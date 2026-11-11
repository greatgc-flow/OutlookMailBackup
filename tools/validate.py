#!/usr/bin/env python3
"""Portable structural checks for the shipped BAS, NOT a VBA compiler.

Only the project's Boolean platform #If language is accepted. Identifier checks
cover MB*_ names and the project's module-global naming conventions, not every
VBA/Outlook identifier, type, call signature, control-flow path or COM behavior.
"""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import sys

ROOT = Path(__file__).resolve().parents[1]
PLATFORMS = {
    "vba7_32": dict(vba7=True, win64=False, win32=True, mac=False),
    "vba7_64": dict(vba7=True, win64=True, win32=True, mac=False),
    "legacy32": dict(vba7=False, win64=False, win32=True, mac=False),
}


def condition(expression: str, names: dict[str, bool]) -> bool:
    """Parse a deliberately limited Boolean grammar without eval or short cuts."""
    tokens = re.findall(r"[A-Za-z_]\w*|[01]|[()]", expression)
    if "".join(tokens).lower() != re.sub(r"\s+", "", expression).lower():
        raise ValueError(f"Unsupported conditional expression: {expression}")
    position = 0

    def atom() -> int:
        nonlocal position
        if position == len(tokens):
            raise ValueError("Incomplete conditional expression")
        token = tokens[position].lower()
        position += 1
        if token == "not":
            return ~atom()
        if token == "(":
            result = disjunction()
            if position == len(tokens) or tokens[position] != ")":
                raise ValueError("Unclosed conditional parenthesis")
            position += 1
            return result
        if token in names:
            return -1 if names[token] else 0
        if token == "true":
            return -1
        if token == "false":
            return 0
        if token in {"0", "1"}:
            return int(token)
        raise ValueError(f"Unknown conditional symbol: {token}")

    def conjunction() -> int:
        nonlocal position
        value = atom()
        while position < len(tokens) and tokens[position].lower() == "and":
            position += 1
            right = atom()
            value = value & right
        return value

    def disjunction() -> int:
        nonlocal position
        value = conjunction()
        while position < len(tokens) and tokens[position].lower() == "or":
            position += 1
            right = conjunction()
            value = value | right
        return value

    result = disjunction()
    if position != len(tokens):
        raise ValueError("Unexpected conditional tokens")
    return bool(result)


def preprocessor(source: str, names: dict[str, bool]) -> str:
    output, stack = [], []
    active = True
    for number, line in enumerate(source.splitlines(), 1):
        if not line.lstrip().startswith("#"):
            output.append(line if active else "")
            continue
        text = line.split("'", 1)[0].strip()
        branch = re.fullmatch(r"#(If|ElseIf)\s+(.+?)\s+Then", text, re.I)
        if branch:
            yes = condition(branch.group(2), names)
            if branch.group(1).lower() == "if":
                stack.append(dict(parent=active, any=yes, otherwise=False))
                active = active and yes
            else:
                if not stack or stack[-1]["otherwise"]:
                    raise ValueError(f"Unexpected #ElseIf at line {number}")
                top = stack[-1]
                active = top["parent"] and not top["any"] and yes
                top["any"] = top["any"] or yes
        elif re.fullmatch(r"#Else", text, re.I):
            if not stack or stack[-1]["otherwise"]:
                raise ValueError(f"Unexpected #Else at line {number}")
            top = stack[-1]
            active = top["parent"] and not top["any"]
            top["otherwise"] = True
        elif re.fullmatch(r"#End\s+If", text, re.I):
            if not stack:
                raise ValueError(f"Unexpected #End If at line {number}")
            active = stack.pop()["parent"]
        else:
            raise ValueError(f"Unsupported conditional directive at line {number}: {text}")
        output.append("")
    if stack:
        raise ValueError("Unclosed conditional directive")
    return "\n".join(output)


def code_only(line: str) -> str:
    """Remove string contents/comments so their fake names cannot satisfy linkage."""
    result, quoted, i = [], False, 0
    while i < len(line):
        char = line[i]
        if char == '"':
            if quoted and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            quoted = not quoted
            result.append(" ")
        elif char == "'" and not quoted:
            break
        elif not quoted:
            result.append(char)
        i += 1
    if quoted:
        raise ValueError("Unclosed string literal")
    return re.split(r"(?:^|:)\s*Rem\b", "".join(result), maxsplit=1, flags=re.I)[0]


def inspect_platform(source: str, names: dict[str, bool]) -> dict:
    errors, warnings, declared, procedures, exposed = [], [], [], [], []
    try:
        selected = preprocessor(source, names)
        selected_lines = selected.splitlines()
        lines = [code_only(line) for line in selected_lines]
    except ValueError as exc:
        return {"errors": [str(exc)], "warnings": []}
    text = "\n".join(lines)
    declared.extend(re.findall(r"^\s*(?:Private|Public)\s+(?:Const|Type|Enum)\s+(\w+)", text, re.M | re.I))
    declared.extend(re.findall(r"^\s*(?:Private|Public)\s+Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+(\w+)", text, re.M | re.I))
    declared.extend(re.findall(r"^\s*(?:Private|Public)\s+(\w+)(?:\([^)]*\))?\s+As\b", text, re.M | re.I))
    current = None
    largest = {"name": None, "source_bytes": 0, "physical_lines": 0}
    seen_procedure = False
    continuation = maximum = 0
    for number, line in enumerate(lines, 1):
        if re.search(r"\s_\s*$", line):
            continuation += 1
            maximum = max(maximum, continuation)
            if continuation > 24:
                errors.append(f"More than 24 line continuations at line {number}")
        else:
            continuation = 0
        start = re.match(r"^\s*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function)\s+(\w+)\b(.*)", line, re.I)
        finish = re.fullmatch(r"\s*End\s+(Sub|Function)\s*", line, re.I)
        if start:
            if current:
                errors.append(f"Nested/unclosed procedure at line {number}")
            visibility, kind, name, signature = start.groups()
            procedures.append(name)
            declared.append(name)
            if not visibility or visibility.lower() != "private":
                exposed.append(name)
                if (name != "MB_01_RunBackup" or kind.lower() != "sub"
                        or not re.fullmatch(r"\s*\(\s*\)\s*", signature)):
                    errors.append(f"Unexpected public/default-visible procedure: {name}")
            current = (name, kind.lower(), number)
            seen_procedure = True
        elif finish:
            if not current or current[1] != finish.group(1).lower():
                errors.append(f"Mismatched End {finish.group(1)} at line {number}")
            if current:
                name, _, first = current
                size = len("\n".join(selected_lines[first - 1:number]).encode("ascii"))
                if size > largest["source_bytes"]:
                    largest = dict(name=name, source_bytes=size, physical_lines=number - first + 1)
            current = None
        elif seen_procedure and re.match(r"^\s*(Private|Public|Global)\s+", line, re.I):
            errors.append(f"Module declaration after procedures at line {number}")
        if re.fullmatch(r"\s*End\s*", line, re.I):
            errors.append(f"Unsafe standalone End at line {number}")
    if current:
        errors.append(f"Unclosed procedure: {current[0]}")
    if exposed != ["MB_01_RunBackup"]:
        errors.append("Exactly one public entry point MB_01_RunBackup is required")
    if largest["source_bytes"] >= 32768:
        warnings.append("Large procedure source; compiled 64 KiB procedure limit requires the Office compiler")
    known = {name.lower() for name in declared}
    duplicates = sorted(name for name, count in Counter(name.lower() for name in declared).items() if count > 1)
    errors.extend(f"Duplicate module symbol: {name}" for name in duplicates)
    used = set(re.findall(r"\bMB[A-Z]*_\w+\b", text, re.I))
    globals_used = set(re.findall(r"\b(?:g[A-Z]\w*|mb[a-z][A-Z]\w*)\b", text))
    errors.extend(f"Undefined project symbol: {name}" for name in sorted(used | globals_used) if name.lower() not in known)
    return dict(errors=errors, warnings=warnings, procedures=len(procedures), public_entries=exposed,
                max_line_continuations=maximum, largest_procedure=largest)


def inspect_bytes(data: bytes) -> dict:
    errors = []
    result = dict(check="Project structural checks; NOT Outlook VBA compilation or runtime",
                  sha256=hashlib.sha256(data).hexdigest(), bytes=len(data), errors=errors,
                  platforms={})
    try:
        source = data.decode("ascii")
    except UnicodeDecodeError:
        errors.append("BAS must be ASCII without a BOM")
        return result
    if not data or not data.endswith(b"\r\n") or re.search(rb"(?<!\r)\n|\r(?!\n)", data):
        errors.append("BAS must use CRLF, including its final line ending")
    if not re.match(r'^Attribute VB_Name = "OutlookMailBackup"\r\n', source):
        errors.append("Missing or incorrect importable module attribute")
    if len(re.findall(r"^Option Explicit\s*$", source, re.M | re.I)) != 1:
        errors.append("Exactly one Option Explicit is required")
    if re.search(r"^\s*Option Private Module\s*$", source, re.M | re.I):
        errors.append("Option Private Module is outside this module's public-entry-point policy")
    result["physical_lines"] = len(source.splitlines())
    result["max_physical_line"] = max(map(len, source.splitlines()), default=0)
    if result["max_physical_line"] > 1023:
        errors.append("A physical line exceeds 1023 characters")
    for platform, names in PLATFORMS.items():
        result["platforms"][platform] = inspect_platform(source, names)
    return result


def successful(result: dict) -> bool:
    return not result["errors"] and bool(result["platforms"]) and all(
        not platform["errors"] for platform in result["platforms"].values())


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("path", nargs="?", type=Path, default=ROOT / "OutlookMailBackup.bas")
    parser.add_argument("--json", action="store_true", help="emit the complete machine-readable result")
    args = parser.parse_args()
    result = inspect_bytes(args.path.read_bytes())
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        for error in result["errors"]:
            print("FAIL: " + error)
        for platform, check in result["platforms"].items():
            print(f"{'FAIL' if check['errors'] else 'PASS'}: {platform}: "
                  f"{check.get('procedures', 0)} procedures")
            for error in check["errors"]:
                print("  " + error)
            for warning in check["warnings"]:
                print("  WARNING: " + warning)
        print("Outlook VBA compiler/runtime: NOT_RUN")
    return 0 if successful(result) else 1


if __name__ == "__main__":
    raise SystemExit(main())
