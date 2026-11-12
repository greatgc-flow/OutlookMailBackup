#!/usr/bin/env python3
"""Reject release metadata or recognized BAS evidence for another artifact."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[1]
BAS_CHECK_PREFIXES = ("Project structural checks", "Python models/source contracts",
                      "ANTLR grammar", "static symbol linkage")


def check_evidence(data: bytes, report: dict) -> list[str]:
    errors = []
    source = data.decode("ascii")
    versions = re.findall(r'^Private Const MB_VERSION As String = "([^"]+)"\s*$', source, re.M)
    if len(versions) != 1:
        return ["Exactly one MB_VERSION constant is required"]
    digest = hashlib.sha256(data).hexdigest()
    expected = dict(name="OutlookMailBackup.bas", sha256=digest, bytes=len(data),
                    lines=len(source.splitlines()), version=versions[0],
                    public_macros=["MB_01_RunBackup"])
    artifact = report.get("artifact") if isinstance(report, dict) else None
    if not isinstance(artifact, dict):
        return ["docs/validation.json must contain an artifact object"]
    for key, value in expected.items():
        if artifact.get(key) != value:
            errors.append(f"Release artifact {key} does not match the shipped BAS")

    def walk(value, location):
        if isinstance(value, dict):
            check = value.get("check", "")
            bas_result = (isinstance(check, str) and check.startswith(BAS_CHECK_PREFIXES))
            if bas_result and value.get("sha256") != digest:
                errors.append(f"Stale/different BAS evidence in {location}.sha256; archive historical results separately")
            for key, child in value.items():
                path = location + "." + key
                if key == "artifact_sha256" and child != digest:
                    errors.append(f"Stale/different BAS evidence in {path}")
                else:
                    walk(child, path)
        elif isinstance(value, list):
            for index, child in enumerate(value):
                walk(child, f"{location}[{index}]")
    walk(report, "validation")
    return errors


def main() -> int:
    try:
        data = (ROOT / "OutlookMailBackup.bas").read_bytes()
        report = json.loads((ROOT / "docs/validation.json").read_text(encoding="utf-8"))
        errors = check_evidence(data, report)
    except (OSError, UnicodeError, ValueError) as exc:
        errors = [str(exc)]
    for error in errors:
        print("FAIL: " + error)
    if not errors:
        print("PASS: current release metadata/evidence refers to the shipped BAS")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
