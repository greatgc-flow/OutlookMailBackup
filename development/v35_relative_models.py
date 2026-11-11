"""Account-relative journal codec, mixed-generation replay and source contracts.

The model specifies Windows lexical path safety and the checked UTF-16LE wire
format independently of Outlook. Tests do not compile VBA, access real MSG
content, or establish that moving an archive to a different root is supported.
"""
from __future__ import annotations

from hashlib import sha256
import ntpath
from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parent
ACCOUNT = r"C:\Archive\Account"
DIGEST = "a" * 64
KEY = "b" * 64
SOURCE = "c" * 64


def utf16_units(text):
    return len(text.encode("utf-16-le", "surrogatepass")) // 2


def valid_component(value):
    if not value or value in (".", "..") or value != value.strip(" "):
        raise ValueError("empty/dot/trimmed component")
    if value[-1] in " ." or utf16_units(value) > 255:
        raise ValueError("trailing punctuation or oversized component")
    if any(ord(c) < 32 or c in '<>:"/\\|?*' for c in value):
        raise ValueError("Windows-forbidden character")
    units = list(value.encode("utf-16-le", "surrogatepass"))
    units = [units[i] | units[i + 1] << 8 for i in range(0, len(units), 2)]
    i = 0
    while i < len(units):
        if 0xD800 <= units[i] <= 0xDBFF:
            if i + 1 >= len(units) or not 0xDC00 <= units[i + 1] <= 0xDFFF:
                raise ValueError("unpaired high surrogate")
            i += 2
        elif 0xDC00 <= units[i] <= 0xDFFF:
            raise ValueError("unpaired low surrogate")
        else:
            i += 1
    stem = value.upper().split(".", 1)[0].rstrip(" .")
    if stem in {"CON", "PRN", "AUX", "NUL", "CLOCK$", "CONIN$", "CONOUT$"}:
        raise ValueError("reserved device")
    if len(stem) == 4 and stem[:3] in ("COM", "LPT") and stem[3] in "123456789¹²³":
        raise ValueError("reserved numbered device")
    return value


def absolute(path):
    path = path.replace("/", "\\")
    if path.startswith(("\\\\?\\", "\\\\.\\")):
        raise ValueError("device namespace")
    drive, tail = ntpath.splitdrive(path)
    if drive.startswith("\\\\"):
        server_share = drive[2:].split("\\")
        if len(server_share) != 2 or not all(server_share):
            raise ValueError("UNC root")
    elif not re.fullmatch(r"[A-Za-z]:", drive) or not tail.startswith("\\"):
        raise ValueError("absolute drive or UNC path required")
    canonical = ntpath.normpath(path)
    if utf16_units(canonical) > 235:
        raise ValueError("path budget")
    body = canonical[2:] if canonical.startswith("\\\\") else canonical[3:]
    for part in body.split("\\") if body else ():
        valid_component(part)
    return canonical


def account(root):
    root = absolute(root)
    _, tail = ntpath.splitdrive(root)
    if tail in ("", "\\"):
        raise ValueError("account cannot be drive/share root")
    return root


def account_path(relative, root):
    root = account(root)
    if not relative or relative.startswith("\\") or "/" in relative or ":" in relative:
        raise ValueError("account-relative path required")
    parts = relative.split("\\")
    for part in parts:
        valid_component(part)
    return absolute(root + "\\" + "\\".join(parts))


def account_relative(path, root):
    root = account(root)
    path = absolute(path)
    prefix = root + "\\"
    if not path.casefold().startswith(prefix.casefold()):
        raise ValueError("outside account boundary")
    relative = path[len(prefix):]
    if account_path(relative, root).casefold() != path.casefold():
        raise ValueError("round-trip failure")
    return relative


ESCAPES = (("%", "%25"), ("\t", "%09"), ("\r", "%0D"), ("\n", "%0A"))
FORMULA = {"=": "%3D", "+": "%2B", "-": "%2D", "@": "%40"}


def escape(value):
    for char, token in ESCAPES:
        value = value.replace(char, token)
    if value and value[0] in FORMULA:
        value = FORMULA[value[0]] + value[1:]
    return value


def unescape(value):
    for char, token in ESCAPES[1:]:
        value = value.replace(token, char)
    for char, token in FORMULA.items():
        value = value.replace(token, char)
    return value.replace("%25", "%")


def checked(fields):
    payload = "\t".join(escape(value) for value in fields)
    return payload + "\t" + sha256(payload.encode("utf-16-le", "surrogatepass")).hexdigest() + "\tEND"


def parse(line):
    raw = line.split("\t")
    if len(raw) < 3 or raw[-1] != "END":
        raise ValueError("incomplete checked row")
    payload = "\t".join(raw[:-2])
    if sha256(payload.encode("utf-16-le", "surrogatepass")).hexdigest() != raw[-2]:
        raise ValueError("bad checksum")
    fields = [unescape(value) for value in raw[:-2]]
    schema = fields[0]
    supported = {"MBS2": (16, {"M", "A", "L"}), "MBS3": (17, {"N", "T", "R"}),
                 "MBS4": (17, {"N", "T", "R", "L"})}
    if schema not in supported:
        raise ValueError("unknown checked schema")
    count, kinds = supported[schema]
    if len(fields) != count or fields[1] not in kinds:
        raise ValueError("unsupported field count/kind")
    return fields


def expand(row):
    if row[0] != "MBS4":
        return row
    if len(row) != 17 or row[1] not in {"N", "T", "R", "L"}:
        raise ValueError("relative record shape")
    root = account(row[16])
    result = list(row)
    result[3] = account_path(row[3], root)
    if row[1] == "L":
        if row[14]:
            raise ValueError("duplicated legacy root")
        result = result[:16]
        result[0], result[14] = "MBS2", root
    else:
        result[0], result[16] = "MBS3", root
    return result


def record(schema="MBS4", kind="N", *, key=KEY, root=ACCOUNT, path=None, reference="", subject="subject"):
    if path is None:
        path = r"2026\11\title.msg" if schema == "MBS4" else root + r"\2026\11\title.msg"
    if schema == "MBS2" and kind in ("M", "A"):
        path = r"messages\ab\old.msg" if path is None or ntpath.isabs(path) else path
    fields = [schema, kind, key, path, "100", DIGEST, SOURCE, "store-A", "entry-A", "Inbox",
              "20261111_010000", "20261110_235959", subject, "0", reference, "VALIDATED_V3"]
    if schema in ("MBS3", "MBS4"):
        fields.append(root)
    elif kind == "L":
        fields[14] = root
    return fields


def index_key(row):
    return ("M" if row[1] in {"N", "R"} else row[1]) + "|" + row[2]


class MixedJournal:
    def __init__(self):
        self.rows = {}
        self.raw = []

    def replay(self, line):
        raw_row = parse(line)
        row = expand(raw_row)
        key = index_key(row)
        proof = row[5], row[4], row[6], row[7]
        if key in self.rows:
            prior = self.rows[key]
            if (prior[5], prior[4], prior[6], prior[7]) != proof:
                raise ValueError("immutable proof conflict")
        if row[1] == "R":
            ref = row[14]
            if ref == key or ref not in self.rows:
                raise ValueError("missing/self alias reference")
            prior = self.rows[ref]
            if (prior[5], prior[4], prior[7]) != (row[5], row[4], row[7]):
                raise ValueError("alias bytes/store conflict")
        self.rows[key] = row
        self.raw.append(line)


class RelativePathCases(unittest.TestCase):
    def test_drive_rooted_account_roundtrip(self):
        relative = r"2026\11\title.msg"
        self.assertEqual(account_relative(account_path(relative, ACCOUNT), ACCOUNT), relative)

    def test_unc_account_roundtrip(self):
        root = r"\\server\share\Account"
        relative = r"2026\11\title.msg"
        self.assertEqual(account_path(relative, root), root + "\\" + relative)
        self.assertEqual(account_relative(root + "\\" + relative, root), relative)

    def test_unicode_combining_emoji_and_percent_are_preserved(self):
        for name in ("한글 中文.msg", "e\u0301 résumé.msg", "😀.msg", "%09_100%.msg", "literal%2e%2e.msg"):
            with self.subTest(name=name):
                self.assertEqual(account_relative(account_path(name, ACCOUNT), ACCOUNT), name)

    def test_sibling_prefix_is_not_inside_account(self):
        with self.assertRaises(ValueError):
            account_relative(ACCOUNT + r"_Other\x.msg", ACCOUNT)

    def test_account_directory_itself_is_not_a_file(self):
        with self.assertRaises(ValueError):
            account_relative(ACCOUNT, ACCOUNT)

    def test_drive_and_share_roots_are_rejected(self):
        for root in ("C:\\", r"\\server\share", "\\\\server\\share\\"):
            with self.subTest(root=root), self.assertRaises(ValueError):
                account_path("x.msg", root)

    def test_nonabsolute_and_device_account_roots_are_rejected(self):
        for root in (r"C:Account", r"Account", r"\Account", r"\\?\C:\Account", r"\\.\C:\Account"):
            with self.subTest(root=root), self.assertRaises(ValueError):
                account_path("x.msg", root)

    def test_relative_escape_and_ambiguous_separators_are_rejected(self):
        for path in ("", ".", "..", r"..\x.msg", r"2026\..\x.msg", r"2026\.\x.msg",
                     r"\x.msg", r"\\host\share\x.msg", r"C:\x.msg", "C:x.msg", "a/b.msg",
                     r"a\\b.msg", "a\\", "file.msg:stream", "a:stream\\x.msg"):
            with self.subTest(path=path), self.assertRaises(ValueError):
                account_path(path, ACCOUNT)

    def test_device_components_are_rejected_even_with_extensions(self):
        for path in ("CON", "con.msg", "AUX.txt", "NUL ", "COM1.msg", "LPT9.msg", "COM¹.msg",
                     "LPT².msg", "CONIN$.msg", r"2026\NUL\x.msg"):
            with self.subTest(path=path), self.assertRaises(ValueError):
                account_path(path, ACCOUNT)

    def test_trailing_dot_space_and_leading_space_are_not_rewritten(self):
        for path in ("a.msg.", "a.msg ", " a.msg", r"2026.\x.msg", r"2026 \x.msg"):
            with self.subTest(path=path), self.assertRaises(ValueError):
                account_path(path, ACCOUNT)

    def test_c0_controls_and_invalid_surrogates_are_rejected(self):
        for char in [chr(n) for n in range(32)] + ["\ud800", "\udfff", "\ud800A"]:
            with self.subTest(char=repr(char)), self.assertRaises(ValueError):
                account_path("file" + char + ".msg", ACCOUNT)

    def test_utf16_path_budget_counts_emoji_as_two_units(self):
        prefix = ACCOUNT + "\\"
        safe = "😀" * 80 + "x" * (235 - utf16_units(prefix) - 160 - 4) + ".msg"
        self.assertEqual(utf16_units(account_path(safe, ACCOUNT)), 235)
        with self.assertRaises(ValueError):
            account_path("x" + safe, ACCOUNT)

    def test_case_insensitive_boundary_preserves_file_spelling(self):
        path = account_path(r"2026\11\MiXeD.msg", ACCOUNT)
        self.assertEqual(account_relative(path, ACCOUNT.lower()), r"2026\11\MiXeD.msg")


class CheckedCodecCases(unittest.TestCase):
    def test_all_new_kinds_have_17_wire_fields(self):
        for kind in ("N", "T", "R", "L"):
            with self.subTest(kind=kind):
                row = record(kind=kind)
                self.assertEqual(parse(checked(row)), row)
                self.assertEqual(len(checked(row).split("\t")), 19)

    def test_relative_file_and_single_root_binding_expand(self):
        row = record()
        expanded = expand(row)
        self.assertEqual(expanded[0], "MBS3")
        self.assertEqual(expanded[3], ACCOUNT + "\\" + row[3])
        self.assertEqual(expanded[16], ACCOUNT)
        self.assertEqual(row[0], "MBS4")
        self.assertEqual(row[3], r"2026\11\title.msg")

    def test_relative_legacy_expands_to_exact_16_field_legacy_contract(self):
        row = record(kind="L")
        expanded = expand(row)
        self.assertEqual(len(expanded), 16)
        self.assertEqual(expanded[0], "MBS2")
        self.assertEqual(expanded[14], ACCOUNT)
        self.assertEqual(row[14], "")
        self.assertEqual(len(row), 17)

    def test_old_mbs2_rows_are_not_reencoded_or_relocated(self):
        for kind in ("M", "A", "L"):
            row = record("MBS2", kind)
            line = checked(row)
            self.assertEqual(expand(parse(line)), row)
            self.assertEqual(checked(expand(parse(line))), line)

    def test_old_mbs3_rows_are_not_reencoded_or_relocated(self):
        for kind in ("N", "T", "R"):
            row = record("MBS3", kind)
            line = checked(row)
            self.assertEqual(expand(parse(line)), row)
            self.assertEqual(checked(expand(parse(line))), line)

    def test_wire_checksum_remains_over_original_relative_payload(self):
        row = record(subject="제목%\tline\r\n=+-@")
        line = checked(row)
        self.assertEqual(parse(line), row)
        expanded = expand(parse(line))
        self.assertNotEqual(checked(expanded), line)
        self.assertEqual(parse(line)[0], "MBS4")

    def test_subject_formula_prefixes_and_literal_escapes_roundtrip(self):
        for subject in ("=A1", "+value", "-value", "@value", "%09", "%2509", "한글\t中文\n😀"):
            with self.subTest(subject=subject):
                self.assertEqual(parse(checked(record(subject=subject)))[12], subject)

    def test_tampered_checked_payload_is_rejected_before_schema(self):
        line = checked(record())
        with self.assertRaises(ValueError):
            parse(line.replace("MBS4", "MBS9"))

    def test_incomplete_terminal_and_bad_field_count_are_rejected(self):
        for line in (checked(record())[:-1], checked(record()[:-1]), checked(record() + ["extra"])):
            with self.subTest(line=line[-30:]), self.assertRaises(ValueError):
                parse(line)

    def test_unknown_valid_schema_or_kind_requires_review(self):
        cases = [record("MBS9", "N"), record("MBS4", "M"), record("MBS4", "A"),
                 record("MBS4", "UNKNOWN"), record("MBS3", "L"), record("MBS2", "R")]
        for row in cases:
            with self.subTest(schema=row[0], kind=row[1]), self.assertRaises(ValueError):
                parse(checked(row))

    def test_duplicate_legacy_root_in_relative_wire_is_rejected(self):
        row = record(kind="L", reference=ACCOUNT)
        with self.assertRaises(ValueError):
            expand(parse(checked(row)))

    def test_absolute_wire_file_path_cannot_escape_relative_semantics(self):
        for path in (ACCOUNT + r"\x.msg", r"\\server\share\x.msg", r"..\x.msg"):
            with self.subTest(path=path), self.assertRaises(ValueError):
                expand(parse(checked(record(path=path))))

    def test_escaped_tab_in_filename_is_rejected_after_wire_decode(self):
        row = record(path="bad\tname.msg")
        self.assertIn("%09", checked(row))
        with self.assertRaises(ValueError):
            expand(parse(checked(row)))

    def test_literal_percent09_in_filename_is_not_double_decoded(self):
        row = record(path="literal%09.msg")
        self.assertIn("%2509", checked(row))
        self.assertTrue(expand(parse(checked(row)))[3].endswith("literal%09.msg"))


class MixedFormatReplayCases(unittest.TestCase):
    def test_relative_r_can_reference_old_and_new_accepted_proofs(self):
        for schema, kind in (("MBS2", "M"), ("MBS2", "L"), ("MBS3", "N"), ("MBS4", "N"), ("MBS4", "L")):
            with self.subTest(schema=schema, kind=kind):
                journal = MixedJournal()
                parent = record(schema, kind)
                journal.replay(checked(parent))
                alias = record(kind="R", key="d" * 64, reference=index_key(parent))
                journal.replay(checked(alias))
                self.assertEqual(journal.rows["M|" + "d" * 64][5], DIGEST)

    def test_same_immutable_proof_survives_absolute_to_relative_format(self):
        journal = MixedJournal()
        journal.replay(checked(record("MBS3", "N")))
        journal.replay(checked(record("MBS4", "N")))
        self.assertEqual(len(journal.rows), 1)
        self.assertEqual(journal.rows["M|" + KEY][5], DIGEST)

    def test_mixed_format_conflict_cannot_replace_historical_hash(self):
        journal = MixedJournal()
        original = checked(record("MBS3", "N"))
        journal.replay(original)
        changed = record(); changed[5] = "e" * 64
        with self.assertRaises(ValueError):
            journal.replay(checked(changed))
        self.assertEqual(journal.raw, [original])
        self.assertEqual(journal.rows["M|" + KEY][5], DIGEST)

    def test_alias_missing_or_self_reference_is_rejected(self):
        for reference in ("M|" + KEY, "L|" + "e" * 64):
            with self.subTest(reference=reference), self.assertRaises(ValueError):
                MixedJournal().replay(checked(record(kind="R", reference=reference)))

    def test_alias_different_hash_size_or_store_is_rejected(self):
        for index, value in ((5, "e" * 64), (4, "101"), (7, "other-store")):
            journal = MixedJournal(); parent = record("MBS3", "N")
            journal.replay(checked(parent))
            alias = record(kind="R", key="d" * 64, reference=index_key(parent))
            alias[index] = value
            with self.subTest(index=index), self.assertRaises(ValueError):
                journal.replay(checked(alias))

    def test_replay_keeps_each_original_checked_line_byte_for_byte(self):
        journal = MixedJournal()
        lines = [checked(record("MBS3", "N")), checked(record(kind="L", key="d" * 64))]
        for line in lines:
            journal.replay(line)
        self.assertEqual(journal.raw, lines)
        self.assertIn("MBS4", journal.raw[1])
        self.assertEqual(journal.rows["L|" + "d" * 64][0], "MBS2")


def procedure(source, name):
    found = re.search(r"(?ms)^(?:Private|Public) (?:Sub|Function) " + re.escape(name) +
                      r"\b.*?^End (?:Sub|Function)\s*$", source)
    if not found:
        raise AssertionError(f"Missing procedure {name}")
    return found.group(0)


class SourceContracts(unittest.TestCase):
    def source(self, filename):
        return (ROOT / filename).read_text(encoding="ascii")

    def test_relative_helpers_are_part_of_final_assembler(self):
        self.assertIn("'mb_relative.basfrag'", self.source("mb_assemble.py"))

    def test_no_hardcoded_default_backup_root_remains(self):
        self.assertNotIn("MB_OUTPUT_ROOT", self.source("mb_main_declarations.basfrag"))
        self.assertNotIn("MB_OUTPUT_ROOT", procedure(self.source("mb_main.basfrag"), "MB_Run"))

    def test_wire_relative_components_use_strict_join_and_bound_check(self):
        body = procedure(self.source("mb_relative.basfrag"), "MBX_AccountPath")
        for required in ('InStr(relative, "/")', 'InStr(relative, ":")', 'MBL_IsFileSystemRoot(value)',
                         'For Each part In Split(relative, "\\")', "MBH_JoinPath(value, CStr(part))",
                         "MBY_CanonicalFilePath(value, accountRoot)"):
            self.assertIn(required, body)

    def test_relative_writer_roundtrips_against_exact_account_root(self):
        body = procedure(self.source("mb_relative.basfrag"), "MBX_AccountRelative")
        self.assertIn("MBY_CanonicalFilePath(path, root)", body)
        self.assertIn("MB_PathWithin(canonical, root)", body)
        self.assertIn("MBX_AccountPath(relative, root)", body)

    def test_expansion_copies_fields_and_leaves_old_schemas_unchanged(self):
        body = procedure(self.source("mb_relative.basfrag"), "MBX_ExpandIndexFields")
        self.assertIn('If CStr(row(0)) <> "MBS4" Then', body)
        self.assertIn("MBX_ExpandIndexFields = row", body)
        self.assertIn("fields(i) = row(i)", body)
        self.assertNotRegex(body, r"(?m)^\s*row\(\d+\)\s*=")
        self.assertIn('fields(0) = "MBS2": fields(14) = root', body)
        self.assertIn('fields(0) = "MBS3": fields(16) = root', body)

    def test_main_writer_uses_relative_new_and_legacy_rows(self):
        body = procedure(self.source("mb_main.basfrag"), "MB_RecordCommit")
        self.assertIn('Array("MBS4", kind, key, MBX_AccountRelative(relativePath, accountRoot)', body)
        self.assertIn('Array("MBS4", kind, key, MBX_AccountRelative(relativePath, originalName)', body)
        self.assertIn('CStr(attNo), "", validation, originalName)', body)
        self.assertIn('Array("MBS2", kind, key, relativePath', body)
        self.assertLess(body.index("MBR_RejectIndexConflict fields"), body.index("MB_OpenJournal"))

    def test_cache_expands_before_path_and_immutable_proof_validation(self):
        body = procedure(self.source("mb_layout_integration.basfrag"), "MBX_CacheRecord")
        self.assertLess(body.index("f = MBX_ExpandIndexFields(f)"), body.index("MBR_RejectIndexConflict f"))
        self.assertLess(body.index("f = MBX_ExpandIndexFields(f)"), body.index("Select Case kind"))

    def test_shared_parser_checks_checksum_before_interpreting_schema(self):
        body = procedure(self.source("mb_recovery_merge.basfrag"), "MBR_ParseIndexFields")
        self.assertLess(body.index("MB_ReadCheckedLine(line, count)"), body.index('CStr(f(0)) = "MBS4"'))
        for schema in ("MBS2", "MBS3", "MBS4"):
            self.assertIn('CStr(f(0)) = "' + schema + '"', body)

    def test_normal_loader_and_salvage_share_checked_parser(self):
        for filename, name in (("mb_main.basfrag", "MB_LoadIndexRow"), ("mb_recovery.basfrag", "MBR_TryIndexRow")):
            body = procedure(self.source(filename), name)
            self.assertIn("f = MBR_ParseIndexFields(line)", body)
            self.assertIn("MBX_CacheRecord f", body)

    def test_restore_keeps_raw_wire_and_pending_cancel_intent(self):
        body = procedure(self.source("mb_recovery_merge.basfrag"), "MBR_MergeJournal")
        self.assertIn("outputStream.WriteLine line", body)
        self.assertIn("If Not MB_ContinueWork() Then Err.Raise", body)
        restore = procedure(self.source("mb_recovery.basfrag"), "MBR_RestoreState")
        self.assertLess(restore.index('intent, "START"'), restore.index("MBR_ApplyOverlay"))
        self.assertLess(restore.index("MBR_ApplyOverlay"), restore.index('intent, "DONE"'))

    def test_store_preflight_reads_mbs4_binding(self):
        body = procedure(self.source("mb_store_binding.basfrag"), "MBS_CheckStoreRow")
        self.assertIn('Left$(line, 5) = "MBS4" & vbTab', body)
        self.assertIn("fieldCount = 17", body)
        self.assertIn("foundStore = CStr(fields(7))", body)

    def test_actual_saveas_still_uses_constructed_absolute_stage(self):
        body = procedure(self.source("mb_main.basfrag"), "MB_ProcessOne")
        self.assertIn("accountRoot = MBY_AccountRoot(CStr(q(1)))", body)
        self.assertIn("messageDir = MBY_MonthDirectory(accountRoot, MBY_MessageDate(mail))", body)
        self.assertIn("stagePath = MBH_NewStagePath(finalPath)", body)
        self.assertIn("mail.SaveAs stagePath, olMSGUnicode", body)
        self.assertNotIn("mail.SaveAs relativePath", body)

    def test_schema_owner_blocks_old_writers_but_accepts_old_readers_input(self):
        declarations = self.source("mb_main_declarations.basfrag")
        self.assertIn('MB_OWNER As String = "OutlookMailBackup.Schema4"', declarations)
        owner = procedure(self.source("mb_main.basfrag"), "MB_CheckOwnership")
        for marker in ("MB_OWNER_V34", "MB_OWNER_PREVIOUS", "MB_OWNER_LEGACY"):
            self.assertIn(marker, owner)
        self.assertIn('"owner-previous-"', owner)
        self.assertIn("MBS_WriteAtomic p, MB_OWNER", owner)

    def test_native_codec_check_is_guarded_and_required_for_acceptance(self):
        body = procedure(self.source("mb_acceptance.basfrag"), "MBT_RunAcceptanceReport")
        call = body.index('relativeOK = MBX_RelativeSelfTest(MBH_JoinPath(runFolder, "account"), testDetail)')
        guard = body.rfind("If helperOK And Not gCancel Then", 0, call)
        self.assertGreaterEqual(guard, 0)
        self.assertIn('"RELATIVE_JOURNAL_CODEC", "NOT_RUN"', body)
        result = next(line for line in body.splitlines() if "MBT_RunAcceptanceReport = envOK" in line)
        self.assertIn("And relativeOK And", result)

    def test_native_codec_check_is_pure_and_preserves_serialized_input(self):
        body = procedure(self.source("mb_relative.basfrag"), "MBX_RelativeSelfTest")
        for forbidden in ("CreateTextFile", "OpenTextFile", "CreateFolder", "GetItemFromID", "OpenSharedItem", "SaveAs", "DeleteFile"):
            self.assertNotIn(forbidden, body)
        self.assertIn("MB_EncodeFields(row) <> payload", body)
        self.assertIn('Array("N", "T", "R", "L")', body)
        self.assertIn("UBound(expanded) <> 15", body)
        self.assertIn("UBound(expanded) <> 16", body)
        self.assertIn("MBX_RelativeRejected(CStr(unsafe), root)", body)


if __name__ == "__main__":
    unittest.main(verbosity=2)
