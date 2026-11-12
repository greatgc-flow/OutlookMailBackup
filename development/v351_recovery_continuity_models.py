"""v3.5.1 historical-prefix recovery models and source contracts.

Python file-backed counterexamples and bounded-reader models, not VBA/Win32 or
Outlook execution. Historical bytes are the independent SHA-256 oracle; valid
checked rows and equal/larger file size must not erase previously saved proof.
"""
from __future__ import annotations

import hashlib
import io
from pathlib import Path
import re
import tempfile
import unittest

ROOT = Path(__file__).resolve().parent
BUFFER_BYTES = 65536
MAX_EXACT = 2**53 - 1


class Review(Exception):
    pass


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def checked(fields) -> str:
    body = "\t".join(str(value) for value in fields)
    return body + "\t" + sha(body.encode("utf-16-le")) + "\tEND"


def row(name: str, schema="MBS3", body=b"original bytes") -> str:
    key = sha(name.encode())
    source = sha(("MBSOURCE2|5:store|" + str(len(name)) + ":" + name).encode("utf-16-le"))
    relative = "2026\\11\\" + name + ".msg"
    fields = [schema, "N", key, relative if schema == "MBS4" else "C:\\Mail\\" + relative,
              str(len(body)), sha(body), source, "store", name, "Inbox", "2026-11-11T10:00:00",
              "2026-11-10T10:00:00", "Unicode \uD55C\uAE00 \U0001F600 %09", "0", "",
              "REOPENED_CONTENT_SHA256_V3", "C:\\Mail"]
    if schema == "MBS2":
        fields[1], fields[14], fields[15] = "L", "C:\\Mail", "LEGACY_COMPARE_SHA256_V3"
        fields.pop()
    return checked(fields)


def wire(rows) -> bytes:
    text = "\r\n".join(rows)
    return b"\xff\xfe" + (text + ("\r\n" if text else "")).encode("utf-16-le")


def valid_checked(content: bytes) -> bool:
    if len(content) % 2:
        return False
    try:
        for line in content.decode("utf-16").splitlines():
            if not line:
                continue
            fields = line.split("\t")
            if len(fields) < 3 or fields[-1] != "END":
                return False
            if fields[-2] != sha("\t".join(fields[:-2]).encode("utf-16-le")):
                return False
        return True
    except UnicodeError:
        return False


def prior_health(content: bytes, saved: bytes) -> bool:
    """Pre-fix behavior relevant to the counterexample, after presence checks."""
    return valid_checked(content) and len(content) >= len(saved)


def hash_prefix(reader, byte_count: int) -> str:
    if isinstance(byte_count, bool) or not isinstance(byte_count, int) or not 0 <= byte_count <= MAX_EXACT:
        raise Review("Invalid exact prefix byte count")
    remaining = byte_count
    result = hashlib.sha256()
    while remaining:
        wanted = min(remaining, BUFFER_BYTES)
        value = reader.read(wanted)
        if not value or len(value) > wanted:
            raise Review("Snapshot boundary unavailable")
        result.update(value)
        remaining -= len(value)
    return result.hexdigest()


def healthy_file(path: Path, saved: bytes, leaf="journal.tsv") -> bool:
    content = path.read_bytes()
    if not valid_checked(content):
        return False
    append_only = leaf.lower() in {"journal.tsv", "moves.tsv"} or leaf.lower().startswith("baseline_")
    if not append_only:
        return True
    if len(content) < len(saved):
        return False
    with path.open("rb") as stream:
        return hash_prefix(stream, len(saved)) == sha(saved)


class RecordingReader:
    def __init__(self, content, *, max_piece=None, fail_after=None):
        self.source = io.BytesIO(content)
        self.max_piece = max_piece
        self.fail_after = fail_after
        self.requests = []
        self.read_bytes = 0

    def read(self, wanted):
        self.requests.append(wanted)
        if self.fail_after is not None and self.read_bytes >= self.fail_after:
            raise OSError("injected read failure")
        if self.max_piece is not None:
            wanted = min(wanted, self.max_piece)
        value = self.source.read(wanted)
        self.read_bytes += len(value)
        return value


def new_plan(live: bytes | None, newest: bytes | None, *, newest_blocks=True,
             older: bytes | None=None, pending=None):
    """Newest readable manifest remains authoritative once known.

    `older` exists to prove that a healthy older generation does not authorize
    losing history from a later readable manifest. The returned plan holds the
    immutable original live overlay. It does not mutate either input.
    """
    if pending is not None:
        return pending
    if live is not None and newest is not None and prior_health(live, newest):
        if hash_prefix(io.BytesIO(live), len(newest)) == sha(newest):
            return "HEALTHY"
    if newest is not None:
        if not newest_blocks:
            raise Review("Newest readable history has unavailable blocks")
        return dict(base=newest, overlay=live)
    return dict(base=None, overlay=live)


def proof_set(content: bytes) -> dict[str, str]:
    if not valid_checked(content):
        raise Review("Unproven rows")
    result = {}
    for line in content.decode("utf-16").splitlines():
        if not line:
            continue
        fields = line.split("\t")
        old = result.get(fields[2])
        if old is not None and old != fields[5]:
            raise Review("Historical expected digest conflict")
        result[fields[2]] = fields[5]
    return result


def merge_proof(plan):
    result = {}
    for content in (plan["base"], plan["overlay"]):
        if content is None:
            continue
        for key, digest in proof_set(content).items():
            if key in result and result[key] != digest:
                raise Review("Immutable expected proof conflict")
            result[key] = digest
    return result


class PrefixCounterexamples(unittest.TestCase):
    def check_live(self, content, saved, leaf="journal.tsv"):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / leaf
            path.write_bytes(content)
            before = path.read_bytes()
            answer = healthy_file(path, saved, leaf)
            self.assertEqual(path.read_bytes(), before, "health checks must not rewrite evidence")
            return answer

    def test_unchanged_history_passes(self):
        saved = wire([row("alpha")])
        self.assertTrue(self.check_live(saved, saved))

    def test_new_valid_append_passes_without_hashing_tail(self):
        saved = wire([row("alpha")])
        current = saved + wire([row("bravo", "MBS4")])[2:]
        self.assertTrue(self.check_live(current, saved))
        reader = RecordingReader(current)
        self.assertEqual(hash_prefix(reader, len(saved)), sha(saved))
        self.assertEqual(reader.read_bytes, len(saved))

    def test_same_length_valid_row_replacement_is_detected(self):
        saved, current = wire([row("alpha")]), wire([row("bravo")])
        self.assertEqual(len(current), len(saved))
        self.assertTrue(prior_health(current, saved))
        self.assertFalse(self.check_live(current, saved))

    def test_deleted_row_hidden_by_longer_tail_is_detected(self):
        saved = wire([row("alpha"), row("bravo")])
        current = wire([row("bravo"), row("charlie-long"), row("delta-long")])
        self.assertTrue(prior_health(current, saved))
        self.assertFalse(self.check_live(current, saved))

    def test_valid_row_reordering_is_detected(self):
        saved = wire([row("alpha"), row("bravo")])
        current = wire([row("bravo"), row("alpha")])
        self.assertTrue(prior_health(current, saved))
        self.assertFalse(self.check_live(current, saved))

    def test_shorter_history_is_still_rejected(self):
        self.assertFalse(self.check_live(wire([row("alpha")]), wire([row("alpha"), row("bravo")])))

    def test_baseline_receives_the_same_continuity_check(self):
        saved = wire([checked(["MBBASE1", "root", "store", "mail", "path", "37", sha(b"old"), "utc", "MSG_ONLY_REBASELINED"])])
        current = wire([checked(["MBBASE1", "root", "store", "mail", "path", "37", sha(b"new"), "utc", "MSG_ONLY_REBASELINED"])])
        self.assertTrue(prior_health(current, saved))
        self.assertFalse(self.check_live(current, saved, "baseline_account.tsv"))

    def test_moves_receives_the_same_continuity_check(self):
        saved, current = wire([checked(["MBMOVE", "source", "old"])]), wire([checked(["MBMOVE", "source", "new"])])
        self.assertFalse(self.check_live(current, saved, "moves.tsv"))

    def test_compacted_audit_is_not_misclassified_as_append_only(self):
        saved = wire([checked(["AUDIT", "one"]), checked(["AUDIT", "two"])])
        current = wire([checked(["AUDIT", "new"])])
        self.assertTrue(self.check_live(current, saved, "archive-audit.tsv"))

    def test_bom_and_unicode_are_part_of_the_exact_digest(self):
        saved = wire([row("alpha")])
        self.assertNotEqual(sha(saved), sha(saved[2:]))
        self.assertTrue(self.check_live(saved, saved))
        self.assertEqual(hash_prefix(io.BytesIO(saved), 3), sha(saved[:3]))

    def test_zero_length_snapshot_prefix_allows_valid_new_bytes(self):
        self.assertTrue(self.check_live(wire([row("alpha")]), b""))
        reader = RecordingReader(b"tail")
        self.assertEqual(hash_prefix(reader, 0), sha(b""))
        self.assertEqual(reader.requests, [])

    def test_odd_utf16_truncation_does_not_pass(self):
        saved = wire([row("alpha")])
        self.assertFalse(self.check_live(saved[:-1], saved))

    def test_mixed_legacy_and_current_serialization_remains_raw(self):
        saved = wire([row("legacy", "MBS2"), row("alpha", "MBS3")])
        current = saved + wire([row("bravo", "MBS4")])[2:]
        self.assertTrue(self.check_live(current, saved))

    def test_legacy_blank_lines_remain_compatible_with_normal_append(self):
        saved = wire(["", row("legacy", "MBS2"), "", row("alpha"), ""])
        current = saved + wire([row("bravo", "MBS4")])[2:]
        self.assertTrue(self.check_live(current, saved))

    def test_prefix_mutations_at_many_offsets_are_detected(self):
        saved = bytes(range(256)) * 601
        for offset in (0, 1, 65535, 65536, 65537, len(saved) - 1):
            changed = bytearray(saved + b"later tail")
            changed[offset] ^= 1
            with self.subTest(offset=offset):
                self.assertNotEqual(hash_prefix(io.BytesIO(changed), len(saved)), sha(saved))


class BoundedPrefixReader(unittest.TestCase):
    def test_larger_than_one_mebibyte_is_read_in_bounded_requests(self):
        saved = bytes(range(256)) * 4096 + b"odd tail"
        reader = RecordingReader(saved + b"suffix ignored")
        self.assertEqual(hash_prefix(reader, len(saved)), sha(saved))
        self.assertGreater(len(reader.requests), 16)
        self.assertLessEqual(max(reader.requests), BUFFER_BYTES)
        self.assertEqual(reader.read_bytes, len(saved))

    def test_positive_short_reads_continue_to_exact_boundary(self):
        saved = bytes(range(251)) * 881
        reader = RecordingReader(saved + b"tail", max_piece=997)
        self.assertEqual(hash_prefix(reader, len(saved)), sha(saved))
        self.assertEqual(reader.read_bytes, len(saved))

    def test_one_byte_reads_preserve_utf16_surrogates_and_boundaries(self):
        saved = wire([row("alpha")])
        self.assertEqual(hash_prefix(RecordingReader(saved, max_piece=1), len(saved)), sha(saved))

    def test_early_eof_is_not_hash_of_a_shorter_history(self):
        with self.assertRaises(Review):
            hash_prefix(io.BytesIO(b"short"), 20)

    def test_io_error_is_not_converted_to_an_empty_prefix(self):
        with self.assertRaises(OSError):
            hash_prefix(RecordingReader(b"a" * 100000, fail_after=65536), 100000)

    def test_oversized_reader_response_is_rejected(self):
        class BadReader:
            def read(self, count):
                return b"x" * (count + 1)
        with self.assertRaises(Review):
            hash_prefix(BadReader(), 17)

    def test_invalid_or_inexact_byte_counts_are_rejected(self):
        for count in (-1, 0.5, MAX_EXACT + 1, True):
            with self.subTest(count=count), self.assertRaises(Review):
                hash_prefix(io.BytesIO(b""), count)

    def test_large_count_does_not_narrow_to_signed_long(self):
        reader = RecordingReader(b"")
        with self.assertRaises(Review):
            hash_prefix(reader, 2**31 + 17)
        self.assertEqual(reader.requests, [65536])


class PinnedRecoveryHistory(unittest.TestCase):
    def test_same_length_replacement_enters_recovery_and_restores_old_proof(self):
        saved, live = wire([row("alpha")]), wire([row("bravo")])
        plan = new_plan(live, saved)
        result = merge_proof(plan)
        self.assertEqual(result, proof_set(wire([row("alpha"), row("bravo")])))
        self.assertEqual(plan["overlay"], live)

    def test_deleted_row_and_new_tail_preserve_both_generations(self):
        saved = wire([row("alpha"), row("bravo")])
        live = wire([row("bravo"), row("charlie-long"), row("delta-long")])
        result = merge_proof(new_plan(live, saved))
        self.assertTrue(proof_set(saved).items() <= result.items())
        self.assertTrue(proof_set(live).items() <= result.items())

    def test_changed_expected_hash_cannot_replace_old_proof(self):
        saved, live = wire([row("alpha", body=b"old")]), wire([row("alpha", body=b"new")])
        with self.assertRaises(Review):
            merge_proof(new_plan(live, saved))

    def test_missing_latest_blocks_do_not_fall_back_to_valid_older_history(self):
        older = wire([row("alpha")])
        saved = wire([row("alpha"), row("bravo")])
        live = wire([row("alpha"), row("charlie-long"), row("delta-long")])
        with self.assertRaises(Review):
            new_plan(live, saved, newest_blocks=False, older=older)

    def test_missing_latest_blocks_do_not_fall_back_to_msg_only(self):
        with self.assertRaises(Review):
            new_plan(None, wire([row("alpha")]), newest_blocks=False)

    def test_valid_active_prefix_needs_no_block_reconstruction(self):
        saved = wire([row("alpha")])
        live = saved + wire([row("bravo")])[2:]
        self.assertEqual(new_plan(live, saved, newest_blocks=False), "HEALTHY")

    def test_pending_plan_reuses_its_immutable_original_overlay(self):
        saved = wire([row("alpha")])
        original = wire([row("bravo")])
        pending = new_plan(original, saved)
        self.assertIs(new_plan(saved, saved, pending=pending), pending)
        self.assertEqual(merge_proof(pending), proof_set(wire([row("alpha"), row("bravo")])))

    def test_without_any_readable_snapshot_surviving_rows_are_not_discarded(self):
        live = wire([row("alpha")])
        self.assertEqual(merge_proof(new_plan(live, None)), proof_set(live))

    def test_normalized_merge_requires_its_own_physical_checkpoint(self):
        saved = wire([row("alpha"), ""])
        live = wire([row("bravo"), row("charlie")])
        plan = new_plan(live, saved)
        expected = merge_proof(plan)
        merged = wire([row("alpha"), row("bravo"), row("charlie")])
        self.assertEqual(proof_set(merged), expected)
        self.assertNotEqual(hash_prefix(io.BytesIO(merged), len(saved)), sha(saved))
        self.assertEqual(new_plan(merged, merged), "HEALTHY")
        self.assertEqual(plan["base"], saved, "older raw evidence must remain immutable")
        self.assertEqual(plan["overlay"], live)


class SourceContracts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = (ROOT / "mb_recovery.basfrag").read_text(encoding="ascii")
        cls.native = (ROOT / "mb_operating_selftest.basfrag").read_text(encoding="ascii")

    def body(self, name):
        match = re.search(r"(?ms)^Private (?:Function|Sub) " + name + r"\b.*?^End (?:Function|Sub)\s*$", self.source)
        self.assertIsNotNone(match)
        return match.group(0)

    def test_health_gate_compares_exact_prefix_to_manifest_digest(self):
        body = self.body("MBR_StateHealthy")
        check = 'MBR_HashPrefix(path, MB_ByteCount(CStr(meta(0)))) <> CStr(meta(1))'
        self.assertIn(check, body)
        self.assertLess(body.index("MBR_CheckedStateValid"), body.index(check))
        self.assertLess(body.index(check), body.index("MBR_StateHealthy = True"))
        for leaf in ('"journal.tsv"', '"moves.tsv"', '"baseline_"'):
            self.assertIn(leaf, body)

    def test_prefix_hash_is_raw_bounded_read_and_denies_writers(self):
        body = self.body("MBR_HashPrefix")
        for expected in ("remaining As Double", "buffer(0 To 65535) As Byte", "request = 65536",
                         "request = CLng(remaining)", "MBH_FILE_SHARE_READ", "MBH_OPEN_EXISTING",
                         "count <= 0 Or count > request", "remaining = remaining - CDbl(count)",
                         "MBH_BCryptHashData(context.Handle, VarPtr(buffer(0)), count, 0)"):
            self.assertIn(expected, body)
        for forbidden in ("ReadLine", "ReadAll", "LOF(", "Get #", "MB_Unescape", "MB_EncodeFields"):
            self.assertNotIn(forbidden, body)
        self.assertLess(body.index("MBY_NoReparse"), body.index("MBH_CreateFileW"))

    def test_hash_reader_has_success_and_failure_cleanup(self):
        body = self.body("MBR_HashPrefix")
        self.assertIn("If rc = 0 Then Err.Raise", body)
        failure = body.split("Failed:", 1)[1]
        self.assertIn("If handle <> -1 Then rc = MBH_CloseHandle(handle)", failure)
        self.assertIn("MBH_HashClose context", failure)
        self.assertIn("Err.Raise n, source, detail", failure)

    def test_restore_does_not_choose_an_older_verified_generation_after_health_check(self):
        body = self.body("MBR_RestoreState")
        self.assertIn("manifest = MBR_FindSnapshot(recoveryRoot, accountRoot, storeId, files, chunks)", body)
        self.assertNotIn("files, chunks, True)", body)
        self.assertLess(body.index("MBR_StateHealthy(files, reason)"), body.index("If healthy And Len(pending) = 0 Then"))
        self.assertLess(body.index("If Not MBR_BlocksValid(recoveryRoot, files, chunks, reason) Then"),
                        body.index("baseReference = manifest"))
        self.assertLess(body.index("baseReference = manifest"), body.index("MBR_CaptureOverlay"))
        self.assertIn("MB_ERR + 1195", body)

    def test_pending_merge_still_reuses_existing_plan_before_new_capture(self):
        body = self.body("MBR_RestoreState")
        self.assertLess(body.index('If Left$(pending, 6) = "MERGE|" Then'), body.index("MBR_CaptureOverlay"))
        self.assertIn("intent = pending", body)
        self.assertLess(body.index("MBR_ReadOverlay"), body.index('intent, "START"'))
        self.assertLess(body.index("MBR_ApplyOverlay"), body.index('intent, "DONE"'))

    def test_successful_merge_is_checkpointed_before_done_without_dropping_original_plan(self):
        body = self.body("MBR_RestoreState")
        self.assertLess(body.index('intent, "START"'), body.index("MBR_ApplyOverlay"))
        self.assertLess(body.index("MBR_ApplyOverlay"), body.index("MBR_SaveSnapshot accountRoot, storeId"))
        self.assertLess(body.index("MBR_SaveSnapshot accountRoot, storeId"), body.index('intent, "DONE"'))
        checkpoint = body.split("MBR_ApplyOverlay", 1)[1].split('intent, "DONE"', 1)[0]
        self.assertNotIn("On Error Resume Next", checkpoint, "a failed checkpoint must retain START")
        self.assertNotIn("Delete", checkpoint, "old snapshot/overlay are independent immutable evidence")

    def test_native_fixture_exercises_large_prefix_and_valid_row_loss(self):
        for expected in ("MBR_HashPrefix(probePath", "Native prefix continuity", "MBR_StateHealthy(files, reason)",
                         "Historical row loss was not detected", "Lost snapshot proof was not recovered",
                         "Verified merge checkpoint did not prevent repeated historical restoration"):
            self.assertIn(expected, self.native)
        self.assertNotRegex(self.native, r"(?m)^Public (Sub|Function) ")


if __name__ == "__main__":
    unittest.main()
