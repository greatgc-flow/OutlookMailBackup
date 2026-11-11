"""Review guidance classification, report lifecycle and VBA source contracts.

These tests exercise an independent event/report model, not Outlook or VBA.
The report is a derived view of observed events and never a work/repair ledger.
Use --integration after the parent has connected the shared main module hooks.
"""
from __future__ import annotations

from collections import Counter
from pathlib import Path
import re
import sys
import unittest


ROOT = Path(__file__).resolve().parent
INTEGRATION = "--integration" in sys.argv
if INTEGRATION:
    sys.argv.remove("--integration")

HEADER = "time\tlevel\tstage\tkey\tpath\terror\tdetail"


def escape(value):
    for char, replacement in (("%", "%25"), ("\t", "%09"), ("\r", "%0D"), ("\n", "%0A")):
        value = value.replace(char, replacement)
    if value[:1] in {"=", "+", "-", "@"}:
        value = {"=": "%3D", "+": "%2B", "-": "%2D", "@": "%40"}[value[0]] + value[1:]
    return value


def unescape(value):
    for token, char in (("%09", "\t"), ("%0D", "\r"), ("%0A", "\n"),
                        ("%3D", "="), ("%2B", "+"), ("%2D", "-"), ("%40", "@"), ("%25", "%")):
        value = value.replace(token, char)
    return value


def wire(values):
    return "\t".join(escape(str(v)) for v in values)


def event(level="WARNING", stage="LEGACY_REVIEW_REQUIRED", key="source-a", path="Account", reason="Review", error="0"):
    return wire(("2026-11-11T12:00:00", level, stage, key, path, error, reason))


def classify(level, stage, reason):
    if level not in {"WARNING", "WARN", "ERROR"}:
        return None
    text = reason.casefold()
    has = lambda part: part.casefold() in text
    stages = {
        "SPECIAL_ATTACHMENT_CONTENT_REVIEW": "CONTENT_LIMITED",
        "SOURCE_CHANGED_RETRY": "SOURCE_CHANGED", "TYPE_CHANGED": "SOURCE_CHANGED",
        "RECOVERY_NEW_BASELINE": "CURRENT_BASELINE", "SCAN_TABLE_FALLBACK": "SCAN_FALLBACK",
    }
    if stage in stages:
        return stages[stage]
    if stage == "DOWNLOAD_STATE" and has("Header-only message"):
        return "DOWNLOAD_REQUIRED"
    if has("No verified old-file match") and has("latest write"):
        return "HISTORY_CUTOFF"
    if stage == "LEGACY_LOOKUP_UNRESOLVED" or has("no readable filename/SentOn lookup"):
        return "LEGACY_UNREADABLE"
    if has("candidate limit reached") or has("bucket exceeds 512 files") or has("Legacy candidate budget exceeded"):
        return "CANDIDATE_LIMIT"
    if any(has(x) for x in ("Indexed file is missing", "Referenced file is unavailable",
                            "previously linked old MSG is unavailable")):
        return "FILE_MISSING"
    if any(has(x) for x in ("SHA-256 mismatch", "changed SHA-256", "SHA-256 differs from",
                            "Indexed file size mismatch", "Archived file size differs", "old MSG has changed size",
                            "Recovered MSG differs from its recorded baseline")):
        return "FILE_CHANGED"
    if has("provider revision already has a file proof"):
        return "REVISION_CONFLICT"
    if any(has(x) for x in ("conflicting", "different file proof",
                            "Multiple content-equal MSG candidates have different file bytes",
                            "Multiple matching legacy files have different bytes")):
        return "PROOF_CONFLICT"
    if any(has(x) for x in ("Custom UserProperties are present", "Embedded/OLE/reference attachment",
                            "special attachment or custom metadata equivalence")):
        return "CONTENT_LIMITED"
    if any(has(x) for x in ("restored identity is uncertain", "provider/import rewriting requires review")):
        return "IDENTITY_UNCERTAIN"
    if stage in {"INVALID_INDEX_ROW", "SOURCE_AUDIT_STATE_REVIEW", "STAGE_LEDGER_REVIEW", "SEMANTIC_INDEX_RECOVERY"}:
        return "STATE_REVIEW"
    if stage == "LEGACY_REVIEW_REQUIRED":
        return "MATCH_UNPROVEN"
    return "OPERATION_FAILED" if level == "ERROR" else "REVIEW_REQUIRED"


def derive(lines):
    lines = iter(lines)
    if next(lines, None) != HEADER:
        raise ValueError("bad header")
    rows, counts = [], Counter()
    context = None
    for number, line in enumerate(lines, 2):
        fields = [unescape(x) for x in line.split("\t")]
        if len(fields) != 7 or fields[1] not in {"INFO", "WARNING", "WARN", "ERROR"}:
            raise ValueError("incomplete/unknown event")
        when, level, stage, key, path, error, reason = fields
        if level == "INFO" and stage == "REVIEW_SUBJECT":
            context = key, reason
            continue
        subject = context[1] if context and context[0] and context[0] == key else ""
        context = None
        code = classify(level, stage, reason)
        if code:
            rows.append(dict(event_row=number, time=when, code=code, subject=subject,
                             location=path, source_key=key, reason=reason, stage=stage, error=error))
            counts[code] += 1
    return rows, counts


class PublishModel:
    """No overwrite; reporting failure cannot manufacture a successful result."""
    def __init__(self, existing=()):
        self.files = {name: "prior bytes" for name in existing}
        self.state = "NOT_CREATED"
        self.count = 0
        self.counts = {}
        self.trace = []

    def publish(self, events, fail_at=None, cancelled=False):
        def step(name):
            self.trace.append(name)
            if name == fail_at:
                raise OSError(name)
        if self.state == "COMPLETE":
            return True
        if self.state == "FAILED":
            return False
        try:
            if self.files:
                raise OSError("existing destination or partial")
            step("open_input")
            if not events or events[0] != HEADER:
                raise ValueError("bad header")
            step("create_partial")
            self.files["partial"] = "header"
            step("write_rows")
            rows, counts = derive(events)
            self.files["partial"] = repr(rows) + "\nEND_OF_REPORT"
            step("close_input")
            step("close_output")
            step("flush_partial")
            step("rename_no_replace")
            self.files["final"] = self.files.pop("partial")
            step("flush_final")
            self.count, self.counts = len(rows), counts
            self.state = "COMPLETE"
            return True
        except (ValueError, OSError):
            self.state = "FAILED"
            return False


class Classification(unittest.TestCase):
    def test_info_is_not_an_outstanding_review(self):
        self.assertIsNone(classify("INFO", "LEGACY_UNMATCHED_NO_EXPORT", "No verified old-file match; latest write"))

    def test_legacy_cutoff_is_actionable_without_authorizing_export(self):
        self.assertEqual(classify("WARNING", "LEGACY_REVIEW_REQUIRED",
            "No verified old-file match; this message predates the old archive's latest write. Automatic re-export is blocked."), "HISTORY_CUTOFF")

    def test_cutoff_requires_both_producer_phrases(self):
        for reason in ("The latest write failed", "No verified old-file match"):
            self.assertEqual(classify("WARNING", "LEGACY_REVIEW_REQUIRED", reason), "MATCH_UNPROVEN")

    def test_unreadable_archive_paths(self):
        for stage, reason in (("LEGACY_LOOKUP_UNRESOLVED", "Cannot read SentOn"),
                              ("LEGACY_REVIEW_REQUIRED", "Some old MSGs have no readable filename/SentOn lookup")):
            self.assertEqual(classify("WARNING", stage, reason), "LEGACY_UNREADABLE")

    def test_identity_and_both_legacy_candidate_caps(self):
        for reason in ("Plausible identity candidate limit reached; unvisited candidates remain",
                       "Legacy candidate bucket exceeds 512 files; all are preserved for review.",
                       "Legacy candidate budget exceeded; unvisited files remain."):
            self.assertEqual(classify("WARNING", "LEGACY_REVIEW_REQUIRED", reason), "CANDIDATE_LIMIT")

    def test_missing_file_wrapped_reason(self):
        self.assertEqual(classify("WARNING", "LEGACY_REVIEW_REQUIRED",
            "Recorded backup needs review: Referenced file is unavailable: X"), "FILE_MISSING")

    def test_hash_and_size_differences(self):
        for reason in ("Indexed SHA-256 mismatch.", "Archived SHA-256 differs from its recorded baseline.",
                       "Indexed file size mismatch.", "A previously linked old MSG has changed size.",
                       "Recovered MSG differs from its recorded baseline."):
            self.assertEqual(classify("WARNING", "ARCHIVE_FILE_REVIEW", reason), "FILE_CHANGED")

    def test_proof_conflict(self):
        self.assertEqual(classify("ERROR", "INVALID_INDEX_ROW", "A key has conflicting historical expected bytes"), "PROOF_CONFLICT")

    def test_same_revision_content_conflict(self):
        self.assertEqual(classify("WARNING", "LEGACY_REVIEW_REQUIRED",
            "Current content differs but its provider revision already has a file proof"), "REVISION_CONFLICT")

    def test_limited_validation_and_unsupported_properties(self):
        self.assertEqual(classify("WARNING", "SPECIAL_ATTACHMENT_CONTENT_REVIEW", "Saved and reopened"), "CONTENT_LIMITED")
        self.assertEqual(classify("WARNING", "LEGACY_REVIEW_REQUIRED", "Custom UserProperties are present"), "CONTENT_LIMITED")

    def test_identity_time_ambiguity(self):
        self.assertEqual(classify("WARNING", "LEGACY_REVIEW_REQUIRED", "Same Message-ID but different times; provider/import rewriting requires review"), "IDENTITY_UNCERTAIN")

    def test_source_changed_and_header_only(self):
        self.assertEqual(classify("WARNING", "SOURCE_CHANGED_RETRY", "Moved"), "SOURCE_CHANGED")
        self.assertEqual(classify("ERROR", "DOWNLOAD_STATE", "Header-only message. Download it fully"), "DOWNLOAD_REQUIRED")

    def test_warn_is_supported_for_existing_scan_producers(self):
        self.assertEqual(classify("WARN", "SCAN_TABLE_FALLBACK", "Use item enumeration"), "SCAN_FALLBACK")
        self.assertEqual(classify("WARN", "ROW_CHANGED", "Modified during enumeration"), "REVIEW_REQUIRED")

    def test_state_and_baseline_have_distinct_guidance(self):
        self.assertEqual(classify("ERROR", "INVALID_INDEX_ROW", "Invalid key shape"), "STATE_REVIEW")
        self.assertEqual(classify("WARNING", "RECOVERY_NEW_BASELINE", "No historical comparison"), "CURRENT_BASELINE")

    def test_unknown_warnings_and_errors_are_retained(self):
        self.assertEqual(classify("WARNING", "FUTURE_REVIEW_STAGE", "New reason"), "REVIEW_REQUIRED")
        self.assertEqual(classify("ERROR", "FUTURE_STAGE", "New reason"), "OPERATION_FAILED")

    def test_reason_matching_is_case_insensitive(self):
        self.assertEqual(classify("WARNING", "LEGACY_REVIEW_REQUIRED", "INDEXED SHA-256 MISMATCH"), "FILE_CHANGED")


class EventDerivation(unittest.TestCase):
    def test_subject_context_is_adjacent_and_same_key(self):
        rows, _ = derive([HEADER, event("INFO", "REVIEW_SUBJECT", reason="A subject"), event(reason="Unclear")])
        self.assertEqual(rows[0]["subject"], "A subject")
        self.assertEqual(rows[0]["event_row"], 3)

    def test_subject_does_not_supply_a_reason_code(self):
        rows, _ = derive([HEADER, event("INFO", "REVIEW_SUBJECT", reason="Indexed SHA-256 mismatch"), event(reason="Unclear")])
        self.assertEqual(rows[0]["code"], "MATCH_UNPROVEN")

    def test_path_does_not_supply_a_reason_code(self):
        rows, _ = derive([HEADER, event(path="Indexed SHA-256 mismatch", reason="Unclear")])
        self.assertEqual(rows[0]["code"], "MATCH_UNPROVEN")

    def test_subject_is_not_applied_to_another_source(self):
        rows, _ = derive([HEADER, event("INFO", "REVIEW_SUBJECT", key="a", reason="Private A"),
                          event(key="b"), event(key="a")])
        self.assertEqual([row["subject"] for row in rows], ["", ""])

    def test_context_expires_on_intervening_info(self):
        rows, _ = derive([HEADER, event("INFO", "REVIEW_SUBJECT", reason="Subject"),
                          event("INFO", "SOMETHING_ELSE"), event()])
        self.assertEqual(rows[0]["subject"], "")

    def test_control_characters_percent_unicode_are_round_tripped(self):
        subject = "=한글\temoji 😀\r\n%09"
        path = "@account\t%0A"
        reason = "Unclear\nline\r\t%09"
        rows, _ = derive([HEADER, event("INFO", "REVIEW_SUBJECT", reason=subject),
                          event(path=path, reason=reason)])
        self.assertEqual((rows[0]["subject"], rows[0]["location"], rows[0]["reason"]), (subject, path, reason))
        emitted = wire((rows[0]["subject"], rows[0]["location"], rows[0]["reason"]))
        self.assertEqual(len(emitted.split("\t")), 3)
        self.assertFalse(emitted.startswith("="))
        self.assertNotIn("\n", emitted)

    def test_counts_are_events_not_unique_mail(self):
        rows, counts = derive([HEADER, event(), event(), event("ERROR", "SAVE_MSG")])
        self.assertEqual(len(rows), 3)
        self.assertEqual(sum(counts.values()), 3)
        self.assertEqual(counts["MATCH_UNPROVEN"], 2)

    def test_empty_event_set_is_a_valid_empty_report(self):
        self.assertEqual(derive([HEADER]), ([], Counter()))

    def test_malformed_header_or_row_is_not_a_partial_success(self):
        for lines in ([], ["wrong"], [HEADER, ""], [HEADER, "a\tb"], [HEADER, event("UNKNOWN")]):
            with self.assertRaises(ValueError):
                derive(lines)

    def test_registration_info_does_not_count_as_unresolved_mail(self):
        rows, _ = derive([HEADER, event("INFO", "LEGACY_UNMATCHED_NO_EXPORT"), event("INFO", "END")])
        self.assertEqual(rows, [])


class Publication(unittest.TestCase):
    def test_success_is_only_after_final_flush(self):
        model = PublishModel()
        self.assertTrue(model.publish([HEADER, event()]))
        self.assertEqual(model.state, "COMPLETE")
        self.assertEqual(model.trace[-1], "flush_final")
        self.assertEqual(model.count, 1)

    def test_every_io_failure_withholds_success_counts(self):
        for stage in ("open_input", "create_partial", "write_rows", "close_input", "close_output",
                      "flush_partial", "rename_no_replace", "flush_final"):
            with self.subTest(stage=stage):
                model = PublishModel()
                self.assertFalse(model.publish([HEADER, event()], fail_at=stage))
                self.assertEqual((model.state, model.count, model.counts), ("FAILED", 0, {}))

    def test_existing_report_or_partial_is_never_overwritten(self):
        for existing in ("partial", "final"):
            model = PublishModel([existing])
            self.assertFalse(model.publish([HEADER, event()]))
            self.assertEqual(model.files, {existing: "prior bytes"})

    def test_failed_partial_is_retained(self):
        model = PublishModel()
        model.publish([HEADER, event()], fail_at="flush_partial")
        self.assertIn("partial", model.files)
        self.assertFalse(model.publish([HEADER]))

    def test_cancelled_run_still_gets_an_event_report(self):
        model = PublishModel()
        self.assertTrue(model.publish([HEADER, event("WARNING", "SOURCE_CHANGED_RETRY")], cancelled=True))
        self.assertEqual(model.counts["SOURCE_CHANGED"], 1)

    def test_repeated_success_does_not_duplicate_report(self):
        model = PublishModel()
        model.publish([HEADER, event()])
        before = dict(model.files)
        self.assertTrue(model.publish([HEADER, event(), event()]))
        self.assertEqual(model.files, before)
        self.assertEqual(model.count, 1)


def procedure(source, name):
    found = re.search(r"(?ims)^Private (?:Sub|Function) " + re.escape(name) +
                      r"\b.*?^End (?:Sub|Function)\s*$", source)
    if not found:
        raise AssertionError("Procedure missing: " + name)
    return found.group(0)


class SourceContracts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = (ROOT / "mb_review.basfrag").read_text(encoding="ascii")

    def test_all_returned_codes_have_actions(self):
        codes = set(re.findall(r'MBG_Code = "([A-Z_]+)"', self.source))
        action = procedure(self.source, "MBG_Action")
        for code in codes - {"REVIEW_REQUIRED"}:
            self.assertIn('"' + code + '"', action)
        self.assertIn("Case Else", action)

    def test_report_publish_orders_close_flush_and_no_replace(self):
        body = procedure(self.source, "MBG_CreateReviewReport")
        operations = ["writer.Close: Set writer = Nothing", "MBM_FlushPath partial",
                      "MBM_MoveFileExW(StrPtr(partial), StrPtr(target), 8)", "MBM_FlushPath target",
                      '\n    gReviewGuideState = "COMPLETE"']
        positions = [body.index(op) for op in operations]
        self.assertEqual(positions, sorted(positions))
        self.assertIn("CreateTextFile(partial, False, True)", body)

    def test_output_is_encoded_and_not_deleted_on_error(self):
        body = procedure(self.source, "MBG_CreateReviewReport")
        self.assertIn("writer.WriteLine MB_EncodeFields(Array", body)
        self.assertNotRegex(body, r"(?i)\b(?:Kill|DeleteFile|DeleteFolder)\b")
        self.assertIn('gReviewGuideState = "FAILED"', body)

    def test_guide_cannot_change_backup_or_coverage(self):
        body = procedure(self.source, "MBG_CreateReviewReport")
        forbidden = ("MB_RecordCommit", "MBW_Commit", "MBR_CaptureBaseline", "MB_NewExportAllowed",
                     ".SaveAs", "GetItemFromID", "gMailReview =", "gWindowCommitted =", "MB_ContinueWork")
        for symbol in forbidden:
            self.assertNotIn(symbol, body)

    def test_subject_context_does_not_modify_reason(self):
        body = procedure(self.source, "MBG_LogReview")
        self.assertIn('"REVIEW_SUBJECT", key, path, 0, subject', body)
        self.assertIn('"WARNING", stage, key, path, errorNumber, reason', body)
        classify_body = procedure(self.source, "MBG_Code")
        self.assertNotRegex(classify_body, r"(?im)^\s*[^']*\bsubject\b")

    def test_actual_producer_phrases_remain_recognized(self):
        reuse = (ROOT / "mb_reuse.basfrag").read_text(encoding="ascii")
        identity = (ROOT / "mb_identity.basfrag").read_text(encoding="ascii")
        self.assertIn("No verified old-file match; this message predates the old archive's latest write.", reuse)
        self.assertIn("no readable filename/SentOn lookup", reuse)
        self.assertIn("Plausible identity candidate limit reached", identity)
        self.assertIn('fields(1) <> "WARN"', self.source)

    def test_manual_msg_can_supply_a_candidate_without_an_old_filename(self):
        legacy = (ROOT / "mb_legacy.basfrag").read_text(encoding="ascii")
        catalog = procedure(legacy, "MBL_CatalogSentStamp")
        self.assertIn("If Len(stamp) = 0 Then stamp = MBL_ReadSentStamp(path, reason)", catalog)
        lookup = procedure(legacy, "MBL_ReadSentStamp")
        self.assertIn("Application.Session.OpenSharedItem(path)", lookup)
        self.assertIn("stamp = mail.SentOn", lookup)
        scan = procedure(legacy, "MBL_BuildIndex")
        self.assertIn("For Each child In folder.SubFolders", scan)
        self.assertIn("MBX_IsManagedPath", scan)

    def test_manual_candidate_does_not_bypass_content_or_prior_baseline(self):
        legacy = (ROOT / "mb_legacy.basfrag").read_text(encoding="ascii")
        verify = procedure(legacy, "MBL_VerifyCandidate")
        self.assertIn("Recovered MSG differs from its recorded baseline", verify)
        self.assertIn("MBL_CompareExisting mail, path", verify)
        compare = procedure(legacy, "MBL_CompareExisting")
        for symbol in ("source.Subject", "source.SentOn", "source.ReceivedTime", "source.Body", "source.Recipients.Count",
                       "MBL_AttachmentHash", "MBE_RequireMetadataComplete source, copy"):
            self.assertIn(symbol, compare)
        match = procedure(legacy, "MBL_TryMatch")
        self.assertIn("If matchCount = 0 Or differentBytes Then", match)
        self.assertIn("MBL_TryMatch = 2", match)

    def test_history_action_is_conditional_and_preserves_existing_evidence(self):
        action = procedure(self.source, "MBG_Action")
        cutoff = action.split('Case "HISTORY_CUTOFF"', 1)[1].split('Case "LEGACY_UNREADABLE"', 1)[0]
        for expected in ("Only if you confirm it was never backed up", "unused name", "Keep the source",
                         "only verified content can clear review", "Never overwrite old files"):
            self.assertIn(expected, cutoff)


if INTEGRATION:
    class IntegrationContracts(unittest.TestCase):
        @classmethod
        def setUpClass(cls):
            cls.main = (ROOT / "mb_main.basfrag").read_text(encoding="ascii")

        def test_report_is_after_stream_close_and_before_coverage(self):
            run = procedure(self.main, "MB_Run")
            self.assertLess(run.index("MB_CloseResources"), run.index("MBG_CreateReviewReport"))
            self.assertLess(run.index("MBG_CreateReviewReport"), run.index("MBW_CommitAfterClose"))

        def test_report_failure_blocks_success_via_cleanup_issue(self):
            run = procedure(self.main, "MB_Run")
            self.assertRegex(run, r"(?s)If Not MBG_CreateReviewReport\([^)]*\) Then.{0,150}gCleanupIssue")

        def test_reset_and_summary_hooks(self):
            self.assertIn("MBG_Reset", procedure(self.main, "MB_Reset"))
            self.assertIn("MBG_Summary()", procedure(self.main, "MB_Summary"))

        def test_existing_and_new_limited_validation_have_subject_guidance(self):
            process = procedure(self.main, "MB_ProcessOne")
            self.assertEqual(process.count('MBG_LogReview "SPECIAL_ATTACHMENT_CONTENT_REVIEW"'), 2)
            self.assertIn('MBG_LogReview "LEGACY_REVIEW_REQUIRED"', process)

        def test_manual_candidate_is_verified_before_the_historical_cutoff(self):
            process = procedure(self.main, "MB_ProcessOne")
            self.assertLess(process.index("MB_ReuseLegacy(mail"), process.index("MB_NewExportAllowed(mail"))
            self.assertIn("ElseIf legacyResult = 2 Then\n                    GoTo LegacyReview", process)


if __name__ == "__main__":
    unittest.main()
