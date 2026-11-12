"""Startup recovery fault models and source contracts, not Outlook/VBA execution.

The retry truth-table executes the actual simple Boolean expression extracted
from the VBA predicate.  Flow models describe required outcomes under injected
failures; source contracts tie those requirements to the production call path.
"""
from dataclasses import dataclass
import ast
from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parent


def procedure(source, name):
    match = re.search(
        rf"(?ms)^(?:Private|Public) (?:Sub|Function) {re.escape(name)}\b.*?^End (?:Sub|Function)\s*$",
        source,
    )
    if not match:
        raise AssertionError(f"VBA procedure not found: {name}")
    return match.group(0)


def executable_text(source):
    return re.sub(r'"(?:""|[^"\n])*"|\'[^\n]*', "", source)


def retry_predicate():
    """Evaluate only the actual scalar predicate, with no VBA runtime claims."""
    source = (ROOT / "mb_startup.basfrag").read_text(encoding="ascii")
    declarations = (ROOT / "mb_acceptance_declarations.basfrag").read_text(encoding="ascii")
    body = procedure(source, "MBT_CanRetry")
    joined = re.sub(r"\s+_\s*\n\s*", " ", body)
    matches = re.findall(r"(?m)^\s*MBT_CanRetry\s*=\s*([^\n]+)", joined)
    if len(matches) != 1:
        raise AssertionError("Retry policy must have one explicit scalar result")
    expression = matches[0].split("'", 1)[0].strip().lower()
    expression = expression.replace("<>", "!=")
    expression = re.sub(r"(?<![<>!=])=(?!=)", "==", expression)
    expression = re.sub(r"\btrue\b", "True", expression)
    expression = re.sub(r"\bfalse\b", "False", expression)
    parsed = ast.parse(expression, mode="eval")
    allowed = (ast.Expression, ast.BoolOp, ast.And, ast.Or, ast.UnaryOp, ast.Not,
               ast.Compare, ast.Lt, ast.LtE, ast.Gt, ast.GtE, ast.Eq, ast.NotEq,
               ast.Name, ast.Load, ast.Constant)
    if any(not isinstance(node, allowed) for node in ast.walk(parsed)):
        raise AssertionError("Retry policy contains unsupported non-scalar syntax")
    constants = {
        name.lower(): int(value)
        for name, value in re.findall(
            r"(?m)^Private Const (\w+) As Long = (\d+)\s*$", declarations
        )
    }
    code = compile(parsed, "MBT_CanRetry scalar expression", "eval")

    def evaluate(completed_attempts, cancelled=False, clean=True):
        values = dict(constants, completedattempts=completed_attempts,
                      cancelled=cancelled, clean=clean)
        return bool(eval(code, {"__builtins__": {}}, values))

    return evaluate


@dataclass(frozen=True)
class Attempt:
    passed: bool
    cancelled: bool = False
    clean: bool = True
    report: str = ""


def startup_model(outcomes, retry):
    """Required bounded flow; historical evidence never becomes a success flag."""
    evidence = []
    for count, outcome in enumerate(outcomes, 1):
        evidence.append(outcome)
        if outcome.passed and not outcome.cancelled and outcome.clean:
            return True, tuple(evidence)
        if not retry(count, outcome.cancelled, outcome.clean):
            break
    return False, tuple(evidence)


class ActualRetryPredicate(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.retry = staticmethod(retry_predicate())

    def test_exhaustive_completed_attempt_cancel_cleanup_truth_table(self):
        for attempts in (1, 2, 3, 100):
            for cancelled in (False, True):
                for clean in (False, True):
                    with self.subTest(attempts=attempts, cancelled=cancelled, clean=clean):
                        expected = attempts < 2 and not cancelled and clean
                        self.assertEqual(self.retry(attempts, cancelled, clean), expected)

    def test_first_clean_failure_is_automatically_retryable(self):
        self.assertTrue(self.retry(1, False, True))

    def test_cancelled_first_attempt_is_terminal(self):
        self.assertFalse(self.retry(1, True, True))

    def test_unknown_cleanup_is_terminal(self):
        self.assertFalse(self.retry(1, False, False))

    def test_second_failure_is_terminal_even_when_clean(self):
        self.assertFalse(self.retry(2, False, True))


class RecoveryFaultModels(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.retry = staticmethod(retry_predicate())

    def run_model(self, *outcomes):
        return startup_model(outcomes, self.retry)

    def test_first_success_starts_archive_after_one_attempt(self):
        result, evidence = self.run_model(Attempt(True, report="run1"), Attempt(False))
        self.assertTrue(result)
        self.assertEqual(len(evidence), 1)

    def test_transient_failure_recovers_without_erasing_first_failure(self):
        result, evidence = self.run_model(Attempt(False, report="run1"), Attempt(True, report="run2"))
        self.assertTrue(result)
        self.assertEqual([(item.passed, item.report) for item in evidence],
                         [(False, "run1"), (True, "run2")])

    def test_persistent_failure_never_reaches_third_attempt_or_archive(self):
        result, evidence = self.run_model(Attempt(False), Attempt(False), Attempt(True))
        self.assertFalse(result)
        self.assertEqual(len(evidence), 2)

    def test_cancelled_first_failure_never_restarts(self):
        result, evidence = self.run_model(Attempt(False, cancelled=True), Attempt(True))
        self.assertFalse(result)
        self.assertEqual(len(evidence), 1)

    def test_unclean_first_failure_never_restarts(self):
        result, evidence = self.run_model(Attempt(False, clean=False), Attempt(True))
        self.assertFalse(result)
        self.assertEqual(len(evidence), 1)

    def test_cancel_after_all_checks_pass_still_blocks_archive(self):
        result, evidence = self.run_model(Attempt(True, cancelled=True), Attempt(True))
        self.assertFalse(result)
        self.assertEqual(len(evidence), 1)

    def test_cleanup_failure_after_all_checks_pass_still_blocks_archive(self):
        result, evidence = self.run_model(Attempt(True, clean=False), Attempt(True))
        self.assertFalse(result)
        self.assertEqual(len(evidence), 1)

    def test_report_setup_failure_can_consume_one_attempt_and_recover(self):
        result, evidence = self.run_model(Attempt(False, report=""), Attempt(True, report="alternate/run2"))
        self.assertTrue(result)
        self.assertEqual(len(evidence), 2)
        self.assertEqual(evidence[0].report, "")

    def test_second_attempt_cancellation_preserves_both_attempt_records(self):
        result, evidence = self.run_model(Attempt(False, report="run1"),
                                          Attempt(False, cancelled=True, report="run2"))
        self.assertFalse(result)
        self.assertEqual([item.report for item in evidence], ["run1", "run2"])


class SourceContracts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.acceptance = (ROOT / "mb_acceptance.basfrag").read_text(encoding="ascii")
        cls.startup = (ROOT / "mb_startup.basfrag").read_text(encoding="ascii")
        cls.automatic = (ROOT / "mb_auto.basfrag").read_text(encoding="ascii")

    def test_public_command_cannot_enter_archive_when_startup_recovery_fails(self):
        body = procedure(self.automatic, "MB_01_RunBackup")
        guard = body.index("If Not MBT_RunStartupChecks(nativeReport, detail) Then")
        archive = body.index('MB_Run "AUTO", False, True, False')
        self.assertLess(guard, archive)
        guarded = body[guard:body.index("End If", guard)]
        self.assertIn("Exit Sub", guarded)
        self.assertLess(body.index("nativeTestsPassed = True"), archive)
        self.assertGreater(body.index("nativeTestsPassed = True"), body.index("End If", guard))

    def test_cached_native_success_cannot_bypass_current_resource_check(self):
        body = procedure(self.automatic, "MB_01_RunBackup")
        cache_end = body.index("End If", body.index("nativeTestsPassed = True"))
        resource_guard = body.index("If Not MBT_ResourcesClean(detail) Then")
        archive = body.index('MB_Run "AUTO", False, True, False')
        self.assertGreater(resource_guard, cache_end)
        self.assertLess(resource_guard, body.index("gAutoTestReport = nativeReport"))
        self.assertLess(resource_guard, archive)
        guarded = body[resource_guard:body.index("End If", resource_guard)]
        self.assertIn("Exit Sub", guarded)
        self.assertNotIn("MB_Reset", executable_text(body[:archive]))

    def test_success_still_requires_every_native_check(self):
        body = procedure(self.acceptance, "MBT_RunAcceptanceReport")
        result = re.search(r"(?m)^\s*MBT_RunAcceptanceReport = envOK[^\n]+", body)
        self.assertIsNotNone(result)
        for required in ("helperOK", "localeOK", "relativeOK", "propertyOK", "moveOK",
                         "recoveryOK", "plainOK", "htmlOK", "cleanupOK", "Not gCancel"):
            self.assertIn(required, result.group(0))
        self.assertNotRegex(body, r"(?m)^\s*MBT_RunAcceptanceReport = True\s*$")

    def test_startup_coordinator_does_not_repair_or_delete_user_archive(self):
        body = executable_text(procedure(self.startup, "MBT_RunStartupChecks"))
        for forbidden in ("DeleteFolder", "DeleteFile", "MB_Run", "MBY_AccountRoot",
                          "MBR_RestoreState", "MBM_ApplyPlan", "SaveSetting", "DeleteSetting",
                          ".SaveAs", ".Delete", ".Move", ".Send"):
            with self.subTest(forbidden=forbidden):
                self.assertNotIn(forbidden, body)

    def test_resources_checked_before_runner_can_reset_global_state(self):
        wrapper = procedure(self.startup, "MBT_RunStartupChecks")
        runner = procedure(self.acceptance, "MBT_RunAcceptanceReport")
        self.assertLess(wrapper.index("clean = MBT_ResourcesClean(cleanup)"),
                        wrapper.index("passed = MBT_RunAcceptanceReport"))
        initial_guard = wrapper[:wrapper.index("passed = MBT_RunAcceptanceReport")]
        self.assertRegex(initial_guard, r"(?s)If Not clean Then.*?Exit For.*?End If")
        self.assertLess(runner.index("If Not MBT_ResourcesClean(detail) Then"),
                        runner.index("    MB_Reset"))

    def test_success_and_retry_both_require_current_cleanup_and_cancel_state(self):
        body = procedure(self.startup, "MBT_RunStartupChecks")
        completed = body.index("passed = MBT_RunAcceptanceReport")
        after = body[completed:]
        self.assertLess(after.index("clean = MBT_ResourcesClean(cleanup)"),
                        after.index("If passed And Not gCancel And clean Then"))
        self.assertIn("If Not MBT_CanRetry(attempt, gCancel, clean) Then Exit For", after)
        self.assertNotIn("gCancel = False", executable_text(after))
        self.assertNotIn('mbtCleanupFault = ""', after)
        note = after.index("mbtStartupNote = MBT_WriteStartupNote")
        post_note = after[note:after.index("Failed:", note)]
        self.assertRegex(post_note, r"(?s)If MBT_RunStartupChecks Then.*?If Not MBT_ResourcesClean\(cleanup\) Then.*?MBT_RunStartupChecks = False")
        note_writer = procedure(self.startup, "MBT_WriteStartupNote")
        self.assertIn('MBT_NoteCleanup "Startup note stream close", Err.Number, Err.Description', note_writer)

    def test_loop_is_bounded_and_keeps_failed_attempt_evidence_before_success(self):
        body = procedure(self.startup, "MBT_RunStartupChecks")
        self.assertIn("For attempt = 1 To MBT_MAX_ATTEMPTS", body)
        self.assertEqual(body.count("passed = MBT_RunAcceptanceReport"), 1)
        self.assertLess(body.index("history = history &"), body.index("If passed And"))
        self.assertIn("attemptDetail & vbCrLf & mbtFailureLines", body)
        self.assertIn("MBT_WriteStartupNote(reportPath, detail & vbCrLf & vbCrLf & history)", body)

    def test_resource_predicate_observes_handles_without_discarding_them(self):
        body = procedure(self.startup, "MBT_ResourcesClean")
        for resource in ("MBH_RunLock", "mbmAccountLocks.Count", "gJournal", "gEvents",
                         "MBH_OpenHashes", "MBH_Stages.Count", "gStageEnabled",
                         "gCleanupIssue", "mbtCleanupFault"):
            self.assertIn(resource, body)
        code = executable_text(body)
        for forbidden in ("MB_Reset", "MB_CloseResources", "MBH_ReleaseLock", "MBH_ReleaseCrypto",
                          "RemoveAll", "= Nothing", "DeleteFolder", "DeleteFile"):
            self.assertNotIn(forbidden, code)
        self.assertLess(body.index("If Len(detail) > 0 Then"),
                        body.index("MBT_ResourcesClean = True"))
        self.assertIn("On Error GoTo Unknown", body)

    def test_all_five_disposable_test_creators_use_selected_parent(self):
        creators = (
            ("mb_main.basfrag", "MBT_HelperSelfTest"),
            ("mb_helpers.basfrag", "MBH_HashSelfTest"),
            ("mb_move_selftest.basfrag", "MBT_YearMonthSelfTest"),
            ("mb_operating_selftest.basfrag", "MBS_NativeRecoverySelfTest"),
            ("mb_acceptance.basfrag", "MBT_SyntheticMsgSelfTest"),
        )
        for filename, name in creators:
            with self.subTest(creator=name):
                source = (ROOT / filename).read_text(encoding="ascii")
                body = procedure(source, name)
                self.assertIn("MBT_TestParent(", body)
                self.assertNotIn("GetSpecialFolder(2)", body)

    def test_workspace_fallback_is_reversed_on_second_attempt(self):
        body = procedure(self.startup, "MBT_SelectTestParent")
        self.assertRegex(body, r'(?s)If attempt = 1 Then\s+candidates = Array\(MBT_TempParent\(fso\), Environ\$\("LOCALAPPDATA"\)\)\s+Else\s+candidates = Array\(Environ\$\("LOCALAPPDATA"\), MBT_TempParent\(fso\)\)')
        self.assertIn("For Each candidate In candidates", body)
        self.assertIn("MBT_TryScratchParent(fso, CStr(candidate), path, reason)", body)
        self.assertIn("If Len(mbtCleanupFault) > 0 Then Exit For", body)

    def test_report_fallback_is_independent_and_reversed_on_second_attempt(self):
        body = procedure(self.startup, "MBT_CreateAcceptanceReport")
        self.assertRegex(body, r'(?s)If attempt = 1 Then\s+candidates = Array\(Environ\$\("LOCALAPPDATA"\), MBT_TempParent\(fso\)\)\s+Else\s+candidates = Array\(MBT_TempParent\(fso\), Environ\$\("LOCALAPPDATA"\)\)')
        self.assertIn("For Each candidate In candidates", body)
        self.assertIn("MBT_TryReport(fso, CStr(candidate), path, reason)", body)
        self.assertIn("If Len(mbtCleanupFault) > 0 Then Exit For", body)

    def test_report_directory_and_file_are_created_without_overwrite(self):
        body = procedure(self.startup, "MBT_TryReport")
        self.assertIn('MBC_FileStamp(Now) & "_" & fso.GetTempName', body)
        self.assertIn("fso.CreateFolder runFolder", body)
        self.assertIn("fso.CreateTextFile(candidate, False, True)", body)
        self.assertLess(body.index("MBP_NoReparse parent, fso"), body.index("fso.CreateFolder runFolder"))
        self.assertLess(body.index("MBP_NoReparse runFolder, fso"), body.index("fso.CreateTextFile"))
        for forbidden in ("DeleteFolder", "DeleteFile", "MoveFile", "MoveFolder"):
            self.assertNotIn(forbidden, executable_text(body))

    def test_probe_can_only_delete_the_directory_it_exclusively_created(self):
        body = procedure(self.startup, "MBT_TryScratchParent")
        self.assertIn('"MBProbe_" & fso.GetTempName', body)
        self.assertLess(body.index("fso.CreateFolder probe"), body.index("fso.DeleteFolder probe, True"))
        self.assertLess(body.index("MBP_NoReparse probe, fso"), body.index("fso.DeleteFolder probe, True"))
        self.assertNotIn("DeleteFolder", executable_text(body.split("Failed:", 1)[1]))
        self.assertIn("fso.CreateTextFile(MBH_JoinPath(probe, \"write.txt\"), False, True)", body)

    def test_failed_check_name_and_detail_are_retained_before_report_io(self):
        body = procedure(self.acceptance, "MBT_ReportRow")
        capture = body.index('If status = "FAIL" And check <> "RUN_RESULT" Then')
        self.assertIn('mbtFailureLines = mbtFailureLines & check & ": " & detail', body)
        self.assertLess(capture, body.index("fso.OpenTextFile"))
        self.assertLess(body.index("MBP_NoReparse path, fso"), body.index("fso.OpenTextFile"))

    def test_failure_opens_diagnostic_and_also_shows_detail_without_copying(self):
        body = procedure(self.startup, "MBT_ShowStartupFailure")
        self.assertIn("path = mbtStartupNote", body)
        self.assertIn("If Len(path) = 0 Then path = reportPath", body)
        self.assertLess(body.index("opened = MBT_OpenDiagnostic(path)"), body.index("MsgBox shown"))
        self.assertIn("shown = Left$(detail, 650)", body)
        self.assertIn("A screenshot is enough", body)

    def test_diagnostic_is_quoted_explicit_text_viewer_not_a_shell_script(self):
        body = procedure(self.startup, "MBT_OpenDiagnostic")
        self.assertIn('GetSpecialFolder(1).Path), "notepad.exe"', body)
        self.assertIn('Shell(Chr$(34) & program & Chr$(34) & " " & Chr$(34) & path & Chr$(34)', body)
        self.assertLess(body.index("MBP_NoReparse path, fso"), body.index("task = Shell("))
        self.assertIn("If Not fso.FileExists(path) Then Exit Function", body)
        for forbidden in ("cmd.exe /c", "powershell", "WScript.Shell", "ShellExecute"):
            self.assertNotIn(forbidden, body)

    def test_operational_not_run_gates_are_not_relabelled_by_retry(self):
        body = procedure(self.acceptance, "MBT_RunAcceptanceReport")
        for name in ("FULL_VBA_PROJECT_COMPILE", "REAL_ARCHIVE_COMPATIBILITY",
                     "SPECIAL_ATTACHMENTS_AND_ALL_MAPI", "100000_ITEM_PERFORMANCE"):
            with self.subTest(check=name):
                self.assertIn(f'MBT_ReportRow reportPath, "{name}", "NOT_RUN"', body)
                self.assertNotIn(f'MBT_ReportRow reportPath, "{name}", "PASS"', body)


if __name__ == "__main__":
    unittest.main(verbosity=2)
