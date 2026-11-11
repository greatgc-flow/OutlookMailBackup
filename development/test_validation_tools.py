"""Mutation tests for validation gates; these do not execute Outlook/VBA."""
from contextlib import redirect_stdout, redirect_stderr
import io
import hashlib
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tools"))
import build
import test as model_runner
import validate
import release_evidence

SOURCE = '''Attribute VB_Name = "OutlookMailBackup"
Option Explicit
Private MBX_Value As Long
Private gValue As Long
#If VBA7 Then
Private Const MBX_Flag As Boolean = True
#Else
Private Const MBX_Flag As Boolean = False
#End If
Public Sub MB_01_RunBackup()
    MBX_Helper
End Sub
Private Sub MBX_Helper()
    gValue = MBX_Value
End Sub
'''


def data(source=SOURCE):
    return source.replace("\n", "\r\n").encode("ascii")


class ArtifactMutationTests(unittest.TestCase):
    def reject(self, source, needle):
        report = validate.inspect_bytes(data(source))
        messages = report["errors"] + [e for p in report["platforms"].values() for e in p["errors"]]
        self.assertFalse(validate.successful(report))
        self.assertTrue(any(needle in message for message in messages), messages)

    def test_valid_artifact_all_three_platforms(self):
        result = validate.inspect_bytes(data())
        self.assertTrue(validate.successful(result), result)
        self.assertEqual(set(result["platforms"]), set(validate.PLATFORMS))

    def test_project_identifiers_are_case_insensitive(self):
        self.assertTrue(validate.successful(validate.inspect_bytes(data(SOURCE.replace("    MBX_Helper", "    mbx_helper")))))

    def test_comments_and_strings_do_not_create_false_dependencies(self):
        source = SOURCE.replace("    gValue = MBX_Value", '    gValue = MBX_Value\n    Debug.Print "MBX_Ghost \' x" \' MBX_Ghost\n    Rem MBX_Ghost')
        self.assertTrue(validate.successful(validate.inspect_bytes(data(source))))

    def test_strings_cannot_supply_a_missing_declaration(self):
        self.reject(SOURCE.replace("    MBX_Helper", '    MBX_Ghost\n    Debug.Print "Private Sub MBX_Ghost()"'), "Undefined project symbol")

    def test_missing_project_procedure_is_rejected(self):
        self.reject(SOURCE.replace("    MBX_Helper", "    MBX_Unknown"), "Undefined project symbol")

    def test_missing_global_is_rejected(self):
        self.reject(SOURCE.replace("    gValue", "    gMissing"), "Undefined project symbol")

    def test_duplicate_procedure_is_rejected(self):
        self.reject(SOURCE + "Private Sub MBX_Helper()\nEnd Sub\n", "Duplicate module symbol")

    def test_duplicate_constant_is_rejected(self):
        self.reject(SOURCE.replace("Private MBX_Value As Long", "Private MBX_Value As Long\nPrivate Const MBX_Value As Long = 1"), "Duplicate module symbol")

    def test_extra_public_macro_is_rejected(self):
        self.reject(SOURCE.replace("Private Sub MBX_Helper", "Public Sub MBX_Helper"), "Unexpected public")

    def test_default_public_macro_is_rejected(self):
        self.reject(SOURCE.replace("Private Sub MBX_Helper", "Sub MBX_Helper"), "Unexpected public")

    def test_entry_point_with_parameter_is_rejected(self):
        self.reject(SOURCE.replace("MB_01_RunBackup()", "MB_01_RunBackup(ByVal arg As Long)"), "Unexpected public")

    def test_missing_entry_point_is_rejected(self):
        self.reject(SOURCE.replace("Public Sub MB_01_RunBackup", "Private Sub MB_01_RunBackup"), "Exactly one public")

    def test_wrong_module_attribute_is_rejected(self):
        self.reject(SOURCE.replace('VB_Name = "OutlookMailBackup"', 'VB_Name = "Other"'), "module attribute")

    def test_missing_option_explicit_is_rejected(self):
        self.reject(SOURCE.replace("Option Explicit", ""), "Option Explicit")

    def test_lf_artifact_is_rejected(self):
        self.assertFalse(validate.successful(validate.inspect_bytes(SOURCE.encode())))

    def test_bare_cr_artifact_is_rejected(self):
        self.assertFalse(validate.successful(validate.inspect_bytes(data().replace(b"\r\n", b"\r", 1))))

    def test_utf8_bom_is_rejected(self):
        self.assertFalse(validate.successful(validate.inspect_bytes(b"\xef\xbb\xbf" + data())))

    def test_standalone_end_is_rejected(self):
        self.reject(SOURCE.replace("    MBX_Helper", "    End"), "standalone End")

    def test_late_global_declaration_is_rejected(self):
        self.reject(SOURCE + "Private MBX_Late As Long\n", "Module declaration after")

    def test_unclosed_procedure_is_rejected(self):
        self.reject(SOURCE.removesuffix("End Sub\n"), "Unclosed procedure")

    def test_mismatched_procedure_end_is_rejected(self):
        self.reject(SOURCE.replace("End Sub", "End Function", 1), "Mismatched End")

    def test_overlong_line_is_rejected(self):
        self.reject(SOURCE + "'" + "x" * 1023 + "\n", "1023 characters")

    def test_excessive_continuation_is_rejected(self):
        self.reject(SOURCE.replace("    MBX_Helper", "    Debug.Print " + ("1 + _\n        " * 25) + "1"), "24 line continuations")

    def test_inactive_platform_bug_is_still_detected(self):
        source = SOURCE.replace("    MBX_Helper", "#If Win64 Then\n    MBX_Unknown\n#Else\n    MBX_Helper\n#End If")
        report = validate.inspect_bytes(data(source))
        self.assertFalse(validate.successful(report))
        self.assertFalse(report["platforms"]["vba7_32"]["errors"])
        self.assertTrue(report["platforms"]["vba7_64"]["errors"])


class ConditionalMutationTests(unittest.TestCase):
    def test_boolean_precedence_and_nested_parentheses(self):
        self.assertTrue(validate.condition("VBA7 And (Not Win64 Or False)", validate.PLATFORMS["vba7_32"]))

    def test_numeric_not_uses_vba_bitwise_semantics(self):
        self.assertTrue(validate.condition("Not 1", {}))
        self.assertFalse(validate.condition("Not True", {}))

    def test_unknown_symbol_not_hidden_by_true_left_operand(self):
        with self.assertRaises(ValueError):
            validate.condition("True Or Unknown", {})

    def test_unknown_expression_is_rejected(self):
        with self.assertRaises(ValueError):
            validate.condition("VBA7 = 1", validate.PLATFORMS["vba7_32"])

    def test_duplicate_else_is_rejected(self):
        with self.assertRaises(ValueError):
            validate.preprocessor("#If VBA7 Then\n#Else\n#Else\n#End If", validate.PLATFORMS["vba7_32"])

    def test_elseif_after_else_is_rejected(self):
        with self.assertRaises(ValueError):
            validate.preprocessor("#If VBA7 Then\n#Else\n#ElseIf Win64 Then\n#End If", validate.PLATFORMS["vba7_32"])

    def test_unknown_directive_is_rejected(self):
        with self.assertRaises(ValueError):
            validate.preprocessor("#Const Anything = 1", {})


class RunnerGateTests(unittest.TestCase):
    def test_clean_unittest_success_is_accepted(self):
        self.assertEqual(model_runner.suite_passed(0, "Ran 3 tests in 0.1s\n\nOK\n{}", 3), (True, 3))

    def test_skipped_tests_are_not_passed_tests(self):
        self.assertFalse(model_runner.suite_passed(0, "Ran 3 tests in 0.1s\n\nOK (skipped=1)\n", 3)[0])

    def test_expected_failures_are_not_passed_tests(self):
        self.assertFalse(model_runner.suite_passed(0, "Ran 3 tests in 0.1s\n\nOK (expected failures=1)\n", 3)[0])

    def test_wrong_count_is_rejected(self):
        self.assertFalse(model_runner.suite_passed(0, "Ran 2 tests in 0.1s\n\nOK\n", 3)[0])

    def test_failure_exit_is_rejected_even_with_success_text(self):
        self.assertFalse(model_runner.suite_passed(1, "Ran 3 tests in 0.1s\n\nOK\n", 3)[0])

    def test_multiple_test_summaries_are_rejected(self):
        self.assertFalse(model_runner.suite_passed(0, "Ran 3 tests in 0.1s\nOK\nRan 1 test in 0.1s\nOK\n", 3)[0])

    def test_build_check_rejects_changed_artifact_without_overwriting_it(self):
        with tempfile.TemporaryDirectory() as temporary:
            artifact = Path(temporary) / "OutlookMailBackup.bas"
            artifact.write_bytes(b"changed artifact")
            with mock.patch.object(build, "ARTIFACT", artifact), mock.patch.object(sys, "argv", ["build.py", "--check"]):
                with redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
                    status = build.main()
            self.assertEqual(status, 1)
            self.assertEqual(artifact.read_bytes(), b"changed artifact")


class EvidenceMutationTests(unittest.TestCase):
    def setUp(self):
        self.payload = data(SOURCE.replace("Option Explicit", 'Option Explicit\nPrivate Const MB_VERSION As String = "test"'))
        self.report = {"artifact": dict(name="OutlookMailBackup.bas", version="test", bytes=len(self.payload),
                                       lines=len(self.payload.decode().splitlines()),
                                       sha256=hashlib.sha256(self.payload).hexdigest(),
                                       public_macros=["MB_01_RunBackup"])}

    def test_current_evidence_is_accepted(self):
        self.assertEqual(release_evidence.check_evidence(self.payload, self.report), [])

    def test_old_bas_digest_is_rejected(self):
        self.report["artifact"]["sha256"] = "0" * 64
        self.assertTrue(release_evidence.check_evidence(self.payload, self.report))

    def test_old_byte_count_is_rejected(self):
        self.report["artifact"]["bytes"] -= 1
        self.assertTrue(release_evidence.check_evidence(self.payload, self.report))

    def test_old_line_count_is_rejected(self):
        self.report["artifact"]["lines"] -= 1
        self.assertTrue(release_evidence.check_evidence(self.payload, self.report))

    def test_old_version_is_rejected(self):
        self.report["artifact"]["version"] = "older"
        self.assertTrue(release_evidence.check_evidence(self.payload, self.report))

    def test_different_entrypoint_is_rejected(self):
        self.report["artifact"]["public_macros"].append("MB_Other")
        self.assertTrue(release_evidence.check_evidence(self.payload, self.report))

    def test_missing_manifest_is_rejected(self):
        self.assertTrue(release_evidence.check_evidence(self.payload, {}))

    def test_nested_old_grammar_result_is_rejected(self):
        self.report["grammar"] = {"check": "ANTLR grammar", "sha256": "0" * 64, "result": "PASS"}
        errors = release_evidence.check_evidence(self.payload, self.report)
        self.assertTrue(any("validation.grammar.sha256" in error for error in errors))

    def test_unrelated_raw_source_hash_is_not_a_bas_hash(self):
        self.report["source_hashes"] = [{"path": "source.basfrag", "sha256": "0" * 64}]
        self.assertEqual(release_evidence.check_evidence(self.payload, self.report), [])


if __name__ == "__main__":
    unittest.main()
