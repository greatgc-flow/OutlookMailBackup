"""Bounded fidelity cases and source contracts; NOT Outlook/VBA execution."""
from dataclasses import dataclass, replace
from pathlib import Path
import json
import re
import sys
import unittest

ROOT = Path(__file__).resolve().parent
SOURCE = (ROOT / "mb_fidelity.basfrag").read_text(encoding="ascii")
MISSING, ACCESS_DENIED = object(), object()


class MetadataError(Exception):
    def __init__(self, code):
        self.code = code


def categories(raw, display):
    if raw is MISSING:
        if display:
            raise MetadataError(1124)
        return frozenset()
    if not isinstance(raw, (list, tuple)) or len(raw) > 1024:
        raise MetadataError(1124)
    for name in raw:
        if not isinstance(name, str) or not 0 < len(name.encode("utf-16-le")) // 2 < 256:
            raise MetadataError(1124)
    names = frozenset(raw)
    if bool(names) != bool(display):
        raise MetadataError(1124)
    return names


@dataclass(frozen=True)
class Mail:
    importance: int = 1
    sensitivity: int = 0
    keywords: object = MISSING
    display_categories: str = ""
    custom_count: int = 0


def complete(first, second):
    if not (0 <= first.importance <= 2 and 0 <= second.importance <= 2):
        raise MetadataError(1124)
    if first.importance != second.importance:
        raise MetadataError(1120)
    if not (0 <= first.sensitivity <= 3 and 0 <= second.sensitivity <= 3):
        raise MetadataError(1124)
    if first.sensitivity != second.sensitivity:
        raise MetadataError(1121)
    if categories(first.keywords, first.display_categories) != categories(second.keywords, second.display_categories):
        raise MetadataError(1122)
    if first.custom_count < 0 or second.custom_count < 0:
        raise MetadataError(1124)
    return first.custom_count == second.custom_count == 0


def legacy_result(first, second):
    try:
        return "EQUAL" if complete(first, second) else "REVIEW"
    except MetadataError as error:
        return "DIFFERENT" if error.code in {1120, 1121, 1122} else "REVIEW"


class FidelityCases(unittest.TestCase):
    def test_ordinary_equal_metadata_can_reuse(self):
        self.assertTrue(complete(Mail(), Mail()))
        self.assertEqual(legacy_result(Mail(), Mail()), "EQUAL")

    def test_importance_only_change_is_a_real_revision(self):
        self.assertEqual(legacy_result(replace(Mail(), importance=2), Mail()), "DIFFERENT")

    def test_sensitivity_only_change_is_a_real_revision(self):
        self.assertEqual(legacy_result(replace(Mail(), sensitivity=3), Mail()), "DIFFERENT")

    def test_category_only_change_is_a_real_revision(self):
        edited = replace(Mail(), keywords=["red"], display_categories="red")
        self.assertEqual(legacy_result(edited, Mail()), "DIFFERENT")

    def test_category_order_and_display_delimiter_do_not_change_membership(self):
        first = Mail(keywords=["comma,name", "semi;colon", "한글"], display_categories="comma,name; semi;colon; 한글")
        second = Mail(keywords=["한글", "semi;colon", "comma,name"], display_categories="한글, semi;colon, comma,name")
        self.assertTrue(complete(first, second))

    def test_a_comma_inside_one_name_is_not_two_names(self):
        first = Mail(keywords=["A,B"], display_categories="A,B")
        second = Mail(keywords=["A", "B"], display_categories="A,B")
        self.assertEqual(legacy_result(first, second), "DIFFERENT")

    def test_whitespace_is_not_trimmed_away(self):
        first = Mail(keywords=[" padded "], display_categories=" padded ")
        second = Mail(keywords=["padded"], display_categories="padded")
        self.assertEqual(legacy_result(first, second), "DIFFERENT")

    def test_unicode_case_and_normalization_remain_exact(self):
        for first_name, second_name in (("Tag", "tag"), ("é", "e\u0301"), ("한글😀", "한글😁")):
            first = Mail(keywords=[first_name], display_categories=first_name)
            second = Mail(keywords=[second_name], display_categories=second_name)
            self.assertEqual(legacy_result(first, second), "DIFFERENT")

    def test_duplicate_category_membership_is_idempotent(self):
        self.assertEqual(categories(["A", "A"], "A"), categories(["A"], "A"))

    def test_missing_and_empty_category_list_mean_no_membership(self):
        self.assertEqual(categories(MISSING, ""), categories([], ""))

    def test_missing_keywords_with_visible_categories_is_review(self):
        first = replace(Mail(), display_categories="hidden original name")
        self.assertEqual(legacy_result(first, Mail()), "REVIEW")

    def test_access_denied_is_not_absence_or_known_difference(self):
        first = replace(Mail(), keywords=ACCESS_DENIED)
        self.assertEqual(legacy_result(first, Mail()), "REVIEW")

    def test_null_scalar_or_nested_array_never_becomes_empty_categories(self):
        for raw in (None, "red", 0, [123], [["red"]], [""], ["x" * 256]):
            with self.subTest(raw=repr(raw)[:50]):
                self.assertEqual(legacy_result(replace(Mail(), keywords=raw), Mail()), "REVIEW")

    def test_category_array_display_disagreement_is_review(self):
        for raw, display in (([], "shown"), (["present"], "")):
            self.assertEqual(legacy_result(Mail(keywords=raw, display_categories=display), Mail()), "REVIEW")

    def test_category_budget_does_not_truncate_to_false_equality(self):
        raw = [str(index) for index in range(1025)]
        self.assertEqual(legacy_result(Mail(keywords=raw, display_categories="many"), Mail()), "REVIEW")

    def test_new_custom_properties_get_limited_validation(self):
        self.assertFalse(complete(replace(Mail(), custom_count=1), Mail()))

    def test_equal_custom_counts_do_not_prove_equal_custom_values(self):
        self.assertFalse(complete(replace(Mail(), custom_count=2), replace(Mail(), custom_count=2)))
        self.assertEqual(legacy_result(replace(Mail(), custom_count=2), replace(Mail(), custom_count=2)), "REVIEW")

    def test_custom_properties_only_in_old_copy_still_block_alias(self):
        self.assertEqual(legacy_result(Mail(), replace(Mail(), custom_count=1)), "REVIEW")

    def test_invalid_enum_or_custom_count_is_unknown(self):
        for edited in (replace(Mail(), importance=9), replace(Mail(), sensitivity=-1), replace(Mail(), custom_count=-1)):
            self.assertEqual(legacy_result(edited, Mail()), "REVIEW")


class SourceContracts(unittest.TestCase):
    def test_exact_category_mapi_array_has_no_locale_split_or_trim(self):
        declarations = (ROOT / "mb_fidelity_declarations.basfrag").read_text(encoding="ascii")
        self.assertIn("{00020329-0000-0000-C000-000000000046}/Keywords", declarations)
        code = re.sub(r'"(?:[^"\r\n]|"")*"', '""', SOURCE)
        code = "\n".join(line.split("'", 1)[0] for line in code.splitlines())
        self.assertNotRegex(code, r"(?i)\b(?:Split|Trim|LCase|UCase)\$?\s*\(")
        self.assertIn("names.CompareMode = vbBinaryCompare", SOURCE)

    def test_both_content_comparators_call_same_metadata_contract(self):
        main = (ROOT / "mb_main.basfrag").read_text(encoding="ascii")
        legacy = (ROOT / "mb_legacy.basfrag").read_text(encoding="ascii")
        comparator = main.split("Private Sub MB_CompareMsg", 1)[1].split("End Sub", 1)[0]
        self.assertIn("If Not MBE_MetadataComplete(source, copy) Then contentComplete = False", comparator)
        comparator = legacy.split("Private Sub MBL_CompareExisting", 1)[1].split("End Sub", 1)[0]
        self.assertIn("MBE_RequireMetadataComplete source, copy", comparator)
        self.assertLess(comparator.index("Legacy body mismatch"), comparator.index("MBE_RequireMetadataComplete"))

    def test_nonzero_custom_count_is_explicitly_limited(self):
        self.assertIn("firstCustom = source.UserProperties.Count: secondCustom = copy.UserProperties.Count", SOURCE)
        self.assertIn("MBE_MetadataComplete = (firstCustom = 0 And secondCustom = 0)", SOURCE)
        self.assertIn("Err.Raise MB_ERR + 1123", SOURCE)

    def test_readonly_comparison_and_disposable_native_test_are_separate(self):
        production = SOURCE.split("Private Function MBE_MetadataResultCode", 1)[0]
        self.assertNotRegex(production, r"(?m)^\s*(?:source|copy)\.\w+\s*=")
        test = SOURCE.split("Private Function MBE_SyntheticMetadataTest", 1)[1]
        self.assertIn('gMode <> "SELFTEST_MSG" Or gRunId <> "synthetic_msg"', test)
        self.assertIn("If Not MB_PathWithin(path, gRoot)", test)
        self.assertIn("If source.Sent Then", test)
        self.assertNotRegex(SOURCE, r"(?i)\.\s*(?:SaveAs|Save|Send)\b")

    def test_full_mapi_and_flag_scope_is_not_claimed(self):
        self.assertIn("Flags/tasks/reminders and arbitrary MAPI properties remain outside this scope.", SOURCE)

    def test_helpers_add_no_public_macro(self):
        self.assertNotRegex(SOURCE, r"(?m)^Public\s+(?:Function|Sub)\b")

    def test_synthetic_custom_property_is_legal_and_not_added_to_folder(self):
        match = re.search(r'UserProperties.Add\("([^"]+)", olText, False\)', SOURCE)
        self.assertIsNotNone(match)
        self.assertLessEqual(len(match[1]), 64)
        self.assertFalse(any(char in match[1] for char in "[]_#"))


class IntegrationContracts(unittest.TestCase):
    def test_metadata_mismatch_but_not_unknown_is_authorized_as_difference(self):
        identity = (ROOT / "mb_identity.basfrag").read_text(encoding="ascii")
        code = identity.split("Private Function MBI_CompareContents", 1)[1].split("End Function", 1)[0]
        self.assertRegex(code, r"Case[^\n]*(?:1120 To 1122|1120, 1121, 1122)")
        self.assertNotRegex(code, r"Case[^\n]*112[34]")

    def test_fidelity_files_are_in_assembly_and_native_test_is_connected(self):
        assembly = (ROOT / "mb_assemble.py").read_text(encoding="ascii")
        self.assertIn("mb_fidelity_declarations.basfrag", assembly)
        self.assertIn("mb_fidelity.basfrag", assembly)
        acceptance = (ROOT / "mb_acceptance.basfrag").read_text(encoding="ascii")
        self.assertIn("MBE_CategorySelfTest()", acceptance)
        self.assertIn("MBE_SyntheticMetadataTest(reopened, copyPath, detail)", acceptance)


if __name__ == "__main__":
    integrated = "--integration" in sys.argv
    if integrated:
        sys.argv.remove("--integration")
    suite = unittest.TestSuite()
    for test in (FidelityCases, SourceContracts) + ((IntegrationContracts,) if integrated else ()):
        suite.addTests(unittest.TestLoader().loadTestsFromTestCase(test))
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    report = {
        "kind": "bounded fidelity models/source contracts; NOT Outlook/VBA execution",
        "integration_hooks_checked": integrated,
        "tests": result.testsRun, "failures": len(result.failures), "errors": len(result.errors),
        "pass": result.wasSuccessful(), "native_outlook_runtime": "NOT_RUN",
        "every_mapi_property_fidelity": "NOT_CLAIMED",
    }
    (ROOT / "vba_parser_review/v34_fidelity_models.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps(report, indent=2))
    raise SystemExit(not result.wasSuccessful())
