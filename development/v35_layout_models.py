"""Default directory namespaces and unchanged legacy absolute bindings.

These independent path/choice models and source contracts do not execute VBA,
Outlook, or a Windows filesystem. Their scope is namespace selection, deletion
stability, conflict refusal, and preserving historical path strings in place.
Explicit owner/hint precedence and ambiguous historical MailBackup roots are
covered by v351_layout_binding_models.py.
"""

import ntpath
from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parent
BASE = r"C:\Archive"
ACCOUNT = r"C:\Archive\Account"
MODERN_ROOT = ntpath.join(BASE, "MailBackup")
LEGACY_ROOT = ntpath.join(BASE, "SafeV2")


class ReviewRequired(ValueError):
    pass


def norm(path: str) -> str:
    return ntpath.normcase(ntpath.normpath(path))


def choose_root(base: str, children=(), remembered: str = "", *, owned_base=False) -> str:
    modern = ntpath.join(base, "MailBackup")
    legacy = ntpath.join(base, "SafeV2")
    children = {norm(path) for path in children}
    if owned_base:
        if remembered and norm(remembered) != norm(base):
            raise ReviewRequired("Owned base conflicts with remembered root")
        return base
    has_modern, has_legacy = norm(modern) in children, norm(legacy) in children
    if has_modern and has_legacy:
        raise ReviewRequired("Choose an exact owned managed root")
    if remembered:
        if norm(remembered) not in {norm(modern), norm(legacy)}:
            raise ReviewRequired("Remembered root is outside the selected base")
        if has_modern and norm(remembered) == norm(legacy):
            raise ReviewRequired("A different root exists")
        if has_legacy and norm(remembered) == norm(modern):
            raise ReviewRequired("A different root exists")
        return remembered
    return legacy if has_legacy else modern


def layout(root: str, account: str = ACCOUNT) -> tuple[str, str]:
    if ntpath.basename(norm(root)) == "mailbackup":
        return ntpath.join(root, "State"), ntpath.join(account, "_Recovery")
    return ntpath.join(root, "_MailBackup_Data"), ntpath.join(account, "_MailBackup_Recovery")


def existing_recovery(account: str, names=()) -> str:
    names = {name.casefold() for name in names}
    modern, legacy = "_recovery" in names, "_mailbackup_recovery" in names
    if modern and legacy:
        raise ReviewRequired("Two independent recovery histories")
    if modern:
        return ntpath.join(account, "_Recovery")
    if legacy:
        return ntpath.join(account, "_MailBackup_Recovery")
    return ""


def validate_layout(root: str, *, populated_states=(), recovery_names=()) -> tuple[str, str]:
    state, recovery = layout(root)
    expected_leaf = ntpath.basename(state).casefold()
    if any(name.casefold() != expected_leaf for name in populated_states):
        raise ReviewRequired("Populated state uses another namespace")
    existing = existing_recovery(ACCOUNT, recovery_names)
    if existing and norm(existing) != norm(recovery):
        raise ReviewRequired("Existing recovery cannot move to the new name")
    return state, recovery


def source(name: str) -> str:
    return (ROOT / name).read_text(encoding="ascii")


def procedure(filename: str, name: str) -> str:
    match = re.search(
        rf"(?ms)^(?:Private|Public) (?:Sub|Function) {re.escape(name)}\b.*?^End (?:Sub|Function)\s*$",
        source(filename),
    )
    if match is None:
        raise AssertionError(f"Missing procedure: {filename}:{name}")
    return "\n".join(line for line in match.group().splitlines()
                     if not line.lstrip().startswith("'"))


class LayoutScenarios(unittest.TestCase):
    def test_fresh_base_uses_versionless_managed_root(self):
        self.assertEqual(choose_root(BASE), MODERN_ROOT)

    def test_existing_legacy_child_is_reused_at_exact_path(self):
        self.assertEqual(choose_root(BASE, [LEGACY_ROOT]), LEGACY_ROOT)

    def test_existing_modern_child_is_reused_at_exact_path(self):
        self.assertEqual(choose_root(BASE, [MODERN_ROOT]), MODERN_ROOT)

    def test_two_managed_children_require_an_explicit_choice(self):
        with self.assertRaises(ReviewRequired):
            choose_root(BASE, [MODERN_ROOT, LEGACY_ROOT])

    def test_deleted_remembered_legacy_root_keeps_its_path(self):
        self.assertEqual(choose_root(BASE, remembered=LEGACY_ROOT), LEGACY_ROOT)

    def test_deleted_remembered_modern_root_keeps_its_path(self):
        self.assertEqual(choose_root(BASE, remembered=MODERN_ROOT), MODERN_ROOT)

    def test_missing_old_root_does_not_adopt_new_sibling(self):
        with self.assertRaises(ReviewRequired):
            choose_root(BASE, [MODERN_ROOT], LEGACY_ROOT)

    def test_missing_new_root_does_not_adopt_old_sibling(self):
        with self.assertRaises(ReviewRequired):
            choose_root(BASE, [LEGACY_ROOT], MODERN_ROOT)

    def test_remembered_path_outside_base_is_rejected(self):
        with self.assertRaises(ReviewRequired):
            choose_root(BASE, remembered=r"D:\Other\SafeV2")

    def test_selecting_exact_owned_root_resolves_sibling_ambiguity(self):
        self.assertEqual(choose_root(LEGACY_ROOT, owned_base=True), LEGACY_ROOT)

    def test_new_archive_uses_state_and_recovery_roles(self):
        self.assertEqual(layout(MODERN_ROOT),
                         (r"C:\Archive\MailBackup\State", r"C:\Archive\Account\_Recovery"))

    def test_old_archive_keeps_full_absolute_tuple(self):
        self.assertEqual(layout(LEGACY_ROOT),
                         (r"C:\Archive\SafeV2\_MailBackup_Data", r"C:\Archive\Account\_MailBackup_Recovery"))

    def test_exactly_owned_custom_legacy_root_keeps_legacy_names(self):
        self.assertEqual(layout(r"D:\HistoricalRoot")[0], r"D:\HistoricalRoot\_MailBackup_Data")

    def test_root_name_case_does_not_change_layout(self):
        self.assertEqual(ntpath.basename(layout(r"C:\Archive\MAILBACKUP")[0]), "State")

    def test_primary_deletion_keeps_same_restore_destination(self):
        for root in (MODERN_ROOT, LEGACY_ROOT):
            with self.subTest(root=root):
                before = layout(root)
                recovery_leaf = ntpath.basename(before[1])
                self.assertEqual(validate_layout(root, recovery_names=[recovery_leaf]), before)

    def test_recovery_deletion_keeps_same_recreation_destination(self):
        for root in (MODERN_ROOT, LEGACY_ROOT):
            with self.subTest(root=root):
                before = layout(root)
                self.assertEqual(validate_layout(root, populated_states=[ntpath.basename(before[0])]), before)

    def test_settings_and_both_metadata_directories_deleted_do_not_change_namespace(self):
        for root in (MODERN_ROOT, LEGACY_ROOT):
            with self.subTest(root=root):
                self.assertEqual(validate_layout(root), layout(root))

    def test_modern_root_with_legacy_state_requires_review(self):
        with self.assertRaises(ReviewRequired):
            validate_layout(MODERN_ROOT, populated_states=["_MailBackup_Data"])

    def test_legacy_root_with_new_state_requires_review(self):
        with self.assertRaises(ReviewRequired):
            validate_layout(LEGACY_ROOT, populated_states=["State"])

    def test_two_recovery_histories_are_not_merged_by_name(self):
        with self.assertRaises(ReviewRequired):
            existing_recovery(ACCOUNT, ["_Recovery", "_MailBackup_Recovery"])

    def test_mismatched_recovery_directory_is_not_silently_relocated(self):
        for root, leaf in ((MODERN_ROOT, "_MailBackup_Recovery"), (LEGACY_ROOT, "_Recovery")):
            with self.subTest(root=root), self.assertRaises(ReviewRequired):
                validate_layout(root, recovery_names=[leaf])

    def test_old_and_new_bound_archives_roundtrip_without_path_changes(self):
        for root in (MODERN_ROOT, LEGACY_ROOT):
            with self.subTest(root=root):
                before = layout(root)
                after = validate_layout(root, populated_states=[ntpath.basename(before[0])],
                                        recovery_names=[ntpath.basename(before[1])])
                self.assertEqual(before, after)


class LayoutSourceContracts(unittest.TestCase):
    def test_new_and_legacy_names_are_declared_separately(self):
        declarations = "\n".join(source(name) for name in (
            "mb_paths_declarations.basfrag", "mb_operating_declarations.basfrag", "mb_recovery_declarations.basfrag"))
        for constant, value in (("MBP_MANAGED_LEAF", "MailBackup"), ("MBP_MANAGED_LEAF_LEGACY", "SafeV2"),
                                ("MBS_DATA_LEAF", "State"), ("MBS_DATA_LEAF_LEGACY", "_MailBackup_Data"),
                                ("MBR_FOLDER", "_Recovery"), ("MBR_FOLDER_LEGACY", "_MailBackup_Recovery")):
            self.assertIn(f'{constant} As String = "{value}"', declarations)

    def test_locator_accepts_both_child_names_with_same_root_binding(self):
        body = procedure("mb_paths.basfrag", "MBP_ReadLocator")
        self.assertIn("MBH_JoinPath(base, MBP_MANAGED_LEAF)", body)
        self.assertIn("MBH_JoinPath(base, MBP_MANAGED_LEAF_LEGACY)", body)
        self.assertIn("mbpLocatorManaged = managed", body)

    def test_known_missing_root_is_passed_through_to_preflight(self):
        resolve = procedure("mb_paths.basfrag", "MBP_ResolveBackupRoot")
        self.assertIn("MBP_TryBackupBase(locatedBase, managed, reason, locatedRoot)", resolve)
        self.assertIn("MBP_TryBackupBase(savedBase, savedRoot, savedReason, locatedRoot)", resolve)
        self.assertIn("managed = rememberedRoot", procedure("mb_paths.basfrag", "MBP_TryBackupBase"))

    def test_two_root_candidates_are_rejected_before_any_write_probe(self):
        body = procedure("mb_paths.basfrag", "MBP_TryBackupBase")
        self.assertLess(body.index("If hasModern And hasLegacy Then Err.Raise"),
                        body.index("MBP_ProbeWritable probeParent, fso"))

    def test_state_path_uses_namespace_resolver(self):
        self.assertIn("MBH_JoinPath(gRoot, MBS_DataLeaf(gRoot))", procedure("mb_operating.basfrag", "MB_StatePath"))

    def test_namespace_uses_persistent_evidence_and_cached_choice(self):
        body = procedure("mb_operating.basfrag", "MBS_IsModernLayout")
        self.assertIn("MBS_LayoutKind(managedRoot)", body)
        resolve = procedure("mb_layout_binding.basfrag", "MBS_ResolveLayout")
        self.assertIn("MBS_LayoutOwnerKind(value)", resolve)
        self.assertIn("MBS_LayoutSettingValue(saved, rootPath)", resolve)
        self.assertIn("MBS_LayoutRecoveryEvidence rootPath, fso, selected", resolve)
        self.assertIn("gLayoutRoots(rootPath) = selected", resolve)
        self.assertIn("If owned And defaultModern Then Err.Raise", resolve)

    def test_mixed_primary_layout_is_checked_before_writes(self):
        preflight = procedure("mb_paths.basfrag", "MBP_TryBackupBase")
        self.assertLess(preflight.index("MBS_ValidateLayout managed, fso"), preflight.index("MBP_ProbeWritable"))
        initialize = procedure("mb_operating.basfrag", "MBS_Initialize")
        self.assertLess(initialize.index("MBS_ValidateLayout gRoot, gFSO"), initialize.index("MBH_EnsureDirectory gStateRoot"))

    def test_recovery_reader_accepts_both_but_writer_checks_expected_path(self):
        reader = procedure("mb_recovery.basfrag", "MBR_ReadOwner")
        self.assertIn("MBH_JoinPath(CStr(fields(2)), MBR_FOLDER)", reader)
        self.assertIn("MBH_JoinPath(CStr(fields(2)), MBR_FOLDER_LEGACY)", reader)
        writer = procedure("mb_recovery.basfrag", "MBR_RecoveryRoot")
        self.assertIn("root = MBR_ExpectedRecoveryPath(accountRoot, gRoot)", writer)
        self.assertLess(writer.index("StrComp(existing, root, vbTextCompare)"), writer.index("MBH_EnsureDirectory root"))

    def test_owner_and_store_discovery_read_both_previous_names(self):
        owner = procedure("mb_operating.basfrag", "MBS_CanRecoverOwner")
        store = procedure("mb_store_binding.basfrag", "MBS_PreflightCandidateEvidence")
        for body in (owner, store):
            self.assertIn('Array(MBS_DATA_LEAF, MBS_DATA_LEAF_LEGACY, "_state")', body)
        self.assertIn("MBS_OwnerRecoveryEvidence", owner)
        for body in (procedure("mb_operating.basfrag", "MBS_OwnerRecoveryEvidence"),
                     procedure("mb_store_binding.basfrag", "MBS_PreflightCandidateAccount")):
            self.assertIn("Array(MBR_FOLDER, MBR_FOLDER_LEGACY)", body)
            self.assertNotIn("MBR_ExistingRecoveryPath", body)

    def test_namespace_resolution_does_not_rename_or_copy_previous_archives(self):
        bodies = "\n".join(procedure(filename, name) for filename, name in (
            ("mb_paths.basfrag", "MBP_TryBackupBase"),
            ("mb_operating.basfrag", "MBS_DataLeaf"),
            ("mb_operating.basfrag", "MBS_ValidateLayout"),
            ("mb_recovery.basfrag", "MBR_ExpectedRecoveryPath"),
            ("mb_recovery.basfrag", "MBR_ExistingRecoveryPath"),
        ))
        for mutation in ("MoveFile", "MoveFolder", "CopyFile", "DeleteFile", "DeleteFolder", "SaveSetting", "SaveAs"):
            self.assertNotIn(mutation, bodies)


if __name__ == "__main__":
    unittest.main(verbosity=2)
