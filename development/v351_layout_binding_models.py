"""Persistent layout-selection regressions, not Outlook/VBA execution.

The scenarios distinguish namespace hints from ownership authority, keep legacy
custom roots in place, and exercise loss/conflict combinations. Source contracts
connect this independent model to the actual preflight, owner writer and cache.
"""

from dataclasses import dataclass
import itertools
import ntpath
from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parent
MODERN = "MODERN"
LEGACY = "LEGACY"
GENERIC = "OutlookMailBackup.Schema4"
NEW_MODERN = GENERIC + ".Modern"
NEW_LEGACY = GENERIC + ".Legacy"
OLD = ("OutlookMailBackup.SafeV2", "OutlookMailBackup.SafeV3.3",
       "OutlookMailBackup.SafeV3.4")
MAIL_ROOT = r"C:\Archive\MailBackup"
OLD_ROOT = r"C:\Archive\SafeV2"
CUSTOM_ROOT = r"C:\Archive\Historical"
ACCOUNT = r"C:\Archive\Account"


class ReviewRequired(ValueError):
    pass


def norm(value):
    return ntpath.normcase(ntpath.normpath(value))


@dataclass(frozen=True)
class Hint:
    root: str
    kind: str
    checked: bool = True


@dataclass(frozen=True)
class Recovery:
    root: str
    account: str
    folder_account: str
    kind: str
    checked: bool = True


def owner_kind(value):
    if value in OLD or value == NEW_LEGACY:
        return LEGACY
    if value == NEW_MODERN:
        return MODERN
    if value == GENERIC:
        return None
    raise ReviewRequired("Unknown owner generation")


def decide(root, *, owner=None, hint=None, states=(), recoveries=(),
           ownership_proven=False, cache=None, state_files=()):
    """Choose a namespace; this function does not establish ownership."""
    selected = cache

    def add(candidate):
        nonlocal selected
        if candidate is None:
            return
        if candidate not in (MODERN, LEGACY):
            raise ReviewRequired("Unknown layout kind")
        if selected is not None and selected != candidate:
            raise ReviewRequired("Conflicting layout evidence")
        selected = candidate

    owned = ownership_proven or owner is not None
    if owner is not None:
        add(owner_kind(owner))
    if hint is not None:
        if not hint.checked or norm(hint.root) != norm(root):
            raise ReviewRequired("Invalid or foreign layout hint")
        add(hint.kind)
    if state_files:
        raise ReviewRequired("File occupies a state-directory name")
    states = set(states)
    if states == {MODERN, LEGACY}:
        raise ReviewRequired("Both state layouts have records")
    if owned:
        for state in states:
            add(state)
        for evidence in recoveries:
            if (evidence.checked and norm(evidence.root) == norm(root)
                    and norm(evidence.account) == norm(evidence.folder_account)):
                add(evidence.kind)
    if selected is None:
        named_modern = ntpath.basename(norm(root)) == "mailbackup"
        if owned and named_modern:
            raise ReviewRequired("Ambiguous historical MailBackup root")
        selected = MODERN if named_modern and not owned else LEGACY
    if any(state != selected for state in states):
        raise ReviewRequired("Populated namespace differs from selected layout")
    return selected


def paths(root, kind):
    return (ntpath.join(root, "State" if kind == MODERN else "_MailBackup_Data"),
            ntpath.join(ACCOUNT, "_Recovery" if kind == MODERN else "_MailBackup_Recovery"))


def procedure(filename, name):
    source = (ROOT / filename).read_text(encoding="ascii")
    match = re.search(rf"(?ms)^(?:Private|Public) (?:Sub|Function) {name}\b.*?^End (?:Sub|Function)\s*$", source)
    if not match:
        raise AssertionError(f"Missing source procedure {filename}::{name}")
    return "\n".join(line for line in match.group().splitlines()
                     if not line.lstrip().startswith("'"))


class LayoutBindingModels(unittest.TestCase):
    def test_fresh_mailbackup_uses_new_namespace(self):
        self.assertEqual(decide(MAIL_ROOT), MODERN)

    def test_fresh_root_matching_is_case_insensitive(self):
        self.assertEqual(decide(MAIL_ROOT.upper()), MODERN)

    def test_fresh_nonstandard_name_keeps_prior_default(self):
        self.assertEqual(decide(CUSTOM_ROOT), LEGACY)

    def test_old_v2_custom_mailbackup_is_legacy(self):
        self.assertEqual(decide(MAIL_ROOT, owner=OLD[0], states=[LEGACY]), LEGACY)

    def test_old_v33_custom_mailbackup_is_legacy(self):
        self.assertEqual(decide(MAIL_ROOT, owner=OLD[1], states=[LEGACY]), LEGACY)

    def test_old_v34_custom_mailbackup_is_legacy(self):
        self.assertEqual(decide(MAIL_ROOT, owner=OLD[2], states=[LEGACY]), LEGACY)

    def test_old_owner_preserves_legacy_when_operating_records_deleted(self):
        for value in OLD:
            with self.subTest(owner=value):
                self.assertEqual(decide(MAIL_ROOT, owner=value), LEGACY)

    def test_explicit_legacy_owner_beats_modern_basename(self):
        self.assertEqual(decide(MAIL_ROOT, owner=NEW_LEGACY), LEGACY)

    def test_explicit_modern_owner_survives_state_deletion(self):
        self.assertEqual(decide(MAIL_ROOT, owner=NEW_MODERN), MODERN)

    def test_generic_owner_with_legacy_state_is_legacy(self):
        self.assertEqual(decide(MAIL_ROOT, owner=GENERIC, states=[LEGACY]), LEGACY)

    def test_generic_owner_with_modern_state_is_modern(self):
        self.assertEqual(decide(MAIL_ROOT, owner=GENERIC, states=[MODERN]), MODERN)

    def test_generic_mailbackup_without_layout_proof_requires_review(self):
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, owner=GENERIC)

    def test_generic_safev2_retains_unambiguous_historical_namespace(self):
        self.assertEqual(decide(OLD_ROOT, owner=GENERIC), LEGACY)

    def test_generic_custom_nonmailbackup_retains_historical_namespace(self):
        self.assertEqual(decide(CUSTOM_ROOT, owner=GENERIC), LEGACY)

    def test_recovered_owner_without_layout_cannot_guess_mailbackup(self):
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, ownership_proven=True)

    def test_root_bound_registry_copy_recovers_legacy_layout(self):
        self.assertEqual(decide(MAIL_ROOT, ownership_proven=True,
                                hint=Hint(MAIL_ROOT, LEGACY)), LEGACY)

    def test_foreign_registry_root_is_rejected(self):
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, hint=Hint(OLD_ROOT, LEGACY))

    def test_bad_registry_checksum_is_rejected(self):
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, hint=Hint(MAIL_ROOT, LEGACY, False))

    def test_unknown_registry_layout_is_rejected(self):
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, hint=Hint(MAIL_ROOT, "UNKNOWN"))

    def test_unknown_owner_is_rejected(self):
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, owner="UnrelatedApp")

    def test_owner_and_registry_disagreement_is_rejected(self):
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, owner=NEW_MODERN, hint=Hint(MAIL_ROOT, LEGACY))

    def test_old_owner_and_modern_state_disagreement_is_rejected(self):
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, owner=OLD[2], states=[MODERN])

    def test_explicit_owner_and_other_state_disagreement_is_rejected(self):
        for owner, opposite in ((NEW_MODERN, LEGACY), (NEW_LEGACY, MODERN)):
            with self.subTest(owner=owner), self.assertRaises(ReviewRequired):
                decide(MAIL_ROOT, owner=owner, states=[opposite])

    def test_two_populated_state_names_are_always_rejected(self):
        for owner in (None, GENERIC, NEW_MODERN, NEW_LEGACY):
            with self.subTest(owner=owner), self.assertRaises(ReviewRequired):
                decide(MAIL_ROOT, owner=owner, states=[MODERN, LEGACY])

    def test_file_occupying_either_state_name_is_rejected(self):
        for kind in (MODERN, LEGACY):
            with self.subTest(kind=kind), self.assertRaises(ReviewRequired):
                decide(MAIL_ROOT, state_files=[kind])

    def test_exact_recovery_marker_recovers_legacy_namespace(self):
        recovery = Recovery(MAIL_ROOT, ACCOUNT, ACCOUNT, LEGACY)
        self.assertEqual(decide(MAIL_ROOT, ownership_proven=True,
                                recoveries=[recovery]), LEGACY)

    def test_exact_recovery_marker_recovers_modern_namespace(self):
        recovery = Recovery(MAIL_ROOT, ACCOUNT, ACCOUNT, MODERN)
        self.assertEqual(decide(MAIL_ROOT, owner=GENERIC,
                                recoveries=[recovery]), MODERN)

    def test_unrelated_sibling_recovery_names_do_not_conflict(self):
        unrelated = [Recovery(OLD_ROOT, ACCOUNT, ACCOUNT, k) for k in (MODERN, LEGACY)]
        self.assertEqual(decide(MAIL_ROOT, owner=NEW_MODERN,
                                recoveries=unrelated), MODERN)

    def test_moved_recovery_marker_is_not_layout_authority(self):
        bad = Recovery(MAIL_ROOT, ACCOUNT, r"C:\Archive\Different", LEGACY)
        self.assertEqual(decide(MAIL_ROOT, owner=NEW_MODERN,
                                recoveries=[bad]), MODERN)

    def test_bad_sibling_recovery_checksum_is_ignored(self):
        bad = Recovery(MAIL_ROOT, ACCOUNT, ACCOUNT, LEGACY, False)
        self.assertEqual(decide(MAIL_ROOT, owner=NEW_MODERN,
                                recoveries=[bad]), MODERN)

    def test_two_exact_recovery_names_conflict(self):
        both = [Recovery(MAIL_ROOT, ACCOUNT, ACCOUNT, k) for k in (MODERN, LEGACY)]
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, owner=GENERIC, recoveries=both)

    def test_state_and_recovery_disagreement_is_rejected(self):
        recovery = Recovery(MAIL_ROOT, ACCOUNT, ACCOUNT, LEGACY)
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, owner=GENERIC, states=[MODERN], recoveries=[recovery])

    def test_each_surviving_copy_retains_layout_after_upgrade(self):
        for kind, owner in ((MODERN, NEW_MODERN), (LEGACY, NEW_LEGACY)):
            for survives in itertools.product((False, True), repeat=4):
                if not any(survives):
                    continue
                with self.subTest(kind=kind, survives=survives):
                    self.assertEqual(decide(
                        MAIL_ROOT, owner=owner if survives[0] else None,
                        hint=Hint(MAIL_ROOT, kind) if survives[1] else None,
                        states=[kind] if survives[2] else [],
                        recoveries=[Recovery(MAIL_ROOT, ACCOUNT, ACCOUNT, kind)] if survives[3] else [],
                        ownership_proven=True), kind)

    def test_loss_of_every_layout_copy_with_old_binding_requires_review(self):
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, ownership_proven=True)

    def test_all_evidence_gone_is_new_observation_not_proof_of_history(self):
        self.assertEqual(decide(MAIL_ROOT), MODERN)

    def test_layout_hint_alone_selects_names_without_adopting_state_evidence(self):
        recovery = Recovery(MAIL_ROOT, ACCOUNT, ACCOUNT, MODERN)
        self.assertEqual(decide(MAIL_ROOT, hint=Hint(MAIL_ROOT, LEGACY),
                                recoveries=[recovery], ownership_proven=False), LEGACY)

    def test_same_run_cache_preserves_choice_after_evidence_disappears(self):
        selected = decide(MAIL_ROOT, owner=OLD[2])
        self.assertEqual(decide(MAIL_ROOT, cache=selected), LEGACY)

    def test_same_run_refresh_rejects_new_conflicting_evidence(self):
        with self.assertRaises(ReviewRequired):
            decide(MAIL_ROOT, owner=NEW_MODERN, cache=LEGACY)

    def test_old_custom_mailbackup_paths_never_move(self):
        selected = decide(MAIL_ROOT, owner=OLD[2], states=[LEGACY])
        self.assertEqual(paths(MAIL_ROOT, selected),
                         (MAIL_ROOT + r"\_MailBackup_Data", ACCOUNT + r"\_MailBackup_Recovery"))


class LayoutBindingSourceContracts(unittest.TestCase):
    def test_current_layout_owner_constants_are_distinct(self):
        source = (ROOT / "mb_operating_declarations.basfrag").read_text(encoding="ascii")
        self.assertIn(f'MB_OWNER_LAYOUT_MODERN As String = "{NEW_MODERN}"', source)
        self.assertIn(f'MB_OWNER_LAYOUT_LEGACY As String = "{NEW_LEGACY}"', source)

    def test_layout_setting_checks_framing_root_and_known_kind(self):
        body = procedure("mb_layout_binding.basfrag", "MBS_LayoutSettingValue")
        for required in ('MB_ReadCheckedLine(line, 3)', '"MBLAYOUT1"',
                         'MBH_NormalizePath(rootPath)', 'MBS_LAYOUT_MODERN', 'MBS_LAYOUT_LEGACY'):
            self.assertIn(required, body)

    def test_old_owner_is_explicit_legacy_evidence(self):
        body = procedure("mb_layout_binding.basfrag", "MBS_LayoutOwnerKind")
        self.assertIn('Case MB_OWNER_LAYOUT_LEGACY, MB_OWNER_V34, MB_OWNER_PREVIOUS, MB_OWNER_LEGACY', body)

    def test_root_name_is_not_the_first_layout_decision(self):
        body = procedure("mb_layout_binding.basfrag", "MBS_ResolveLayout")
        self.assertLess(body.index("MBS_LayoutOwnerKind(value)"), body.index("defaultModern ="))
        self.assertLess(body.index("MBS_LayoutSettingValue(saved, rootPath)"), body.index("defaultModern ="))
        self.assertIn('If owned And defaultModern Then Err.Raise MB_ERR + 1233', body)

    def test_namespace_accessors_use_cached_binding(self):
        body = procedure("mb_operating.basfrag", "MBS_IsModernLayout")
        self.assertIn("MBS_LayoutKind(managedRoot)", body)
        cache = procedure("mb_layout_binding.basfrag", "MBS_LayoutKind")
        self.assertLess(cache.index("gLayoutRoots.Exists"), cache.index("CreateObject"))

    def test_preflight_and_initialize_refresh_before_writes(self):
        gate = procedure("mb_operating.basfrag", "MBS_ValidateLayout")
        self.assertIn("MBS_ResolveLayout(managedRoot, fso, True)", gate)
        body = procedure("mb_paths.basfrag", "MBP_TryBackupBase")
        self.assertLess(body.index("MBS_ValidateLayout managed, fso"), body.index("MBP_ProbeWritable"))
        body = procedure("mb_operating.basfrag", "MBS_Initialize")
        self.assertLess(body.index("MBS_ValidateLayout gRoot, gFSO"), body.index("MBH_EnsureDirectory gStateRoot"))

    def test_cached_choice_participates_in_refresh_conflict_checks(self):
        body = procedure("mb_layout_binding.basfrag", "MBS_ResolveLayout")
        self.assertLess(body.index("selected = CStr(gLayoutRoots(rootPath))"), body.index("MBS_LayoutOwnerKind(value)"))
        self.assertIn("If Not refresh Then MBS_ResolveLayout = selected: Exit Function", body)

    def test_reset_drops_only_in_memory_layout_cache(self):
        body = procedure("mb_operating.basfrag", "MBS_Reset")
        self.assertIn('Set gLayoutRoots = CreateObject("Scripting.Dictionary")', body)
        for forbidden in ("DeleteSetting", "DeleteFile", "DeleteFolder"):
            self.assertNotIn(forbidden, body)

    def test_layout_hint_is_not_root_ownership_authority(self):
        body = procedure("mb_operating.basfrag", "MBS_CanRecoverOwner")
        self.assertNotIn("MBS_LayoutSetting", body)
        self.assertNotIn("OperatingLayout_", body)
        resolver = procedure("mb_layout_binding.basfrag", "MBS_ResolveLayout")
        self.assertIn("owned = MBS_CanRecoverOwner(rootPath, fso)", resolver)

    def test_recovery_hint_requires_exact_root_and_account_binding(self):
        body = procedure("mb_layout_binding.basfrag", "MBS_LayoutRecoveryAccount")
        self.assertIn('MBS_OwnerEvidenceFile(marker, rootPath, "RECOVERY", fso, accountPath)', body)
        self.assertNotIn("MBR_ExistingRecoveryPath", body)

    def test_resolver_does_not_mutate_state_or_settings(self):
        body = procedure("mb_layout_binding.basfrag", "MBS_ResolveLayout")
        for forbidden in ("SaveSetting", "MBS_WriteAtomic", "CreateTextFile", "MoveFile", "MoveFolder", "DeleteFile", "DeleteFolder", "SaveAs"):
            self.assertNotIn(forbidden, body)

    def test_remember_checks_real_owner_before_persisting_hint(self):
        body = procedure("mb_layout_binding.basfrag", "MBS_RememberLayout")
        self.assertLess(body.index("MBP_ReadOneLine(ownerPath, gFSO)"), body.index("MBP_SaveSettingText"))
        self.assertIn('MBH_HashText(payload)', body)

    def test_chooser_accepts_explicit_layout_owner_tags(self):
        body = procedure("mb_paths.basfrag", "MBP_ExactOwner")
        self.assertIn("MB_OWNER_LAYOUT_MODERN", body)
        self.assertIn("MB_OWNER_LAYOUT_LEGACY", body)

    def test_unowned_nonempty_root_is_rejected_before_layout_resolution(self):
        body = procedure("mb_paths.basfrag", "MBP_TryBackupBase")
        self.assertIn("If Not MBP_ExactOwner(managed, fso) Then MBP_CheckUnownedEmpty managed, fso", body)
        self.assertLess(body.index("MBP_CheckUnownedEmpty managed, fso"),
                        body.index("MBS_ValidateLayout managed, fso"))
        owner = procedure("mb_paths.basfrag", "MBP_ExactOwner")
        self.assertNotIn("MBS_LayoutSetting", owner)

    def test_owner_upgrade_selects_layout_before_preserving_previous_marker(self):
        body = procedure("mb_main.basfrag", "MB_CheckOwnership")
        self.assertLess(body.index("MBS_ValidateLayout rootPath, gFSO"),
                        body.index("expectedOwner = MBS_LayoutOwner(rootPath)"))
        self.assertLess(body.index("expectedOwner = MBS_LayoutOwner(rootPath)"),
                        body.index('"owner-previous-"'))
        upgrade = body.index("If value <> expectedOwner & vbCrLf Then")
        self.assertLess(body.index('"owner-previous-"', upgrade),
                        body.index("MBS_WriteAtomic p, expectedOwner & vbCrLf", upgrade))

    def test_each_owner_write_or_reuse_remembers_same_checked_layout(self):
        body = procedure("mb_main.basfrag", "MB_CheckOwnership")
        self.assertEqual(body.count("MBS_RememberLayout rootPath"), 3)
        self.assertEqual(body.count("MBS_WriteAtomic p, expectedOwner & vbCrLf"), 2)
        self.assertIn("ts.WriteLine expectedOwner", body)
        self.assertNotIn("MBS_WriteAtomic p, MB_OWNER", body)
        for marker in ("MB_OWNER_LAYOUT_MODERN", "MB_OWNER_LAYOUT_LEGACY"):
            self.assertIn(marker, body)

    def test_legacy_discovery_recognizes_both_layout_owner_tags(self):
        body = procedure("mb_legacy.basfrag", "MBL_HasOwner")
        self.assertIn("MB_OWNER_LAYOUT_MODERN", body)
        self.assertIn("MB_OWNER_LAYOUT_LEGACY", body)

    def test_recovery_path_uses_same_namespace_decision(self):
        body = procedure("mb_recovery.basfrag", "MBR_ExpectedRecoveryPath")
        self.assertIn("MBS_IsModernLayout(managedRoot)", body)


if __name__ == "__main__":
    unittest.main(verbosity=2)
