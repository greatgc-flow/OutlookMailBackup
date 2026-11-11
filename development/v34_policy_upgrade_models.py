"""V3 proof-policy upgrade and compatibility models; not Outlook/VBA execution.

Focused regressions cover old proof labels behind a new checkpoint, simultaneous
checkpoint/snapshot loss, restart progress, orphan history, and guarded validation
upgrades that never re-export existing MSG bytes. Source contracts check the
production branches supporting each invariant.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import json
import re
import unittest

HERE = Path(__file__).resolve().parent
FULL3 = "REOPENED_CONTENT_SHA256_V3"
FULL2 = "REOPENED_CONTENT_SHA256_V2"
BASIC3 = "REOPENED_METADATA_SHA256_V3"
BASIC2 = "REOPENED_METADATA_SHA256_V2"
LEGACY3 = "LEGACY_COMPARE_SHA256_V3"
LEGACY2 = "LEGACY_COMPARE_SHA256_V2"
LEGACY1 = "LEGACY_COMPARE_SHA256_V1"
POLICY3 = "AUTO_INVENTORY_UTC_V2_FIDELITY_V3"
POLICY2 = "AUTO_INVENTORY_UTC_V2"


def src(name):
    return (HERE / name).read_text(encoding="ascii")


def proc(text, name):
    match = re.search(r"(?m)^(?:Private|Public) (Sub|Function) " + re.escape(name) + r"\b", text)
    if not match:
        raise AssertionError(name)
    end = re.search(r"(?m)^End " + match[1], text[match.end():])
    return text[match.start():match.end() + end.end()]


@dataclass(frozen=True)
class Proof:
    source: str
    key: str
    label: str
    digest: str = "unchanged-file-hash"


class Inventory:
    def __init__(self, rows=()):
        self.heads = {}
        self.index = {}
        self.known = set()
        for row in rows:
            self.commit(row)

    def commit(self, row):
        self.heads[row.source] = row.key
        self.index[row.key] = row
        self.known.add(row.source)

    def needs_policy(self, source):
        # Unknown identities are always queued by the inventory's first branch.
        if source not in self.heads:
            return False
        key = self.heads[source]
        if key not in self.index:
            return True
        expected = LEGACY3 if key.startswith("L|") else FULL3
        return self.index[key].label != expected

    def queue(self, source, *, full=False, changed=False, recent=False, audit=False, pending=False):
        return (pending or source not in self.known or changed or self.needs_policy(source)
                or full or audit or recent)

    def require_content(self, source, *, full=False, audit=False, pending=False):
        return full or self.needs_policy(source) or pending or audit


def full_scan(checkpoint, known_count, force_recovery=False):
    if force_recovery:
        return True
    if checkpoint is None:
        return known_count > 0
    return checkpoint != POLICY3


class Upgrade:
    def __init__(self, old_label):
        self.label = old_label
        self.old_bytes = b"existing MSG"
        self.events = []
        self.saves = 0
        self.committed = False
        self.closed = False

    def validate(self, *, complete=True, hash_matches=True, source_stable=True, compare_error=False):
        self.events.append("GUARD")
        try:
            self.events.append("COMPARE")
            if compare_error:
                raise ValueError("comparison failed")
            self.events.append("CHECK_EXPECTED_HASH")
            if not hash_matches:
                raise ValueError("file changed")
            self.events.append("CHECK_SOURCE_REVISION")
            if not source_stable:
                raise ValueError("source changed")
            self.label = FULL3 if complete else BASIC3
            self.events.append("APPEND_PROOF")
            self.committed = True
        finally:
            self.events.append("CLOSE_GUARD")
            self.closed = True


class PolicyBehaviorModels(unittest.TestCase):
    def test_v2_new_and_v1_v2_legacy_labels_are_readable_but_not_current_policy(self):
        rows = [Proof("new", "M|n", FULL2), Proof("basic", "M|b", BASIC2),
                Proof("legacy1", "L|a", LEGACY1), Proof("legacy2", "L|b", LEGACY2)]
        state = Inventory(rows)
        for row in rows:
            with self.subTest(row=row):
                self.assertTrue(state.needs_policy(row.source))
                self.assertTrue(state.queue(row.source))
                self.assertTrue(state.require_content(row.source))

    def test_current_full_v3_labels_retain_incremental_fast_path(self):
        state = Inventory([Proof("new", "M|a", FULL3), Proof("legacy", "L|b", LEGACY3)])
        for key in state.known:
            self.assertFalse(state.needs_policy(key))
            self.assertFalse(state.queue(key))
            self.assertFalse(state.require_content(key))

    def test_current_basic_proof_cannot_hide_behind_v3_checkpoint(self):
        state = Inventory([Proof("limited", "M|a", BASIC3)])
        self.assertFalse(full_scan(POLICY3, len(state.known)))
        self.assertTrue(state.queue("limited"))
        self.assertTrue(state.require_content("limited"))

    def test_old_current_proof_is_queued_even_with_new_checkpoint_and_old_unchanged_mail(self):
        state = Inventory([Proof("old-source", "M|a", FULL2)])
        self.assertFalse(full_scan(POLICY3, len(state.known)))
        self.assertTrue(state.queue("old-source", full=False, changed=False, recent=False, audit=False))

    def test_lost_checkpoint_and_snapshot_with_managed_history_forces_full(self):
        # No legacy catalog and no restoration trigger are required.
        state = Inventory([Proof("managed", "M|a", FULL2)])
        self.assertTrue(full_scan(None, len(state.known), force_recovery=False))
        self.assertTrue(state.queue("managed", full=True))

    def test_missing_checkpoint_with_no_history_still_queues_every_unknown_source(self):
        state = Inventory()
        self.assertFalse(full_scan(None, 0))
        self.assertTrue(state.queue("new-old-dated-mail", recent=False))

    def test_previous_checkpoint_policy_always_forces_full(self):
        self.assertTrue(full_scan(POLICY2, 1))
        self.assertTrue(full_scan(POLICY2, 0))

    def test_successful_upgrade_updates_latest_head_and_exits_policy_queue(self):
        state = Inventory([Proof("source", "M|a", FULL2)])
        self.assertTrue(state.needs_policy("source"))
        state.commit(Proof("source", "M|a", FULL3))
        self.assertFalse(state.needs_policy("source"))
        self.assertFalse(state.queue("source"))

    def test_latest_source_revision_controls_policy_not_old_historical_revision(self):
        state = Inventory([Proof("source", "M|old", FULL2), Proof("source", "M|new", FULL3)])
        self.assertIn("M|old", state.index)
        self.assertFalse(state.needs_policy("source"))

    def test_orphan_old_proofs_do_not_force_live_v3_mail_into_perpetual_full_scan(self):
        state = Inventory([Proof("live", "M|live", FULL3)] +
                          [Proof(f"orphan-{i}", f"M|{i}", FULL2) for i in range(1000)])
        self.assertFalse(full_scan(POLICY3, len(state.known)))
        self.assertEqual([s for s in ["live"] if state.queue(s)], [])

    def test_reappearing_orphan_with_old_label_is_compared_when_inventory_sees_it(self):
        state = Inventory([Proof("returned", "M|old", FULL2)])
        self.assertTrue(state.queue("returned", recent=False))
        self.assertTrue(state.require_content("returned"))

    def test_interrupted_policy_upgrade_retries_only_remaining_old_proofs_after_restart(self):
        rows = [Proof(str(i), f"M|{i}", FULL2) for i in range(10)]
        state = Inventory(rows)
        upgraded = [Proof(str(i), f"M|{i}", FULL3) for i in range(4)]
        restarted = Inventory(rows + upgraded)
        self.assertEqual([str(i) for i in range(10) if restarted.queue(str(i))], [str(i) for i in range(4, 10)])

    def test_dangling_known_head_is_review_candidate_not_silent_skip(self):
        state = Inventory([Proof("source", "M|a", FULL3)])
        del state.index["M|a"]
        self.assertTrue(state.needs_policy("source"))
        self.assertTrue(state.queue("source"))

    def test_validation_upgrade_never_reexports_or_changes_msg_bytes(self):
        upgrade = Upgrade(FULL2)
        upgrade.validate()
        self.assertEqual(upgrade.saves, 0)
        self.assertEqual(upgrade.old_bytes, b"existing MSG")
        self.assertEqual(upgrade.label, FULL3)
        self.assertEqual(upgrade.events, ["GUARD", "COMPARE", "CHECK_EXPECTED_HASH", "CHECK_SOURCE_REVISION", "APPEND_PROOF", "CLOSE_GUARD"])

    def test_file_race_source_race_or_comparison_error_cannot_publish_v3_label(self):
        for failure in ({"hash_matches": False}, {"source_stable": False}, {"compare_error": True}):
            with self.subTest(failure=failure):
                upgrade = Upgrade(FULL2)
                with self.assertRaises(ValueError):
                    upgrade.validate(**failure)
                self.assertEqual(upgrade.label, FULL2)
                self.assertFalse(upgrade.committed)
                self.assertTrue(upgrade.closed)
                self.assertEqual(upgrade.saves, 0)

    def test_limited_comparison_is_labelled_basic_and_remains_retryable(self):
        upgrade = Upgrade(FULL2)
        upgrade.validate(complete=False)
        self.assertEqual(upgrade.label, BASIC3)
        state = Inventory([Proof("source", "M|a", upgrade.label)])
        self.assertTrue(state.queue("source"))


class PolicySourceContracts(unittest.TestCase):
    def test_old_labels_are_explicitly_accepted_by_existing_and_legacy_loaders(self):
        main = proc(src("mb_main.basfrag"), "MB_ExistingValid")
        for token in ('"REOPENED_SHA256"', "MB_NEW_CONTENT_VALIDATION_V2", "MB_NEW_BASIC_VALIDATION_V2"):
            self.assertIn(token, main)
        legacy = proc(src("mb_reuse.basfrag"), "MB_CacheLegacy")
        self.assertIn("MB_LEGACY_VALIDATION_V2", legacy)
        self.assertIn("MB_LEGACY_VALIDATION_V1", legacy)
        cache = proc(src("mb_layout_integration.basfrag"), "MBX_CacheRecord")
        self.assertIn("MB_NEW_CONTENT_VALIDATION_V2", cache)

    def test_upgrade_guards_compare_expected_hash_and_source_before_append(self):
        upgrade = proc(src("mb_main.basfrag"), "MB_UpgradeMsgValidation")
        order = ["guard = MBH_GuardFile(path)", "MB_CompareMsg mail, path, complete", "MBH_HashFile(path) <> CStr(fields(2))",
                 "MB_SourceStable(mail, q, revision)", "MB_RecordCommit", "MBH_CloseGuard guard"]
        self.assertEqual([upgrade.index(x) for x in order], sorted(upgrade.index(x) for x in order))
        self.assertNotRegex(upgrade, r"(?i)\bmail\.SaveAs\b")
        self.assertIn("validation <> MB_NEW_CONTENT_VALIDATION Or MBW_RequireContentCheck", upgrade)
        self.assertIn("Else validation = MB_NEW_BASIC_VALIDATION", upgrade)

    def test_policy_helper_reads_current_head_in_o1_without_scanning_orphan_history(self):
        helper = proc(src("mb_revision.basfrag"), "MBV_RequiresCurrentPolicy")
        self.assertIn("gRevisionHeads(sourceKey)", helper)
        self.assertIn("fields = MB_IndexFields(key)", helper)
        self.assertIn("CStr(fields(3)) <> MB_LEGACY_VALIDATION", helper)
        self.assertIn("CStr(fields(3)) <> MB_NEW_CONTENT_VALIDATION", helper)
        self.assertNotIn("For Each", helper)
        self.assertNotIn("gWindowFull =", helper)
        self.assertIn("If Not gIndex.Exists(key) Then MBV_RequiresCurrentPolicy = True", helper)

    def test_inventory_and_full_compare_both_apply_current_policy_gate(self):
        window = src("mb_window.basfrag")
        queue = proc(window, "MBW_ShouldQueue")
        self.assertLess(queue.index("MBV_RequiresCurrentPolicy(sourceKey)"), queue.index("ElseIf gWindowFull"))
        self.assertLess(queue.index("MBV_RequiresCurrentPolicy(sourceKey)"), queue.index("ElseIf (effectiveDate"))
        self.assertIn("MBV_RequiresCurrentPolicy(sourceKey)", proc(window, "MBW_RequireContentCheck"))
        self.assertIn("gWindowPolicyQueued = 0", proc(window, "MBW_Reset"))

    def test_no_checkpoint_with_existing_history_forces_full_independently_of_legacy_count(self):
        prepare = proc(src("mb_window.basfrag"), "MBW_Prepare")
        branch = prepare[prepare.index("If gWindowKnown.Count > 0 Then"):]
        self.assertLess(branch.index("gWindowFull = True"), branch.index("ElseIf"))
        self.assertNotIn("gLegacyFileCount", prepare)
        self.assertIn("FIDELITY_V3", src("mb_window_declarations.basfrag"))

    def test_legacy_fastpath_requires_current_v3_label_and_no_forced_comparison(self):
        legacy = proc(src("mb_reuse.basfrag"), "MB_ReuseLegacy")
        self.assertIn('key = "L|" & revision And CStr(fields(3)) = MB_LEGACY_VALIDATION', legacy)
        self.assertIn("Not MBW_RequireContentCheck", legacy)
        self.assertIn("guard = MBH_GuardFile(path)", legacy)
        self.assertIn("MBL_VerifyCandidate", legacy)

    def test_old_root_skip_boundaries_accept_previous_v33_owner(self):
        has_owner = proc(src("mb_legacy.basfrag"), "MBL_HasOwner")
        self.assertIn("MB_OWNER_PREVIOUS", has_owner)
        self.assertIn("MBL_HasOwner(current)", proc(src("mb_move.basfrag"), "MBM_BuildPlan"))
        self.assertIn("MBL_HasOwner(current)", proc(src("mb_legacy.basfrag"), "MBL_BuildIndex"))

    def test_head_reset_is_shared_with_normal_and_repaired_index_resets(self):
        identity = proc(src("mb_identity.basfrag"), "MBI_Reset")
        self.assertIn("MBV_Reset", identity)
        self.assertIn("MBI_Reset", proc(src("mb_main.basfrag"), "MB_Reset"))
        self.assertIn("MBI_Reset", proc(src("mb_operating.basfrag"), "MBS_ClearLoadedIndex"))
        self.assertIn("MBV_CacheRecord f", proc(src("mb_layout_integration.basfrag"), "MBX_CacheRecord"))

    def test_limited_content_stops_coverage_commit(self):
        self.assertIn("gContentReview <> 0", proc(src("mb_window.basfrag"), "MBW_CanCommit"))
        limited_branches = re.findall(
            r"If validation = MB_NEW_BASIC_VALIDATION Then\s+"
            r"gContentReview = gContentReview \+ 1\s+"
            r'MBG_LogReview "SPECIAL_ATTACHMENT_CONTENT_REVIEW"',
            proc(src("mb_main.basfrag"), "MB_ProcessOne"))
        # Both reused BASIC proof and a newly saved limited MSG must block
        # successful coverage and explain their unresolved fidelity limit.
        self.assertEqual(len(limited_branches), 2)


if __name__ == "__main__":
    result = unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromModule(__import__(__name__)))
    report = {"kind": __doc__, "tests": result.testsRun, "failures": len(result.failures), "errors": len(result.errors),
              "not_run": ["Office VBA compiler", "Outlook/MAPI runtime", "Win32 file guard execution"]}
    (HERE / "vba_parser_review" / "v34_policy_upgrade_models.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))
    raise SystemExit(not result.wasSuccessful())
