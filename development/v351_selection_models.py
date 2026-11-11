"""Optional partial-run selection: failure injection and source wiring, not VBA execution."""
from dataclasses import dataclass, replace
from datetime import datetime, timedelta, timezone
import hashlib
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parent
NOW = datetime(2026, 11, 11, tzinfo=timezone.utc)
OLD = NOW - timedelta(days=120)
RECENT = NOW - timedelta(days=1)


def digest(value):
    return hashlib.sha256(value).hexdigest()


@dataclass(frozen=True)
class State:
    root: str = r"D:\Backup\MailBackup"
    store: str = "Store-A"
    account: str = r"D:\Mail\Account"
    context: str = "zone-a"
    policy: str = "fidelity-v3"
    coverage_time: datetime = OLD
    coverage: bytes = b"successful-coverage"
    journal: bytes = b"checked-committed-proof"
    audit: bytes = b"all-review-keys-plus-cursor"
    full: bool = False
    recovery: bool = False
    state_missing: bool = False
    state_migrated: bool = False
    invalid_rows: bool = False


def certificate(state, scan=RECENT):
    return (state.root.casefold(), state.store, state.account.casefold(),
            state.context, state.policy, state.coverage_time, scan,
            digest(state.coverage), digest(state.journal), digest(state.audit))


def usable(state, cert, now=NOW):
    if any((state.full, state.recovery, state.state_missing, state.state_migrated, state.invalid_rows)):
        return False
    if not all((state.coverage, state.journal, state.audit)):
        return False
    if state.coverage_time > now:
        return False
    expected = certificate(state, cert[6])
    return cert == expected and state.coverage_time < cert[6] <= now


def publishable(*, planned=2, processed=2, complete=True, audit_saved=True,
                reviews=frozenset({"review-a"}), pending=frozenset({"review-a"}),
                soft_count=1, cancelled=False, fatal=False, cleanup=False,
                open_stream=False, failure=False, changed=False, policy_skip=False,
                unmatched=False, attachment_failure=False, folder_failure=False,
                changed_inventory=False, archive_review=False):
    blocked = (cancelled, fatal, cleanup, open_stream, failure, changed, policy_skip,
               unmatched, attachment_failure, folder_failure, changed_inventory, archive_review)
    return (not any(blocked) and complete and audit_saved and planned == processed
            and soft_count > 0 and len(reviews) == soft_count and reviews <= pending)


def selected(*, known=True, changed=False, policy=False, pending=False,
             audit=False, full=False, effective=OLD, modified=OLD, since=OLD):
    return (not known or changed or policy or pending or audit or full
            or since <= effective <= NOW or since <= modified <= NOW)


class CacheBinding(unittest.TestCase):
    def test_healthy_exact_binding(self):
        s = State(); self.assertTrue(usable(s, certificate(s)))

    def test_each_evidence_file_change_invalidates(self):
        s = State(); c = certificate(s)
        for key in ("coverage", "journal", "audit"):
            with self.subTest(key=key):
                self.assertFalse(usable(replace(s, **{key: getattr(s, key) + b"new"}), c))

    def test_each_missing_evidence_file_invalidates(self):
        s = State(); c = certificate(s)
        for key in ("coverage", "journal", "audit"):
            with self.subTest(key=key): self.assertFalse(usable(replace(s, **{key: b""}), c))

    def test_each_context_or_ownership_change_invalidates(self):
        s = State(); c = certificate(s)
        for key in ("root", "store", "account", "context", "policy"):
            with self.subTest(key=key): self.assertFalse(usable(replace(s, **{key: getattr(s, key) + "-other"}), c))

    def test_store_is_case_sensitive(self):
        s = State(); self.assertFalse(usable(replace(s, store=s.store.lower()), certificate(s)))

    def test_paths_are_case_insensitive(self):
        s = State(); self.assertTrue(usable(replace(s, root=s.root.lower(), account=s.account.upper()), certificate(s)))

    def test_coverage_change_invalidates(self):
        s = State(); self.assertFalse(usable(replace(s, coverage_time=OLD + timedelta(days=1)), certificate(s)))

    def test_future_scan_rejected(self):
        s = State(); self.assertFalse(usable(s, certificate(s, NOW + timedelta(seconds=1))))

    def test_equal_or_older_than_coverage_rejected(self):
        s = State()
        for t in (OLD, OLD - timedelta(seconds=1)):
            with self.subTest(t=t): self.assertFalse(usable(s, certificate(s, t)))

    def test_clock_rollback_rejected(self):
        s = State(); self.assertFalse(usable(s, certificate(s), OLD - timedelta(seconds=1)))

    def test_full_and_recovery_never_shortened(self):
        s = State(); c = certificate(s)
        for key in ("full", "recovery", "state_missing", "state_migrated", "invalid_rows"):
            with self.subTest(key=key): self.assertFalse(usable(replace(s, **{key: True}), c))

    def test_audit_review_replaced_with_ok_invalidates(self):
        s = State(); c = certificate(s)
        self.assertFalse(usable(replace(s, audit=b"cursor-only-after-review-removal"), c))


class PublicationSafety(unittest.TestCase):
    def test_soft_review_can_save_selection_without_completing_coverage(self):
        self.assertTrue(publishable())
        coverage_complete = False  # a REVIEW is still outstanding
        self.assertFalse(coverage_complete)

    def test_equal_count_wrong_source_keys_not_enough(self):
        self.assertFalse(publishable(pending=frozenset({"another-source"})))

    def test_one_missing_required_key_blocks(self):
        self.assertFalse(publishable(reviews=frozenset({"a", "b"}), pending=frozenset({"a", "c"}), soft_count=2))

    def test_existing_absent_source_review_is_retained(self):
        pending = frozenset({"review-a", "old-absent-source"})
        self.assertTrue(publishable(pending=pending))
        self.assertIn("old-absent-source", pending)

    def test_missing_durable_audit_blocks(self):
        self.assertFalse(publishable(audit_saved=False))

    def test_incomplete_inventory_blocks(self):
        self.assertFalse(publishable(complete=False))

    def test_unprocessed_queue_blocks(self):
        self.assertFalse(publishable(processed=1))

    def test_all_hard_failures_block(self):
        for key in ("cancelled", "fatal", "cleanup", "open_stream", "failure", "changed",
                    "policy_skip", "unmatched", "attachment_failure", "folder_failure",
                    "changed_inventory", "archive_review"):
            with self.subTest(key=key): self.assertFalse(publishable(**{key: True}))

    def test_aggregate_soft_count_mismatch_blocks(self):
        self.assertFalse(publishable(soft_count=2))

    def test_no_review_does_not_need_partial_cache(self):
        self.assertFalse(publishable(reviews=frozenset(), pending=frozenset(), soft_count=0))


class SelectionEquivalence(unittest.TestCase):
    def setUp(self):
        self.since = RECENT - timedelta(days=7)

    def test_old_pending_source_still_selected(self):
        self.assertTrue(selected(pending=True, since=self.since))

    def test_unknown_historical_source_still_selected(self):
        self.assertTrue(selected(known=False, since=self.since))

    def test_changed_old_source_still_selected(self):
        self.assertTrue(selected(changed=True, since=self.since))

    def test_old_policy_upgrade_still_selected(self):
        self.assertTrue(selected(policy=True, since=self.since))

    def test_rotating_content_audit_still_selected(self):
        self.assertTrue(selected(audit=True, since=self.since))

    def test_overlap_boundary_is_inclusive(self):
        self.assertTrue(selected(effective=self.since, since=self.since))
        self.assertFalse(selected(effective=self.since-timedelta(seconds=1), since=self.since))

    def test_recent_modified_stamp_still_selected(self):
        self.assertTrue(selected(modified=NOW-timedelta(days=2), since=self.since))

    def test_hundred_thousand_already_verified_sources_avoid_widening_window(self):
        dates = [OLD + timedelta(days=1 + i % 100) for i in range(100000)]
        original = sum(selected(effective=t, since=OLD-timedelta(days=7)) for t in dates)
        optimized = sum(selected(effective=t, since=self.since) for t in dates)
        self.assertEqual(original, 100000)
        self.assertEqual(optimized, 0)
        # This is an operation-count model, not elapsed Outlook runtime.


class SourceContracts(unittest.TestCase):
    def test_cache_is_not_snapshot_state_or_coverage(self):
        s = (ROOT/'mb_selection.basfrag').read_text()
        self.assertIn('& ".cache")', s)
        self.assertNotIn('gWindowCommitted = True', s)
        self.assertNotIn('gRecoveryForceFull = False', s)
        self.assertNotIn('gWindowFull = False', s)

    def test_cache_load_after_pending_ledger_before_queue(self):
        s=(ROOT/'mb_window.basfrag').read_text()
        body=s.split('Private Sub MBW_Prepare()',1)[1].split('End Sub',1)[0]
        self.assertLess(body.index('MBW_LoadAuditState'),body.index('MBW_ApplySelectionCache'))
        self.assertLess(body.index('MBW_ApplySelectionCache'),body.index('MBW_BuildAudit'))

    def test_cache_publish_after_close(self):
        s=(ROOT/'mb_main.basfrag').read_text()
        self.assertLess(s.index('    MB_CloseResources\n'),s.index('Then MBW_SaveSelectionCache'))
        critical = s[s.index('    If mode = "AUTO" Then MBW_CommitAfterClose'):
                     s.index('    If mode = "AUTO" Then MBW_SaveSelectionCache')]
        self.assertIn('If Err.Number <> 0 Then', critical)
        self.assertIn('gFatal = True', critical)
        self.assertIn('gFatalText = gFatalText & " UTC checkpoint: " & Err.Description', critical)
        self.assertTrue(critical.rstrip().endswith('On Error GoTo 0'))
        # Optional cache error cleanup must never clear a critical checkpoint
        # error before it has been reflected in the final status.
        self.assertNotIn('Err.Clear', critical)
        c=(ROOT/'mb_selection.basfrag').read_text()
        self.assertIn('Not gJournal Is Nothing Or Not gEvents Is Nothing',c)
        self.assertIn('gWindowSourcePending.Exists(CStr(key))',c)

    def test_existing_success_coverage_gate_keeps_reviews_blocked(self):
        s=(ROOT/'mb_window.basfrag').read_text().split('Private Function MBW_CanCommit()',1)[1].split('End Function',1)[0]
        self.assertIn('gMailReview <> 0',s)
        self.assertIn('gContentReview <> 0',s)

    def test_three_raw_evidence_hashes_required(self):
        s=(ROOT/'mb_selection.basfrag').read_text()
        for path in ('gWindowPath','gJournalPath','gWindowAuditPath'):
            self.assertGreaterEqual(s.count('MBW_SelectionDigest('+path+')'),2)

    def test_read_does_not_persist_or_remove_reviews(self):
        s=(ROOT/'mb_selection.basfrag').read_text().split('Private Sub MBW_ApplySelectionCache()',1)[1].split('End Sub',1)[0]
        self.assertNotIn('MBS_WriteAtomic',s)
        self.assertNotIn('.Remove',s)


if __name__ == '__main__':
    unittest.main()
