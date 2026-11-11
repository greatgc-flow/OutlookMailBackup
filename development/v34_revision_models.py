"""v3.4 same-source revision regression/failure models and source contracts.

These checks exercise an independent Python policy model, adversarial event order,
and the actual VBA routing source. They do NOT execute Outlook, MAPI, Win32 file
sharing or the Office VBA compiler. The archive content model represents only the
explicitly declared bounded comparator, not every MAPI property.
"""
from dataclasses import dataclass, replace
from enum import IntEnum
from hashlib import sha256
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parent


class Decision(IntEnum):
    NONE = 0
    REUSED = 1
    REVIEW = 2
    CANCELLED = 3
    CHANGED = 4


@dataclass(frozen=True)
class Message:
    source: str = 'source-A'
    store: str = 'store-A'
    revision: str = 'revision-2'
    subject: str = 'old subject'
    body: bytes = b'old body'
    metadata: str = 'read'

    @property
    def bounded(self):
        return self.subject, self.body


@dataclass(frozen=True)
class Proof:
    key: str
    source: str
    store: str
    subject: str
    body: bytes
    path: str
    expected: str
    size: int
    provider_time: int = 0
    root: str = 'account-root'

    @classmethod
    def from_message(cls, message, kind='N', path='old.msg', provider_time=0):
        # The disk byte model is purposefully distinct from the bounded content:
        # metadata revisions may point to one unchanged prior MSG serialization.
        data = b'MSG:' + message.subject.encode() + b':' + message.body
        key_kind = 'M' if kind in ('N', 'R') else kind
        return cls(key_kind + '|' + message.revision, message.source, message.store,
                   message.subject, message.body, path, sha256(data).hexdigest(),
                   len(data), provider_time), data

    @property
    def bounded(self):
        return self.subject, self.body


class Archive:
    """Small state machine specified by archival invariants, not Outlook APIs."""
    def __init__(self):
        self.records = {}
        self.head = {}
        self.files = {}
        self.events = []
        self.new_files = 0
        self.aliases = 0
        self.guarded = False

    def accept(self, proof, data, *, kind='N'):
        self.records[proof.key] = proof
        self.files[proof.path] = data
        if kind in ('M', 'N', 'L', 'R'):
            self.head[proof.source] = proof.key

    def known(self, message, *, comparison=None, stable=True, cancelled=False,
              returned=None, writer_fails=False, attempt_write=False):
        if message.source not in self.head:
            return Decision.NONE
        key = self.head[message.source]
        proof = self.records.get(key)
        if proof is None or (proof.source, proof.store) != (message.source, message.store):
            return Decision.REVIEW
        # Immutable expected bytes are captured before adversarial ByRef outputs.
        expected = (proof.path, proof.root, proof.expected, proof.size)
        if proof.path not in self.files:
            return Decision.REVIEW
        self.guarded = True
        self.events.append('guard')
        try:
            data = self.files[proof.path]
            if data is None or sha256(data).hexdigest() != proof.expected or len(data) != proof.size:
                return Decision.REVIEW
            self.events.append('verify_expected_bytes')
            if cancelled:
                return Decision.CANCELLED
            result = comparison if comparison is not None else int(proof.bounded == message.bounded)
            self.events.append('compare_content')
            if result not in (0, 1):
                return Decision.REVIEW
            outputs = returned or (proof.path, proof.root, sha256(data).hexdigest(), len(data))
            if outputs != expected or not stable:
                return Decision.REVIEW
            self.events.append('source_stable')
            if attempt_write:
                # Specified Windows sharing invariant; native implementation is
                # tested separately by the existing actual Win32 guard self-test.
                if self.guarded:
                    self.events.append('write_denied')
                else:
                    raise AssertionError('unprotected expected file')
            target = 'M|' + message.revision
            if result == 0:
                return Decision.REVIEW if target in self.records else Decision.CHANGED
            if target in self.records:
                old = self.records[target]
                if (old.path, old.expected, old.size) != (proof.path, proof.expected, proof.size):
                    return Decision.REVIEW
                self.events.append('already_current_no_self_reference')
            else:
                self.events.append('commit_alias')
                if writer_fails:
                    raise OSError('journal append failed')
                alias = replace(proof, key=target, source=message.source, store=message.store)
                self.records[target] = alias
                self.head[message.source] = target
                self.aliases += 1
            return Decision.REUSED
        finally:
            self.guarded = False
            self.events.append('release')

    def workflow(self, message, *, registration=False, journal_ok=True, catalog_ok=True,
                 mapped=True, stats_present=True, unknown_legacy=0, old_date=True,
                 legacy_exists=True, **faults):
        result = self.known(message, **faults)
        if result in (Decision.REUSED, Decision.REVIEW, Decision.CANCELLED):
            return result
        # A subject-key search still runs after a verified changed latest head.
        # This closes A -> B -> A without another physical copy of A.
        matches = [p for p in self.records.values()
                   if p.store == message.store and p.subject == message.subject
                   and p.bounded == message.bounded and p.path in self.files
                   and self.files[p.path] is not None
                   and sha256(self.files[p.path]).hexdigest() == p.expected]
        if matches:
            p = matches[0]
            alias = replace(p, key='M|' + message.revision, source=message.source)
            self.records[alias.key] = alias
            self.head[message.source] = alias.key
            self.aliases += 1
            return Decision.REUSED
        if registration:
            return 'UNMATCHED_NO_EXPORT'
        # New-version evidence does not waive global account/catalog integrity.
        if not all((journal_ok, catalog_ok, mapped, stats_present)):
            return Decision.REVIEW
        if result != Decision.CHANGED and (unknown_legacy or (legacy_exists and old_date)):
            return Decision.REVIEW
        proof, data = Proof.from_message(message, path='new-' + message.revision + '.msg')
        if proof.path in self.files:
            raise AssertionError('publication must not overwrite')
        self.accept(proof, data)
        self.new_files += 1
        return 'NEW'


def fixture(kind='N'):
    archive = Archive()
    old = Message(revision='revision-1')
    proof, data = Proof.from_message(old, kind=kind)
    archive.accept(proof, data, kind=kind)
    return archive, proof, data


def source(name):
    return (ROOT / name).read_text(encoding='ascii')


def procedure(text, name):
    start = re.search(r'(?m)^(?:Private|Public) (?:Function|Sub) ' + re.escape(name) + r'\b', text)
    if start is None:
        raise AssertionError('missing procedure ' + name)
    end = re.search(r'(?m)^End (?:Function|Sub)', text[start.end():])
    return text[start.start():start.end() + end.end()]


class RevisionScenarios(unittest.TestCase):
    def test_known_legacy_old_body_edit_exits_permanent_review(self):
        a, p, data = fixture('L')
        changed = Message(body=b'edited body')
        # Previous direct-L branch treated any explicit content mismatch as REVIEW.
        prior_result = Decision.REVIEW if changed.bounded != p.bounded else Decision.REUSED
        self.assertEqual(prior_result, Decision.REVIEW)
        self.assertEqual(a.workflow(changed), 'NEW')
        self.assertEqual(a.new_files, 1)
        self.assertEqual(a.files['old.msg'], data)
        self.assertEqual(a.records[p.key], p)

    def test_known_new_format_old_subject_edit_exits_date_cutoff(self):
        a, p, data = fixture('N')
        current = Message(subject='new full subject')
        self.assertEqual(a.workflow(current, old_date=True), 'NEW')
        self.assertEqual(a.files[p.path], data)
        self.assertIn(p.key, a.records)

    def test_subject_independent_head_has_no_search_scan(self):
        a, p, _ = fixture()
        a.records['irrelevant'] = replace(p, key='irrelevant', source='other', subject='irrelevant')
        self.assertEqual(a.known(Message(subject='changed')), Decision.CHANGED)
        self.assertEqual(a.events.count('compare_content'), 1)

    def test_metadata_only_edit_links_same_msg(self):
        a, p, _ = fixture('L')
        self.assertEqual(a.workflow(Message(metadata='unread')), Decision.REUSED)
        self.assertEqual(a.new_files, 0)
        self.assertEqual(len(a.files), 1)
        self.assertEqual(a.records['M|revision-2'].expected, p.expected)
        self.assertEqual(a.aliases, 1)

    def test_rerun_same_revision_does_not_self_reference(self):
        a, _, _ = fixture()
        message = Message()
        self.assertEqual(a.workflow(message), Decision.REUSED)
        self.assertEqual(a.workflow(message), Decision.REUSED)
        self.assertEqual(a.aliases, 1)
        self.assertIn('already_current_no_self_reference', a.events)

    def test_newer_r_head_supersedes_stale_l_source_pointer(self):
        a, old, _ = fixture('L')
        b = Message(body=b'version B')
        self.assertEqual(a.workflow(b), 'NEW')
        self.assertEqual(a.workflow(replace(b, revision='revision-3')), Decision.REUSED)
        self.assertEqual(a.head[b.source], 'M|revision-3')
        self.assertEqual(a.known(replace(b, revision='revision-4')), Decision.REUSED)
        self.assertIn(old.key, a.records)
        self.assertEqual(a.new_files, 1)

    def test_edit_then_revert_reuses_older_file(self):
        for changed_subject in (False, True):
            with self.subTest(changed_subject=changed_subject):
                a, old, data = fixture('L')
                b = Message(body=b'version B', subject='B' if changed_subject else 'old subject')
                self.assertEqual(a.workflow(b), 'NEW')
                reverted = Message(revision='revision-3')
                self.assertEqual(a.workflow(reverted), Decision.REUSED)
                self.assertEqual(a.new_files, 1)
                self.assertEqual(a.records['M|revision-3'].path, old.path)
                self.assertEqual(a.files[old.path], data)

    def test_restored_entryid_reuses_same_store_old_file(self):
        a, old, _ = fixture('L')
        restored = Message(source='restored-entry')
        self.assertEqual(a.known(restored), Decision.NONE)
        self.assertEqual(a.workflow(restored), Decision.REUSED)
        self.assertEqual(a.new_files, 0)
        self.assertEqual(a.records['M|revision-2'].path, old.path)

    def test_unknown_old_mail_cannot_bypass_duplicate_cutoff(self):
        a, _, _ = fixture()
        self.assertEqual(a.workflow(Message(source='unknown', body=b'unknown body')), Decision.REVIEW)
        self.assertEqual(a.new_files, 0)

    def test_unknown_recent_mail_can_be_exported(self):
        a, _, _ = fixture()
        self.assertEqual(a.workflow(Message(source='unknown', body=b'new body'), old_date=False), 'NEW')
        self.assertEqual(a.new_files, 1)

    def test_registration_never_saves_new_file(self):
        for kind in ('L', 'N'):
            for msg in (Message(), Message(body=b'changed'), Message(source='new', body=b'new')):
                with self.subTest(kind=kind, message=msg):
                    a, _, _ = fixture(kind)
                    result = a.workflow(msg, registration=True)
                    self.assertIn(result, (Decision.REUSED, 'UNMATCHED_NO_EXPORT'))
                    self.assertEqual(a.new_files, 0)
                    self.assertEqual(len(a.files), 1)

    def test_all_global_gates_precede_known_change_bypass(self):
        for gate in ('journal_ok', 'catalog_ok', 'mapped', 'stats_present'):
            with self.subTest(gate=gate):
                a, _, _ = fixture()
                self.assertEqual(a.workflow(Message(body=b'changed'), **{gate: False}), Decision.REVIEW)
                self.assertEqual(a.new_files, 0)

    def test_verified_change_may_bypass_unrelated_unknown_legacy_and_old_date(self):
        a, _, _ = fixture()
        self.assertEqual(a.workflow(Message(body=b'changed'), unknown_legacy=100, old_date=True), 'NEW')

    def test_subject_or_messageid_hints_cannot_authorize_unknown_old_content(self):
        a, _, _ = fixture()
        current = Message(source='different-source', body=b'different bounded payload')
        self.assertEqual(a.workflow(current), Decision.REVIEW)
        self.assertEqual(a.new_files, 0)

    def test_provider_revision_conflict_cannot_overwrite_old_proof(self):
        a, old, data = fixture('N')
        self.assertEqual(a.workflow(Message(revision='revision-1', body=b'changed')), Decision.REVIEW)
        self.assertEqual(a.records[old.key], old)
        self.assertEqual(a.files[old.path], data)
        self.assertEqual(a.new_files, 0)

    def test_last_accepted_journal_order_beats_provider_clock(self):
        a, old, _ = fixture()
        later, data = Proof.from_message(Message(), provider_time=-1000, path='later.msg')
        a.accept(later, data)
        self.assertLess(later.provider_time, old.provider_time)
        self.assertEqual(a.head[old.source], later.key)

    def test_attachment_rows_do_not_displace_source_head(self):
        a, old, _ = fixture()
        attachment = replace(old, key='A|attachment', path='attachment.bin')
        a.accept(attachment, b'attachment', kind='A')
        self.assertEqual(a.head[old.source], old.key)


class RevisionFailureInjection(unittest.TestCase):
    def test_missing_corrupt_empty_or_unreadable_prior_file_never_authorizes_export(self):
        for state in ('missing', b'corrupt', b'', None):
            for body in (b'old body', b'changed'):
                with self.subTest(state=state, body=body):
                    a, p, _ = fixture()
                    if state == 'missing':
                        del a.files[p.path]
                    else:
                        a.files[p.path] = state
                    self.assertEqual(a.workflow(Message(body=body)), Decision.REVIEW)
                    self.assertEqual(a.new_files, 0)
                    self.assertFalse(a.guarded)

    def test_unsupported_or_unknown_comparison_cannot_create_output(self):
        for result in (2, -1, 99):
            a, _, _ = fixture()
            self.assertEqual(a.workflow(Message(body=b'changed'), comparison=result), Decision.REVIEW)
            self.assertEqual(a.new_files, 0)

    def test_cancellation_cannot_create_alias_or_new_file(self):
        a, _, _ = fixture()
        self.assertEqual(a.workflow(Message(body=b'changed'), cancelled=True), Decision.CANCELLED)
        self.assertEqual(a.new_files + a.aliases, 0)
        self.assertFalse(a.guarded)

    def test_changed_live_source_cannot_authorize_old_or_new_output(self):
        for body in (b'old body', b'changed'):
            a, _, _ = fixture()
            self.assertEqual(a.workflow(Message(body=body), stable=False), Decision.REVIEW)
            self.assertEqual(a.new_files + a.aliases, 0)

    def test_mismatched_source_or_store_head_is_rejected(self):
        for change in ({'source': 'wrong'}, {'store': 'wrong'}):
            a, p, _ = fixture()
            a.records[p.key] = replace(p, **change)
            self.assertEqual(a.workflow(Message(body=b'changed')), Decision.REVIEW)
            self.assertEqual(a.new_files + a.aliases, 0)

    def test_missing_index_head_is_not_absent_history(self):
        a, p, _ = fixture()
        del a.records[p.key]
        self.assertEqual(a.workflow(Message(body=b'changed')), Decision.REVIEW)
        self.assertEqual(a.new_files, 0)

    def test_byref_result_cannot_replace_immutable_expected_proof(self):
        for index, value in enumerate(('other.msg', 'other-root', 'new-digest', 99999)):
            for comparison in (0, 1):
                with self.subTest(index=index, comparison=comparison):
                    a, p, _ = fixture()
                    outputs = [p.path, p.root, p.expected, p.size]
                    outputs[index] = value
                    self.assertEqual(a.workflow(Message(), comparison=comparison,
                                                returned=tuple(outputs)), Decision.REVIEW)
                    self.assertEqual(a.new_files + a.aliases, 0)
                    self.assertEqual(a.records[p.key], p)

    def test_equal_candidate_guard_stays_held_through_journal_commit(self):
        a, _, _ = fixture()
        self.assertEqual(a.known(Message(), attempt_write=True), Decision.REUSED)
        self.assertEqual(a.events, ['guard', 'verify_expected_bytes', 'compare_content',
                                   'source_stable', 'write_denied', 'commit_alias', 'release'])

    def test_writer_failure_releases_guard_and_preserves_old_proof(self):
        a, p, data = fixture()
        with self.assertRaises(OSError):
            a.known(Message(), writer_fails=True)
        self.assertFalse(a.guarded)
        self.assertEqual(a.records, {p.key: p})
        self.assertEqual(a.files[p.path], data)
        self.assertEqual(a.new_files + a.aliases, 0)

    def test_conflicting_current_target_proof_is_never_replaced(self):
        a, p, data = fixture()
        target = replace(p, key='M|revision-2', path='different.msg', expected='other')
        a.records[target.key] = target
        self.assertEqual(a.known(Message()), Decision.REVIEW)
        self.assertEqual(a.records[target.key], target)
        self.assertEqual(a.files[p.path], data)


class ActualSourceContracts(unittest.TestCase):
    def test_cache_is_one_direct_head_and_only_accepted_mail_kinds(self):
        body = procedure(source('mb_revision.basfrag'), 'MBV_CacheRecord')
        self.assertIn('gRevisionHeads(CStr(f(6))) = MBX_IndexKey(kind, CStr(f(2)))', body)
        self.assertNotIn('For Each', body)
        self.assertNotIn('New Collection', body)
        for kind in ('M', 'N', 'L', 'R'):
            self.assertIn('kind <> "' + kind + '"', body)
        layout = procedure(source('mb_layout_integration.basfrag'), 'MBX_CacheRecord')
        self.assertLess(layout.index('gRecordMeta(key) ='), layout.index('MBV_CacheRecord f'))

    def test_prior_hash_snapshot_predates_byref_and_survives_both_results(self):
        b = procedure(source('mb_revision.basfrag'), 'MBV_TryKnownRevision')
        order = ['expectedHash = CStr(fields(2))', 'guard = MBH_GuardFile(expectedPath)',
                 'result = MBI_CheckCandidate', 'digest <> expectedHash Or bytes <> expectedBytes',
                 'If Not MB_SourceStable(mail, q, revision)', 'If result = 0 Then']
        self.assertEqual([b.index(x) for x in order], sorted(b.index(x) for x in order))
        self.assertEqual(len(re.findall(r'^\s*expectedHash\s*=', b, re.M)), 1)

    def test_alias_guard_spans_commit_and_failure_cleanup(self):
        b = procedure(source('mb_revision.basfrag'), 'MBV_TryKnownRevision')
        begin = b.index('guard = MBH_GuardFile(expectedPath)')
        commit = b.index('MB_RecordCommit "R"')
        release = b.index('MBH_CloseGuard guard', commit)
        self.assertTrue(begin < commit < release)
        self.assertIn('If gFatal Then Err.Raise', b[b.index('Failed:'):])
        self.assertIn('MBH_CloseGuard guard', b[b.index('Failed:'):])
        self.assertNotIn('.SaveAs', b)
        self.assertNotIn('DeleteFile', b)

    def test_same_revision_and_conflict_cannot_create_self_reference(self):
        b = procedure(source('mb_revision.basfrag'), 'MBV_TryKnownRevision')
        changed = b[b.index('If result = 0 Then'):b.index('MBV_TryKnownRevision = MBV_CHANGED')]
        self.assertIn('If gIndex.Exists(targetKey) Then', changed)
        self.assertIn('GoTo NeedsReview', changed)
        commit = b.index('MB_RecordCommit "R"')
        equal = b[b.index('MBV_TryKnownRevision = MBV_CHANGED'):commit]
        self.assertIn('If gIndex.Exists(targetKey) Then', equal)
        self.assertIn('Else', equal)

    def test_known_change_bypass_preserves_global_integrity_gates(self):
        b = procedure(source('mb_reuse.basfrag'), 'MB_NewExportAllowed')
        bypass = b.index('If verifiedKnownChange Then')
        for condition in ('If gJournalIssues > 0', 'If Not gLegacyScanComplete',
                          'If Len(root) = 0', 'gLegacyAccountStats.Exists(root)',
                          'stats = gLegacyAccountStats(root)'):
            self.assertLess(b.index(condition), bypass)
        self.assertLess(bypass, b.index('If CLng(stats(1)) > 0'))
        self.assertLess(bypass, b.index('MB_EffectiveDate(mail) <= CDate(stats(2))'))

    def test_reset_clears_reconstructed_head(self):
        b = procedure(source('mb_identity.basfrag'), 'MBI_Reset')
        self.assertIn('MBV_Reset', b)
        reset = procedure(source('mb_revision.basfrag'), 'MBV_Reset')
        self.assertIn('Set gRevisionHeads = Nothing', reset)

    def test_fidelity_mismatch_and_unsupported_metadata_are_distinguished(self):
        b = procedure(source('mb_identity.basfrag'), 'MBI_CompareContents')
        classification = re.search(r'(?m)^\s*Case 158[^\r\n]*', b).group(0)
        self.assertIn('1120 To 1122', classification)
        self.assertNotIn('1123', classification)
        self.assertNotIn('1124', classification)
        self.assertIn('MBI_CompareContents = 2', b)

    def test_main_prioritizes_same_source_then_all_aliases(self):
        b = procedure(source('mb_main.basfrag'), 'MB_ProcessOne')
        order = ['MBV_TryKnownRevision', 'MBI_TryAlias', 'MB_ReuseLegacy',
                 'If gMode = "REGISTER_LEGACY" Then', 'MB_NewExportAllowed', 'mail.SaveAs']
        self.assertEqual([b.index(x) for x in order], sorted(b.index(x) for x in order))
        self.assertNotIn('If Not gLegacySources.Exists(CStr(q(7))) Then', b)
        self.assertIn('If Not verifiedKnownChange Then', b)
        self.assertIn('MB_NewExportAllowed(mail, legacyReason, verifiedKnownChange)', b)
        registration = b[b.index('If gMode = "REGISTER_LEGACY" Then'):b.index('MB_NewExportAllowed')]
        self.assertIn('Exit Function', registration)


if __name__ == '__main__':
    unittest.main(verbosity=2)
