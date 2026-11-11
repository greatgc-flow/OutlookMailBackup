"""Independent restoration-proof merge models plus actual VBA source contracts.

No Outlook, Win32, Office VBA compiler or real restore is executed. Models cover
historical evidence preservation and the old-snapshot/newer-current counterexamples.
"""
from dataclasses import dataclass, replace
from pathlib import Path
import ntpath
import re
import unittest

ROOT = Path(__file__).resolve().parent


class Conflict(Exception):
    pass


@dataclass(frozen=True)
class Move:
    op: str
    phase: str
    identity: str = 'immutable-path-size-hash'
    stamp: int = 0


def merge_moves(snapshot, current):
    a = [r for r in snapshot if r is not None]
    b = [r for r in current if r is not None]
    if any(x != y for x, y in zip(a, b)):
        raise Conflict('ordered move histories diverge')
    chosen = a if len(a) >= len(b) else b
    states, identities = {}, {}
    transitions = {
        None: {'PREPARE'}, 'PREPARE': {'COMMIT', 'ABORT'},
        'ABORT': {'PREPARE'}, 'COMMIT': {'UNDO_PREPARE'},
        'UNDO_PREPARE': {'UNDO_COMMIT', 'UNDO_ABORT'},
        'UNDO_ABORT': {'UNDO_PREPARE'}, 'UNDO_COMMIT': {'PREPARE'},
    }
    for row in chosen:
        if not isinstance(row, Move):
            raise Conflict('unchecked move row')
        if row.op in identities and identities[row.op] != row.identity:
            raise Conflict('operation expected bytes or paths changed')
        if row.phase not in transitions.get(states.get(row.op), set()):
            raise Conflict('invalid ordered phase')
        states[row.op] = row.phase
        identities[row.op] = row.identity
    return chosen


@dataclass(frozen=True)
class Stage:
    stage: str = r'C:\Account\~mb_owned.tmp'
    sequence: int = 1
    phase: str = 'RESERVE'
    final: str = r'C:\Account\mail.msg'
    digest: str = ''
    size: int = 0
    held: str = ''
    stamp: int = 1
    root: str = r'C:\Managed'

    @property
    def identity(self):
        return self.stage.lower(), self.sequence


@dataclass(frozen=True)
class Damaged:
    payload: object
    complete: bool = True


def valid_stage(row):
    if not isinstance(row, Stage):
        raise Conflict('no stage schema')
    if row.root != r'C:\Managed' or not 1 <= row.sequence <= 2_000_000_000:
        raise Conflict('stage owner/sequence')
    if type(row.sequence) is not int or row.stamp <= 0:
        raise Conflict('stage sequence/time')
    if ntpath.dirname(row.stage).lower() != ntpath.dirname(row.final).lower():
        raise Conflict('stage/final not siblings')
    if not ntpath.basename(row.stage).lower().startswith('~mb_') or row.stage.lower() == row.final.lower():
        raise Conflict('stage name')
    if row.phase in ('RESERVE', 'ABORTED'):
        if row.digest or row.size != 0 or row.held:
            raise Conflict('empty phase with byte proof')
    elif row.phase in ('PREPARED', 'PUBLISHED', 'RECOVERED', 'CANDIDATE', 'HOLD_PREPARE', 'HELD'):
        if not row.digest or row.size < 0:
            raise Conflict('byte proof absent')
        if row.phase in ('HOLD_PREPARE', 'HELD'):
            if row.held.lower() != (row.stage + '.held').lower():
                raise Conflict('held path')
        elif row.held:
            raise Conflict('unexpected held path')
    else:
        raise Conflict('unknown stage phase')
    return row.identity


def merge_stages(snapshot, current, mirror=()):
    proofs, origins, bad, destinations = {}, {}, [], {}
    for origin, rows in ((1, snapshot), (2, current), (4, mirror)):
        for row in rows:
            if isinstance(row, Damaged):
                if not row.complete or not isinstance(row.payload, Stage):
                    raise Conflict('no full framing-repair payload')
                bad.append((row.payload, origin))
                continue
            key = valid_stage(row)
            destination_key = row.stage.lower()
            if destination_key in destinations and destinations[destination_key] != row.final.lower():
                raise Conflict('stage later retargeted')
            destinations[destination_key] = row.final.lower()
            if key in proofs and proofs[key] != row:
                raise Conflict('same stage/sequence expected payload differs')
            proofs[key] = row
            origins[key] = origins.get(key, 0) | origin
    for payload, origin in bad:
        key = payload.identity
        if proofs.get(key) != payload or not origins.get(key, 0) & (7 ^ origin):
            raise Conflict('no exact independently checked counterpart')
    return list(proofs.values())


@dataclass(frozen=True)
class Baseline:
    path: str
    size: int
    digest: str
    context: tuple = ('managed', 'store', 'account')


def merge_baselines(a, b):
    result = {}
    context = None
    for row in a + b:
        if not isinstance(row, Baseline) or row.size <= 0 or not row.digest:
            raise Conflict('invalid baseline')
        if context is not None and row.context != context:
            raise Conflict('incoherent file context')
        context = row.context
        key = row.path.lower()
        if key in result and (result[key].size, result[key].digest) != (row.size, row.digest):
            raise Conflict('historical baseline byte conflict')
        result.setdefault(key, row)
    return list(result.values())


@dataclass(frozen=True)
class Audit:
    key: str
    status: str
    proof: tuple = ()
    origin: str = 'COMMITTED_INDEX'
    detail: str = 'reason'
    context: tuple = ('managed', 'store', 'account')


CURSOR = Audit('', 'CURSOR', detail='old-cursor')


def merge_audits(a, b):
    obligations = {}
    context = None
    for rows in (a, b):
        if rows is None:
            continue
        terminal = False
        for row in rows:
            if terminal or not isinstance(row, Audit):
                raise Conflict('bad or post-terminal audit data')
            if context is not None and row.context != context:
                raise Conflict('audit bindings changed')
            context = row.context
            if row.status == 'CURSOR':
                terminal = True
            elif row.status == 'REVIEW':
                if row.key in obligations:
                    old = obligations[row.key]
                    if old.proof != row.proof:
                        raise Conflict('proof identity mismatch')
                    if old.origin == 'COMMITTED_INDEX':
                        continue
                obligations[row.key] = row
            elif row.status not in ('OK', 'ABSENT'):
                raise Conflict('unknown audit status')
            # OK and ABSENT deliberately do not clear recovered obligations.
        if not terminal:
            raise Conflict('incomplete compact snapshot')
    if context is None:
        raise Conflict('audit has no binding context')
    return list(obligations.values()) + [replace(CURSOR, detail='', context=context)]


def merge_bindings(snapshot, current, *, selected=('store-A', r'C:\Account-A'), markers=()):
    roots = dict(snapshot)
    roots.update(current)
    for store, path in roots.items():
        if store == selected[0]:
            if path.lower() != selected[1].lower():
                raise Conflict('current owner disagrees with selected account')
        elif (store, path) not in markers:
            raise Conflict('other owner lacks independent marker')
        for other, other_path in roots.items():
            if other != store:
                a, b = path.lower().rstrip('\\'), other_path.lower().rstrip('\\')
                if a == b or a.startswith(b + '\\') or b.startswith(a + '\\'):
                    raise Conflict('account roots overlap')
    return roots


def source(name):
    return (ROOT / name).read_text(encoding='ascii')


def proc(text, name):
    match = re.search(r'(?m)^Private (?:Sub|Function) ' + name + r'\b', text)
    if match is None:
        raise AssertionError('missing ' + name)
    end = re.search(r'(?m)^End (?:Sub|Function)', text[match.end():])
    return text[match.start():match.end() + end.end()]


class MoveProofCases(unittest.TestCase):
    def test_longer_current_tail_preserves_later_commit_and_undo(self):
        p, c, u, done = [Move('x', s, stamp=i) for i, s in enumerate(('PREPARE', 'COMMIT', 'UNDO_PREPARE', 'UNDO_COMMIT'))]
        current = [p, c, u, done]
        self.assertEqual(merge_moves([p, c], current), current)
        self.assertEqual(merge_moves(current, [p]), current)

    def test_blank_lines_do_not_fake_divergence(self):
        p, c = Move('x', 'PREPARE'), Move('x', 'COMMIT')
        self.assertEqual(merge_moves([None, p, None], [p, None, c]), [p, c])

    def test_divergent_valid_branches_do_not_become_unordered_union(self):
        p = Move('x', 'PREPARE')
        with self.assertRaises(Conflict):
            merge_moves([p, Move('x', 'COMMIT')], [p, Move('x', 'ABORT')])

    def test_raw_prefix_with_semantically_invalid_extra_phase_stops(self):
        p = Move('x', 'PREPARE')
        for tail in ([p], [Move('x', 'UNDO_COMMIT')], ['unchecked']):
            with self.subTest(tail=tail), self.assertRaises(Conflict):
                merge_moves([p], [p] + tail)

    def test_different_timestamp_on_same_position_is_not_same_history(self):
        with self.assertRaises(Conflict):
            merge_moves([Move('x', 'PREPARE', stamp=1)], [Move('x', 'PREPARE', stamp=2)])

    def test_different_expected_bytes_cannot_change_existing_operation(self):
        with self.assertRaises(Conflict):
            merge_moves([], [Move('x', 'PREPARE'), Move('x', 'COMMIT', identity='different-hash')])

    def test_interleaved_independent_operations_remain_in_order(self):
        rows = [Move('a', 'PREPARE'), Move('b', 'PREPARE'), Move('a', 'COMMIT'), Move('b', 'ABORT')]
        self.assertEqual(merge_moves(rows[:2], rows), rows)


class StageProofCases(unittest.TestCase):
    def setUp(self):
        self.reserve = Stage()
        self.prepared = replace(self.reserve, sequence=2, phase='PREPARED', digest='hash-A', size=100)
        self.published = replace(self.prepared, sequence=3, phase='PUBLISHED')

    def test_later_publication_survives_older_snapshot(self):
        merged = merge_stages([self.reserve], [self.reserve, self.prepared, self.published])
        self.assertEqual(merged, [self.reserve, self.prepared, self.published])

    def test_later_aborted_row_does_not_erase_published_expected_proof(self):
        aborted = replace(self.reserve, sequence=4, phase='ABORTED')
        merged = merge_stages([self.published], [aborted])
        self.assertIn(self.published, merged)
        self.assertIn(aborted, merged)

    def test_same_sequence_conflicting_payload_stops(self):
        for altered in (replace(self.published, digest='hash-B'), replace(self.published, stamp=99)):
            with self.subTest(altered=altered), self.assertRaises(Conflict):
                merge_stages([self.published], [altered])

    def test_current_framing_damage_needs_exact_other_copy(self):
        self.assertEqual(merge_stages([self.published], [Damaged(self.published)]), [self.published])
        with self.assertRaises(Conflict):
            merge_stages([self.prepared], [Damaged(self.published)])

    def test_same_file_duplicate_cannot_repair_its_own_damaged_proof(self):
        with self.assertRaises(Conflict):
            merge_stages([], [self.published, Damaged(self.published)])

    def test_immutable_independent_mirror_repairs_primary_with_no_snapshot_proof(self):
        self.assertEqual(merge_stages([], [Damaged(self.published)], [self.published]), [self.published])

    def test_primary_repairs_exact_mirror_framing_damage(self):
        self.assertEqual(merge_stages([], [self.published], [Damaged(self.published)]), [self.published])

    def test_conflicting_independent_mirror_stops_even_if_snapshot_is_valid(self):
        with self.assertRaises(Conflict):
            merge_stages([self.published], [self.published], [replace(self.published, digest='other')])

    def test_only_one_surviving_mirror_is_one_origin_not_two(self):
        # Root passes the sole mirror as priorRaw and leaves optional mirror
        # empty, so a damaged duplicate from that same file cannot self-prove.
        with self.assertRaises(Conflict):
            merge_stages([], [self.published, Damaged(self.published)], [])

    def test_partial_unknown_or_conflicting_damage_never_creates_authority(self):
        for damaged in (Damaged(self.published, False), Damaged('partial'), Damaged(replace(self.published, digest='wrong'))):
            with self.subTest(damaged=damaged), self.assertRaises(Conflict):
                merge_stages([self.published], [damaged])

    def test_unsafe_phase_path_binding_hash_and_sequence_are_rejected(self):
        mutations = [dict(root='wrong'), dict(sequence=0), dict(sequence=1.5), dict(phase='UNKNOWN'),
                     dict(stage=r'C:\Elsewhere\~mb_x.tmp'), dict(stage=r'C:\Account\ordinary.tmp'),
                     dict(digest=''), dict(held=r'C:\Account\wrong'), dict(stamp=0)]
        for delta in mutations:
            with self.subTest(delta=delta), self.assertRaises(Conflict):
                merge_stages([], [replace(self.published, **delta)])

    def test_same_stage_cannot_later_target_other_final(self):
        with self.assertRaises(Conflict):
            merge_stages([self.reserve], [replace(self.prepared, final=r'C:\Account\other.msg')])


class BaselineAndAuditCases(unittest.TestCase):
    def test_newer_msg_only_baseline_is_union_not_rollback(self):
        a, b = Baseline('a.msg', 10, 'A'), Baseline('b.msg', 20, 'B')
        self.assertEqual(merge_baselines([a], [a, b]), [a, b])

    def test_same_baseline_path_cannot_accept_newly_observed_hash(self):
        a = Baseline('a.msg', 10, 'A')
        for delta in (dict(digest='B'), dict(size=11)):
            with self.subTest(delta=delta), self.assertRaises(Conflict):
                merge_baselines([a], [replace(a, **delta)])

    def test_baseline_other_account_is_preserved_in_its_own_context(self):
        row = Baseline('x.msg', 10, 'A', ('managed', 'other-store', 'other-account'))
        self.assertEqual(merge_baselines([], [row]), [row])
        with self.assertRaises(Conflict):
            merge_baselines([row], [replace(row, context=('managed', 'wrong-store', 'other-account'))])

    def test_newer_review_and_old_unresolved_review_both_survive(self):
        a, b = Audit('A', 'REVIEW', ('a.msg', 10, 'hashA')), Audit('B', 'REVIEW', ('b.msg', 20, 'hashB'))
        result = merge_audits([a, CURSOR], [b, CURSOR])
        self.assertEqual({r.key for r in result[:-1]}, {'A', 'B'})
        self.assertEqual(result[-1].status, 'CURSOR')
        self.assertEqual(result[-1].detail, '')
        self.assertEqual(sum(r.status == 'CURSOR' for r in result), 1)

    def test_current_ok_or_absent_cannot_erase_old_review_on_recovery(self):
        review = Audit('A', 'REVIEW', ('a.msg', 10, 'hashA'))
        for status in ('OK', 'ABSENT'):
            result = merge_audits([review, CURSOR], [replace(review, status=status), CURSOR])
            self.assertEqual(result[0], review)

    def test_expected_proof_tombstone_survives_missing_index_and_baseline(self):
        tombstone = Audit('missing-index-proof', 'REVIEW', ('missing.msg', 500, 'expected'))
        index, baseline = {}, []
        self.assertFalse(index or baseline)
        result = merge_audits([CURSOR], [tombstone, CURSOR])
        self.assertEqual(result[0].proof, tombstone.proof)

    def test_committed_proof_is_not_downgraded_to_fresh_msg_only_baseline(self):
        strong = Audit('A', 'REVIEW', ('a', 1, 'hash'))
        weak = replace(strong, origin='MSG_ONLY_REBASELINED')
        for a, b in ((strong, weak), (weak, strong)):
            self.assertEqual(merge_audits([a, CURSOR], [b, CURSOR])[0].origin, 'COMMITTED_INDEX')

    def test_incomplete_double_terminal_or_postcursor_audit_stops(self):
        review = Audit('A', 'REVIEW')
        for malformed in ([], [review], [CURSOR, review], [CURSOR, CURSOR], ['broken', CURSOR]):
            with self.subTest(rows=malformed), self.assertRaises(Conflict):
                merge_audits([CURSOR], malformed)

    def test_missing_snapshot_audit_recovers_complete_current(self):
        row = Audit('A', 'REVIEW')
        self.assertEqual(merge_audits(None, [row, CURSOR]), [row, replace(CURSOR, detail='')])

    def test_restore_merge_is_idempotent_without_removing_obligations(self):
        a, b = Audit('A', 'REVIEW'), Audit('B', 'REVIEW')
        once = merge_audits([a, CURSOR], [b, CURSOR])
        twice = merge_audits(once, [b, CURSOR])
        self.assertEqual(once, twice)


class ActualSourceContracts(unittest.TestCase):
    def test_checked_stage_publication_is_after_complete_merge(self):
        b = proc(source('mb_recovery_proofs.basfrag'), 'MBR_MergeProofFile')
        order = ['MBH_NewStagePath(target)', 'MBR_CollectProof priorRaw', 'For Each key In bad.Keys',
                 'writer.Close', 'MBM_FlushPath stage', 'MBR_PublishOverlayStage stage, target']
        self.assertEqual([b.index(x) for x in order], sorted(b.index(x) for x in order))
        self.assertIn('MBH_DeleteStage stage', b[b.index('Failed:'):])
        self.assertNotIn('DeleteFile', b)

    def test_moves_use_one_real_semantic_replay_and_restore_isolated_caches(self):
        b = proc(source('mb_recovery_proofs.basfrag'), 'MBR_MergeMoveProof')
        self.assertIn('If a <> b Then Err.Raise', b)
        self.assertIn('MBM_ValidateMoveRow chosen', b)
        self.assertIn('MBM_CacheMove chosen', b)
        self.assertIn('Set oldOps = mbmOps: Set oldRedirects = mbmRedirects', b)
        self.assertIn('Set mbmOps = oldOps: Set mbmRedirects = oldRedirects', b)
        code = '\n'.join(line for line in b.splitlines() if not line.strip().startswith("'"))
        self.assertNotIn('MBM_Recover ', code)
        self.assertNotIn('MBM_LoadMoves', code)

    def test_stage_rows_not_collapsed_to_latest_sequence(self):
        b = proc(source('mb_recovery_proofs.basfrag'), 'MBR_CollectProof')
        self.assertIn('identity = MBR_StageProofIdentity(f)', b)
        self.assertIn('rows(identity) = payload', b)
        self.assertIn('origins(identity) = CLng(origins(identity)) Or origin', b)
        self.assertNotIn('sequence >', b)
        stage = proc(source('mb_recovery_proofs.basfrag'), 'MBR_StageProofIdentity')
        for phase in ('PREPARED', 'PUBLISHED', 'RECOVERED', 'CANDIDATE', 'HOLD_PREPARE', 'HELD'):
            self.assertIn('"' + phase + '"', stage)
        self.assertIn('MBSR_Identity(f)', stage)
        main = proc(source('mb_recovery_proofs.basfrag'), 'MBR_MergeProofFile')
        self.assertIn('Optional ByVal stageMirrorRaw As String = ""', main)
        self.assertIn('context, 4, accountRoot, storeId', main)
        self.assertIn('(7 Xor CLng(f(2)))', main)
        self.assertIn('StrComp(stageMirrorRaw, priorRaw, vbTextCompare) = 0', main)

    def test_recovered_ok_does_not_remove_review_and_one_cursor_is_written(self):
        b = proc(source('mb_recovery_proofs.basfrag'), 'MBR_CollectProof')
        self.assertIn('If terminal Then Err.Raise', b)
        self.assertIn('If Not terminal Then Err.Raise', b)
        self.assertNotIn('rows.Remove', b)
        main = proc(source('mb_recovery_proofs.basfrag'), 'MBR_MergeProofFile')
        self.assertEqual(main.count('Array("MBSOURCEAUDIT1"'), 1)
        self.assertEqual(main.count('Array("MBARCHIVEAUDIT1"'), 1)
        self.assertLess(main.index('Next key'), main.index('If kind = "SOURCE" Then'))

    def test_baseline_checks_old_digest_and_size_without_current_msg_read(self):
        b = proc(source('mb_recovery_proofs.basfrag'), 'MBR_CollectProof')
        self.assertIn('CStr(prior(6)) <> CStr(f(6))', b)
        self.assertIn('MB_ByteCount(CStr(prior(5))) <> MB_ByteCount(CStr(f(5)))', b)
        self.assertNotIn('MBH_HashFile', b)
        self.assertNotIn('OpenSharedItem', b)
        self.assertNotIn('.SaveAs', b)

    def test_binding_uses_leaf_owner_context_without_uninitialized_audit_globals(self):
        text = source('mb_recovery_proofs.basfrag')
        b = proc(text, 'MBR_ProofContext')
        self.assertIn('StrComp(leaf, expectedLeaf, vbTextCompare)', b)
        for global_name in ('gWindowStore', 'gArchiveAuditStore', 'gArchiveAuditAccount'):
            self.assertNotIn(global_name, text)
        self.assertIn('MBY_CheckMarker', text)
        self.assertIn('Recovered mappings overlap different stores.', text)

    def test_no_public_macro_and_uses_ascii_fragment(self):
        text = source('mb_recovery_proofs.basfrag')
        self.assertNotRegex(text, r'(?m)^Public ')
        self.assertLess(max(map(len, text.splitlines())), 1024)


class AccountBindingCases(unittest.TestCase):
    def test_safe_current_selected_binding_supersedes_older_root(self):
        selected = ('store-A', r'C:\Account-A')
        merged = merge_bindings({'store-A': r'C:\Old-Account'}, dict([selected]))
        self.assertEqual(merged, dict([selected]))

    def test_other_independently_marked_owner_is_retained(self):
        other = ('store-B', r'C:\Account-B')
        merged = merge_bindings({'store-A': r'C:\Account-A'}, dict([other]), markers=[other])
        self.assertEqual(len(merged), 2)

    def test_unproved_other_owner_binding_stops(self):
        with self.assertRaises(Conflict):
            merge_bindings({}, {'store-B': r'C:\Account-B'})

    def test_different_store_overlapping_paths_stop(self):
        for path in (r'C:\Account-A', r'C:\Account-A\nested'):
            other = ('store-B', path)
            with self.subTest(path=path), self.assertRaises(Conflict):
                merge_bindings({'store-A': r'C:\Account-A'}, dict([other]), markers=[other])

    def test_stale_current_store_root_cannot_override_independent_owner(self):
        with self.assertRaises(Conflict):
            merge_bindings({'store-A': r'C:\Old-Account'}, {})


if __name__ == '__main__':
    unittest.main(verbosity=2)
