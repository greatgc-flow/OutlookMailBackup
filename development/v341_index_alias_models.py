"""Immutable journal and indexed alias policy regression models.

The slow decision oracle and independent in-memory lookup cover selection,
failure/cancellation boundaries and operation counts. Source contracts bind the
invariants to the VBA entry points. This does not execute Outlook or compile VBA.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, replace
from itertools import product
from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parent
LIMIT = 128


@dataclass(frozen=True)
class Proof:
    key: str = "M|revision"
    digest: str = "original-sha"
    size: int = 10
    source: str = "source-A"
    store: str = "store-A"
    path: str = "old.msg"
    policy: str = "V2"


class ImmutableJournal:
    def __init__(self):
        self.cache = {}
        self.journal = []

    def accept(self, proof):
        key = proof.key
        if key.startswith(("N|", "R|")):
            key = "M|" + key[2:]
        prior = self.cache.get(key)
        if prior is not None:
            if (prior.digest, prior.size, prior.source, prior.store) != (
                    proof.digest, proof.size, proof.source, proof.store):
                raise ValueError("conflicting immutable expected proof")
        self.journal.append(proof)
        self.cache[key] = proof


@dataclass(frozen=True)
class Message:
    subject: str = "Daily notification"
    store: str = "store-A"
    has_id: bool = True
    message_id: str = "id-A"
    sent: str = "sent-A"
    received: str = "received-A"
    stamp: str = "received-A"
    body: str = "content-A"

    @property
    def bucket(self):
        return self.store, self.subject


@dataclass(frozen=True)
class Record:
    key: str
    mail: Message = Message()
    path: str = "a.msg"
    digest: str = "sha-A"
    header_ok: bool = True
    file_ok: bool = True
    supported: bool = True

    @property
    def physical(self):
        return self.path.casefold(), self.digest


def plausible(source, candidate):
    if source.has_id and candidate.has_id and source.message_id and candidate.message_id:
        return source.message_id == candidate.message_id
    if not source.has_id and not candidate.has_id:
        return source.sent == candidate.sent
    return True


def compare(source, record):
    target = record.mail
    if not record.file_ok:
        return "REVIEW"
    if source.subject != target.subject:
        return "DIFFERENT"
    if source.sent != target.sent or source.received != target.received:
        if source.has_id and target.has_id and source.message_id and target.message_id:
            return "REVIEW" if source.message_id == target.message_id else "DIFFERENT"
        if not source.has_id and not target.has_id and source.sent != target.sent:
            return "DIFFERENT"
        return "REVIEW"
    if source.has_id != target.has_id or source.message_id != target.message_id:
        return "DIFFERENT"
    if source.body != target.body:
        return "DIFFERENT"
    return "EQUAL" if record.supported else "REVIEW"


def evaluate_pass(source, records, seen, checked, uncertain):
    matches = []
    visits = 0
    for record in records:
        visits += 1
        if record.physical in seen:
            continue
        seen.add(record.physical)
        if not record.header_ok:
            uncertain = True
            continue
        if not plausible(source, record.mail):
            continue
        if checked >= LIMIT:
            return "REVIEW", "", checked, uncertain, visits
        checked += 1
        verdict = compare(source, record)
        if verdict == "REVIEW":
            uncertain = True
        elif verdict == "EQUAL":
            matches.append(record)
    if len({row.digest for row in matches}) > 1:
        return "REVIEW", "", checked, uncertain, visits
    if matches:
        return "REUSE", min(row.path for row in matches), checked, uncertain, visits
    return "NONE", "", checked, uncertain, visits


def slow_oracle(source, records):
    seen, checked, uncertain = set(), 0, False
    for exact in (True, False):
        candidates = [r for r in records if r.mail.bucket == source.bucket
                      and ((r.mail.stamp == source.stamp) == exact)]
        verdict, path, checked, uncertain, _ = evaluate_pass(
            source, candidates, seen, checked, uncertain)
        if verdict != "NONE":
            return verdict, path
    return ("REVIEW" if uncertain else "NONE"), ""


class IndexedLookup:
    def __init__(self):
        self.records = {}
        self.buckets = defaultdict(list)
        self.dates = defaultdict(list)
        self.seen_keys = set()
        self.date_keys = set()
        self.entries = defaultdict(list)
        self.counts = defaultdict(int)
        self.proofs = set()
        self.unknown = set()
        self.headers = set()
        self.build_visits = 0
        self.query_visits = 0
        self.header_reads = 0

    def add(self, record):
        self.records[record.key] = record
        identity = record.mail.bucket, record.key
        if identity not in self.seen_keys:
            self.seen_keys.add(identity)
            self.buckets[record.mail.bucket].append(record.key)
        date_identity = record.mail.bucket, record.mail.stamp, record.key
        if date_identity not in self.date_keys:
            self.date_keys.add(date_identity)
            self.dates[(record.mail.bucket, record.mail.stamp)].append(record.key)

    @staticmethod
    def groups(source):
        if source.has_id and source.message_id:
            return ("ID", source.message_id), ("ABSENT",), ("EMPTY",)
        if not source.has_id:
            return ("SENT", source.sent), ("PRESENT",), ("EMPTY",)
        return ("PRESENT",), ("ABSENT",), ("EMPTY",)

    def ensure(self, scope, keys, cancel_at=None):
        position = self.counts[scope]
        while position < len(keys):
            if cancel_at is not None and position >= cancel_at:
                return False
            self.build_visits += 1
            row = self.records[keys[position]]
            proof = scope, row.physical
            if proof not in self.proofs:
                if row.physical not in self.headers:
                    self.header_reads += 1
                if not row.header_ok:
                    self.unknown.add(scope)
                else:
                    self.headers.add(row.physical)
                    if row.mail.has_id and row.mail.message_id:
                        selectors = (("ID", row.mail.message_id), ("PRESENT",))
                    elif not row.mail.has_id:
                        selectors = (("SENT", row.mail.sent), ("ABSENT",))
                    else:
                        selectors = (("EMPTY",),)
                    for selector in selectors:
                        self.entries[(scope, selector)].append(row.key)
                self.proofs.add(proof)
            position += 1
            self.counts[scope] = position
        return True

    def query(self, source, *, cancel_at=None, stable=True, final_ok=True):
        seen, checked, uncertain = set(), 0, False
        for exact in (True, False):
            scope = source.bucket, source.stamp if exact else None
            keys = self.dates.get((source.bucket, source.stamp), ()) if exact else self.buckets[source.bucket]
            if not self.ensure(scope, keys, cancel_at):
                return "CANCELLED", ""
            uncertain |= scope in self.unknown
            candidates = (self.records[key]
                          for selector in self.groups(source)
                          for key in self.entries.get((scope, selector), ())
                          if ((self.records[key].mail.stamp == source.stamp) == exact))
            verdict, path, checked, uncertain, visits = evaluate_pass(
                source, candidates, seen, checked, uncertain)
            self.query_visits += visits
            if verdict == "REUSE" and (not stable or not final_ok):
                return "REVIEW", ""
            if verdict != "NONE":
                return verdict, path
        return ("REVIEW" if uncertain else "NONE"), ""


class ImmutableProofCases(unittest.TestCase):
    def setUp(self):
        self.archive = ImmutableJournal()
        self.first = Proof()
        self.archive.accept(self.first)

    def reject(self, **change):
        before = list(self.archive.journal)
        with self.assertRaises(ValueError):
            self.archive.accept(replace(self.first, **change))
        self.assertEqual(self.archive.journal, before)
        self.assertEqual(self.archive.cache[self.first.key], self.first)

    def test_changed_hash_rejected(self):
        self.reject(digest="reobserved-sha")

    def test_changed_size_rejected(self):
        self.reject(size=11)

    def test_changed_source_rejected(self):
        self.reject(source="source-B")

    def test_changed_store_rejected(self):
        self.reject(store="store-B")

    def test_normalized_n_and_r_keys_cannot_replace_m_proof(self):
        for kind in ("N", "R"):
            with self.subTest(kind=kind):
                self.reject(key=kind + "|revision", digest="new-sha")

    def test_validation_upgrade_keeps_expected_bytes(self):
        changed = replace(self.first, policy="V3")
        self.archive.accept(changed)
        self.assertEqual(self.archive.cache[self.first.key], changed)

    def test_verified_path_rebinding_keeps_expected_bytes(self):
        self.archive.accept(replace(self.first, path="2026/11/old.msg"))
        self.assertEqual(self.archive.cache[self.first.key].digest, self.first.digest)

    def test_idempotent_replay_keeps_proof(self):
        self.archive.accept(self.first)
        self.assertEqual(self.archive.cache[self.first.key], self.first)


class AliasPolicyCases(unittest.TestCase):
    def check(self, source, records):
        index = IndexedLookup()
        for row in records:
            index.add(row)
        self.assertEqual(index.query(source), slow_oracle(source, records))
        return index.query(source)

    def test_exact_id_presence_sent_and_date_grid_matches_slow_oracle(self):
        variants = list(product((False, True), ("", "id-A", "id-B"), ("sent-A", "sent-B"),
                                ("received-A", "other-date")))
        for sh, sid, sent, stamp in variants:
            source = Message(has_id=sh, message_id=sid if sh else "", sent=sent, stamp=stamp)
            for ch, cid, csent, cstamp in variants:
                candidate = Message(has_id=ch, message_id=cid if ch else "", sent=csent, stamp=cstamp)
                with self.subTest(source=source, candidate=candidate):
                    self.check(source, [Record("k", candidate)])

    def test_case_sensitive_id_does_not_alias(self):
        self.assertEqual(self.check(Message(message_id="ID-A"), [Record("k")]), ("NONE", ""))

    def test_missing_id_same_sent_is_compared(self):
        source = Message(has_id=False, message_id="")
        self.assertEqual(self.check(source, [Record("k", source)]), ("REUSE", "a.msg"))

    def test_missing_id_different_sent_is_excluded(self):
        source = Message(has_id=False, message_id="")
        self.assertEqual(self.check(source, [Record("k", replace(source, sent="other"))]), ("NONE", ""))

    def test_cross_presence_different_times_remains_review(self):
        source = Message()
        target = replace(source, has_id=False, message_id="", sent="other")
        self.assertEqual(self.check(source, [Record("k", target)]), ("REVIEW", ""))

    def test_present_empty_id_remains_plausible(self):
        source = Message(message_id="", sent="different")
        self.assertEqual(self.check(source, [Record("k")]), ("REVIEW", ""))

    def test_equal_contents_with_different_bytes_require_review(self):
        rows = [Record("a"), Record("b", path="b.msg", digest="different-sha")]
        self.assertEqual(self.check(Message(), rows), ("REVIEW", ""))

    def test_equal_bytes_choose_stable_path(self):
        rows = [Record("b", path="b.msg"), Record("a")]
        self.assertEqual(self.check(Message(), rows), ("REUSE", "a.msg"))

    def test_unknown_header_never_becomes_absent_id(self):
        self.assertEqual(self.check(Message(), [Record("bad", header_ok=False)]), ("REVIEW", ""))

    def test_exact_date_match_does_not_open_other_date_bucket(self):
        index = IndexedLookup()
        index.add(Record("a"))
        index.add(Record("other", replace(Message(), stamp="old-zone"), path="bad.msg", header_ok=False))
        self.assertEqual(index.query(Message()), ("REUSE", "a.msg"))
        self.assertEqual(index.header_reads, 1)

    def test_other_date_fallback_preserves_timezone_compatibility(self):
        self.assertEqual(self.check(Message(), [Record("old", replace(Message(), stamp="old-zone"))]),
                         ("REUSE", "a.msg"))

    def test_same_physical_alias_keeps_each_metadata_date_priority(self):
        rows = [Record("old", replace(Message(), stamp="old-zone")), Record("alias"),
                Record("conflict", path="b.msg", digest="other-sha")]
        self.assertEqual(self.check(Message(), rows), ("REVIEW", ""))

    def test_exact_date_success_preserves_prior_other_date_conflict_policy(self):
        rows = [Record("a"), Record("b", replace(Message(), stamp="old-zone"), path="b.msg", digest="other")]
        self.assertEqual(self.check(Message(), rows), ("REUSE", "a.msg"))

    def test_unsupported_content_is_not_certified(self):
        self.assertEqual(self.check(Message(), [Record("a", supported=False)]), ("REVIEW", ""))

    def test_same_subject_different_store_never_aliases(self):
        self.assertEqual(self.check(Message(store="other"), [Record("a")]), ("NONE", ""))

    def test_plausible_limit_still_blocks_export(self):
        rows = [Record(str(i), path=f"{i}.msg") for i in range(LIMIT + 1)]
        self.assertEqual(self.check(Message(), rows), ("REVIEW", ""))


class LookupLifecycleCases(unittest.TestCase):
    def test_incremental_suffix_classified_once(self):
        index = IndexedLookup()
        for i in range(20):
            source = Message(message_id=str(i), stamp=str(i))
            self.assertEqual(index.query(source), ("NONE", ""))
            index.add(Record(str(i), source, path=f"{i}.msg"))
        self.assertEqual(index.build_visits, 19)

    def test_cancelled_partial_build_resumes_without_restart(self):
        index = IndexedLookup()
        for i in range(10):
            index.add(Record(str(i), Message(message_id=str(i)), path=f"{i}.msg"))
        source = Message(message_id="none", stamp="absent")
        self.assertEqual(index.query(source, cancel_at=4), ("CANCELLED", ""))
        self.assertEqual(index.build_visits, 4)
        self.assertEqual(index.query(source), ("NONE", ""))
        self.assertEqual(index.build_visits, 10)

    def test_source_changed_after_lookup_blocks_alias(self):
        index = IndexedLookup(); index.add(Record("a"))
        self.assertEqual(index.query(Message(), stable=False), ("REVIEW", ""))

    def test_final_byte_revalidation_failure_blocks_alias(self):
        index = IndexedLookup(); index.add(Record("a"))
        self.assertEqual(index.query(Message(), final_ok=False), ("REVIEW", ""))

    def test_failed_header_retried_next_run(self):
        index = IndexedLookup(); index.add(Record("a", header_ok=False))
        self.assertEqual(index.query(Message()), ("REVIEW", ""))
        index.records["a"] = Record("a")
        self.assertEqual(index.query(Message()), ("REVIEW", ""))
        next_run = IndexedLookup(); next_run.add(Record("a"))
        self.assertEqual(next_run.query(Message()), ("REUSE", "a.msg"))

    def test_metadata_date_refresh_is_visible_without_all_bucket_reset(self):
        index = IndexedLookup(); index.add(Record("a", replace(Message(), stamp="old-zone")))
        self.assertEqual(index.query(Message(message_id="other")), ("NONE", ""))
        prior = index.build_visits
        index.add(Record("a"))
        self.assertEqual(index.query(Message()), ("REUSE", "a.msg"))
        self.assertEqual(index.build_visits, prior + 1)

    def test_100000_same_subject_distinct_ids_are_linear(self):
        index = IndexedLookup()
        for i in range(100000):
            source = Message(message_id=str(i), stamp=str(i))
            if index.query(source) != ("NONE", ""):
                self.fail(f"Unexpected candidate for {i}")
            index.add(Record(str(i), source, path=f"{i}.msg"))
        self.assertEqual(index.build_visits, 99999)
        self.assertEqual(index.header_reads, 99999)
        self.assertEqual(index.query_visits, 0)

    def test_100000_same_subject_absent_ids_use_sent_index(self):
        index = IndexedLookup()
        for i in range(100000):
            source = Message(has_id=False, message_id="", sent=str(i), stamp=str(i))
            if index.query(source) != ("NONE", ""):
                self.fail(f"Unexpected candidate for {i}")
            index.add(Record(str(i), source, path=f"{i}.msg"))
        self.assertEqual(index.build_visits, 99999)
        self.assertEqual(index.query_visits, 0)

    def test_100000_reimport_aliases_share_one_physical_candidate(self):
        index = IndexedLookup(); index.add(Record("original"))
        for i in range(100000):
            if index.query(Message()) != ("REUSE", "a.msg"):
                self.fail(f"Unexpected alias decision for {i}")
            index.add(Record(str(i)))
        self.assertEqual(index.build_visits, 100000)
        self.assertEqual(index.header_reads, 1)
        self.assertEqual(index.query_visits, 100000)


def procedure(text, name):
    found = re.search(r"(?ms)^Private (?:Sub|Function) " + re.escape(name) +
                      r"\b.*?^End (?:Sub|Function)\s*$", text)
    if not found:
        raise AssertionError(f"Missing procedure {name}")
    return found.group(0)


class ProductionContracts(unittest.TestCase):
    def source(self, filename):
        return (ROOT / filename).read_text(encoding="ascii")

    def test_every_index_cache_path_rejects_conflict_before_mutation(self):
        body = procedure(self.source("mb_layout_integration.basfrag"), "MBX_CacheRecord")
        check = body.index("MBR_RejectIndexConflict f")
        self.assertLess(check, body.index("Select Case kind"))

    def test_journal_append_is_after_immutable_validation(self):
        body = procedure(self.source("mb_main.basfrag"), "MB_RecordCommit")
        self.assertLess(body.index("MBR_RejectIndexConflict fields"), body.index("MB_OpenJournal"))
        self.assertLess(body.index("MBR_RejectIndexConflict fields"), body.index("MB_WriteCheckedLine gJournal"))

    def test_conflict_guard_preserves_hash_size_source_store_contract(self):
        body = procedure(self.source("mb_recovery_merge.basfrag"), "MBR_RejectIndexConflict")
        for fragment in ("prior(2)", "prior(1)", "meta(2)", "meta(0)", "MB_ERR + 1150"):
            self.assertIn(fragment, body)
        self.assertNotIn("gIndex(key) =", body)

    def test_lookup_context_reset_and_recovery_isolation(self):
        reset = procedure(self.source("mb_identity.basfrag"), "MBI_Reset")
        self.assertIn("Set gIdentityLookupState = Nothing", reset)
        for filename, name in (("mb_recovery.basfrag", "MBR_SalvageJournal"),
                               ("mb_recovery_merge.basfrag", "MBR_MergeJournal")):
            body = procedure(self.source(filename), name)
            self.assertIn("Set oldLookup = gIdentityLookupState", body)
            self.assertIn("Set gIdentityLookupState = oldLookup", body)

    def test_incremental_cursor_commits_after_complete_classification(self):
        body = procedure(self.source("mb_identity.basfrag"), "MBI_EnsureLookup")
        self.assertIn("position = CLng(counts(scope))", body)
        self.assertIn("Do While position < candidates.Count", body)
        self.assertLess(body.index("If result = 3 Then Exit Function"), body.index("counts(scope) = position"))
        self.assertLess(body.index("proofs.Add proof, True"), body.index("counts(scope) = position"))

    def test_physical_dedup_precedes_header_io(self):
        body = procedure(self.source("mb_identity.basfrag"), "MBI_EnsureLookup")
        self.assertIn('proof = scope & "|" & MB_LegacyOwnerKey', body)
        self.assertLess(body.index("If Not proofs.Exists(proof) Then"), body.index("MBI_ReadHeader"))

    def test_query_preserves_cross_presence_and_empty_id_groups(self):
        body = procedure(self.source("mb_identity.basfrag"), "MBI_QueryGroups")
        self.assertIn('"ABSENT", "EMPTY"', body)
        self.assertIn('"PRESENT", "EMPTY"', body)
        self.assertIn('Array("PRESENT", "ABSENT", "EMPTY")', body)

    def test_final_guard_and_content_check_remain_authoritative(self):
        body = procedure(self.source("mb_identity.basfrag"), "MBI_TryAlias")
        final = body.index("guard = MBH_GuardFile(matchedPath)")
        commit = body.index('MB_RecordCommit "R"')
        self.assertLess(final, commit)
        self.assertIn("MBI_CheckCandidate(mail, matchedKey", body[final:commit])
        self.assertIn("If Not MB_SourceStable(mail, q, revision)", body[final:commit])
        self.assertIn("checked >= MBI_MAX_CANDIDATES", body)


if __name__ == "__main__":
    unittest.main(verbosity=2)
