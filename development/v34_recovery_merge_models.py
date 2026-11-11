"""v3.4 recovery merge fault/behavior models and source contracts.

These are filesystem-backed Python state models, NOT an execution of Outlook,
VBA, Windows fsync/rename semantics, or a claim that native tests have passed.
"""
from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
from pathlib import Path
import re
import tempfile
import unittest

ROOT = Path(__file__).resolve().parent
BINDING = ("C:\\OwnedState", "C:\\OwnedMail", "store-one")
FULL_VALIDATION = "REOPENED_CONTENT_SHA256_V3"


class FrameError(Exception):
    pass


class Review(Exception):
    pass


class Crash(Exception):
    pass


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def text_hash(value: str) -> str:
    return digest(value.encode("utf-16-le"))


def checked(fields) -> str:
    payload = "\t".join(str(value) for value in fields)
    return payload + "\t" + text_hash(payload) + "\tEND"


def read_checked(line: str) -> list[str]:
    fields = line.split("\t")
    if len(fields) < 3 or fields[-1] != "END":
        raise FrameError("Incomplete framing")
    if fields[-2] != text_hash("\t".join(fields[:-2])):
        raise FrameError("Checksum mismatch")
    return fields[:-2]


def encoded(rows) -> bytes:
    text = "\r\n".join(rows)
    if text:
        text += "\r\n"
    return b"\xff\xfe" + text.encode("utf-16-le")


def decoded(content: bytes) -> list[str]:
    if len(content) % 2:
        raise Review("Odd UTF-16 byte count")
    try:
        return [line for line in content.decode("utf-16").splitlines() if line]
    except UnicodeError as error:
        raise Review("Invalid UTF-16") from error


def leaf_ok(name: str) -> bool:
    return bool(name) and name.endswith(".tsv") and not name.startswith((".", "~mb_")) and not any(char in name for char in "/\\:")


def source_key(store: str, entry: str) -> str:
    store, entry = store.lower(), entry.lower()
    return text_hash(f"MBSOURCE2|{len(store)}:{store}|{len(entry)}:{entry}")


def index_row(revision="revision-a", *, kind="N", body=b"mail", store=BINDING[2], entry="entry-one", reference="", path=None):
    key = text_hash(revision)
    fields = ["MBS3", kind, key, path or BINDING[1] + "\\2026\\11\\" + revision + ".msg", str(len(body)), digest(body),
              source_key(store, entry), store, entry, "Inbox", "2026-11-11T10:00:00", "2026-11-10T10:00:00",
              "Subject", "0", reference, FULL_VALIDATION, BINDING[1]]
    return checked(fields)


@dataclass
class IndexProof:
    byte_count: int
    digest: str
    source: str
    store: str
    path: str


class IndexState:
    """Independent expected-proof oracle; last accepted observation chooses head."""
    def __init__(self, expected_store=BINDING[2]):
        self.expected_store = expected_store
        self.proofs = {}
        self.heads = {}

    def accept(self, line):
        fields = read_checked(line)
        if len(fields) not in (16, 17):
            raise Review("Checked but unsupported field count")
        if fields[0] == "MBS3" and len(fields) == 17 and fields[1] in {"N", "T", "R"}:
            pass
        elif fields[0] == "MBS2" and len(fields) == 16 and fields[1] in {"M", "A", "L"}:
            pass
        else:
            raise Review("Checked but unsupported schema/kind")
        if fields[7].lower() != self.expected_store.lower():
            raise Review("Foreign StoreID cannot be published into this root")
        if source_key(fields[7], fields[8]) != fields[6]:
            raise Review("Source identity checksum mismatch")
        if not all(re.fullmatch("[0-9a-f]{64}", fields[index]) for index in (2, 5, 6)):
            raise Review("Invalid expected key/digest")
        if not fields[4].isdigit() or int(fields[4]) <= 0:
            raise Review("Invalid message byte count")
        kind = "M" if fields[1] in {"N", "R"} else ("A" if fields[1] == "T" else fields[1])
        key = kind + "|" + fields[2]
        proof = IndexProof(int(fields[4]), fields[5], fields[6], fields[7].lower(), fields[3])
        prior = self.proofs.get(key)
        if prior and (prior.byte_count, prior.digest, prior.source, prior.store) != (proof.byte_count, proof.digest, proof.source, proof.store):
            raise Review("Same immutable key has conflicting historical expected proof")
        if fields[1] == "R":
            reference = self.proofs.get(fields[14])
            if not reference or fields[14] == key:
                raise Review("R requires prior independent M/L proof")
            if (reference.byte_count, reference.digest, reference.store) != (proof.byte_count, proof.digest, proof.store):
                raise Review("R expected bytes/store differ from reference")
            if fields[15] != FULL_VALIDATION:
                raise Review("Unsupported R validation")
        self.proofs[key] = proof
        if kind != "A":
            self.heads[proof.source] = key


def merge_journal(base_bytes: bytes, current_bytes: bytes, *, expected_store=BINDING[2]):
    base, current = decoded(base_bytes), decoded(current_bytes)
    result, rejected = [], []
    state = IndexState(expected_store)
    for line in base:
        try:
            state.accept(line)
        except FrameError as error:
            raise Review("Invalid restored semantic base") from error
        result.append(line)
    # Remove only the common ORDERED prefix. A repeated later observation must
    # update the head even if the identical row occurred earlier in history.
    prefix = 0
    while prefix < min(len(base), len(current)) and base[prefix] == current[prefix]:
        prefix += 1
    for line in current[prefix:]:
        try:
            state.accept(line)
        except FrameError:
            rejected.append(line)
            continue
        result.append(line)
    return encoded(result), state, rejected


class Recovery:
    """File-backed durable-plan/restart model with explicit mutation cutpoints."""
    def __init__(self, folder: Path):
        self.folder = folder
        self.primary = folder / "primary"
        self.overlays = folder / "recovery" / "overlays"
        self.snapshots = folder / "recovery" / "snapshots"
        self.preserved = folder / "preserved"
        for path in (self.primary, self.overlays, self.snapshots, self.preserved):
            path.mkdir(parents=True, exist_ok=True)
        self.intent_path = folder / "recovery" / "intent.json"
        self.needs_path = folder / "recovery" / "needs.txt"
        self.events = []
        self.fail_at = None
        self.cache = {"sentinel": object()}

    def hit(self, label):
        self.events.append(label)
        if self.fail_at == label:
            raise Crash(label)

    def save_snapshot(self, files, name="snapshot_000000000001.tsv"):
        body = {"binding": BINDING, "files": {leaf: data.hex() for leaf, data in files.items()}}
        path = self.snapshots / name
        path.write_bytes(json.dumps(body, sort_keys=True).encode())
        return path

    def pending(self):
        if self.intent_path.exists():
            entry = json.loads(self.intent_path.read_text())
            if entry["phase"] == "START":
                return entry["pointer"]
        return None

    def write_intent(self, pointer, phase):
        self.intent_path.write_text(json.dumps({"pointer": pointer, "phase": phase}))

    def capture(self, base):
        number = len(list(self.overlays.glob("plan_*.tsv"))) + 1
        plan = self.overlays / f"plan_{number:04}.tsv"
        reference = str(base) if base is not None else "MSG_ONLY"
        base_digest = digest(base.read_bytes()) if base is not None else ""
        lines = [checked(["MBOVERLAY1", "HEADER", *BINDING, reference, base_digest, "2026-11-11T00:00:00Z"])]
        files = sorted(self.primary.glob("*.tsv"))
        for index, path in enumerate(files):
            name = f"raw_{number:04}_{index:04}.tsv"
            data = path.read_bytes()
            raw = self.overlays / name
            raw.write_bytes(data)
            if raw.read_bytes() != data:
                raise Review("Raw copy changed")
            lines.append(checked(["MBOVERLAY1", "FILE", path.name, name, len(data), digest(data), "", ""]))
        lines.append(checked(["MBOVERLAY1", "END", len(files), "", "", "", "", ""]))
        data = encoded(lines)
        plan.write_bytes(data)
        pointer = f"MERGE|{plan.name}|{len(data)}|{digest(data)}"
        self.hit("captured")
        return pointer

    def read_plan(self, pointer, binding=BINDING):
        parts = pointer.split("|")
        if len(parts) != 4 or parts[0] != "MERGE" or not leaf_ok(parts[1]):
            raise Review("Unsafe plan pointer")
        path = self.overlays / parts[1]
        try:
            data = path.read_bytes()
        except FileNotFoundError as error:
            raise Review("Plan disappeared") from error
        if len(data) != int(parts[2]) or digest(data) != parts[3]:
            raise Review("Plan digest changed")
        files = {}
        header = footer = False
        reference = base_digest = None
        for line in decoded(data):
            fields = read_checked(line)
            if len(fields) != 8 or fields[0] != "MBOVERLAY1" or footer:
                raise Review("Unknown/incomplete plan row")
            if fields[1] == "HEADER":
                if header or files or tuple(fields[2:5]) != tuple(binding):
                    raise Review("Duplicate/foreign plan header")
                reference, base_digest = fields[5:7]
                header = True
            elif fields[1] == "FILE":
                leaf, raw_name = fields[2:4]
                if not header or not leaf_ok(leaf) or not leaf_ok(raw_name) or leaf in files:
                    raise Review("Duplicate/unsafe file leaf")
                files[leaf] = (raw_name, int(fields[4]), fields[5])
            elif fields[1] == "END":
                if not header or int(fields[2]) != len(files):
                    raise Review("Wrong footer count")
                footer = True
            else:
                raise Review("Unknown plan record")
        if not header or not footer:
            raise Review("Plan is incomplete")
        if reference == "MSG_ONLY":
            if base_digest:
                raise Review("Unexpected MSG-only digest")
            base_files = {}
        else:
            base = Path(reference)
            if base.parent != self.snapshots or digest(base.read_bytes()) != base_digest:
                raise Review("Original base changed")
            payload = json.loads(base.read_text())
            if tuple(payload["binding"]) != tuple(binding):
                raise Review("Foreign snapshot")
            base_files = {leaf: bytes.fromhex(data) for leaf, data in payload["files"].items()}
            if any(not leaf_ok(leaf) for leaf in base_files):
                raise Review("Unsafe base filename")
        return files, base_files

    def replace_primary(self, leaf, data, phase):
        target = self.primary / leaf
        if target.exists() and target.read_bytes() == data:
            return
        if target.exists():
            saved = self.preserved / f"{len(list(self.preserved.iterdir())):05}_{leaf}"
            target.rename(saved)
            self.hit(f"{phase}:{leaf}:preserved")
        self.hit(f"{phase}:{leaf}:before_publish")
        target.write_bytes(data)
        self.hit(f"{phase}:{leaf}:published")

    def restore(self, latest_base, *, fail_at=None, fail_generic=None):
        self.fail_at = fail_at
        self.needs_path.write_text("NEEDS")
        pointer = self.pending()
        if pointer is None:
            pointer = self.capture(latest_base)
        files, base_files = self.read_plan(pointer)
        self.write_intent(pointer, "START")
        self.hit("started")
        for path in sorted(self.primary.glob("*.tsv")):
            if path.name not in base_files:
                saved = self.preserved / f"{len(list(self.preserved.iterdir())):05}_{path.name}"
                path.rename(saved)
                self.hit(f"extra:{path.name}:preserved")
        for leaf, data in sorted(base_files.items()):
            self.replace_primary(leaf, data, "base")
        if not (self.primary / "journal.tsv").exists():
            (self.primary / "journal.tsv").write_bytes(encoded([]))
        order = sorted(files, key=lambda leaf: (0 if leaf == "moves.tsv" else 1 if leaf == "journal.tsv" else 2, leaf))
        for leaf in order:
            name, size, expected = files[leaf]
            raw = (self.overlays / name).read_bytes()
            if len(raw) != size or digest(raw) != expected:
                raise Review("Immutable raw evidence changed")
            target = self.primary / leaf
            base = target.read_bytes() if target.exists() else encoded([])
            if leaf == "journal.tsv":
                original_cache = self.cache
                self.cache = {}
                try:
                    merged, state, rejected = merge_journal(base, raw)
                    self.cache.update(state.proofs)
                finally:
                    self.cache = original_cache
            else:
                if leaf == fail_generic:
                    raise Review("Injected semantic nonjournal proof conflict")
                # Other proofs have a separate schema-aware VBA merger. Here the
                # raw copy stands for its successful result; faults test ordering.
                merged = raw
            self.replace_primary(leaf, merged, "merge")
        self.hit("merged")
        self.write_intent(pointer, "DONE")
        self.hit("done")
        return pointer


def historical_stage_expectations(rows):
    latest, expected = {}, set()
    for stage, sequence, phase, path, sha in rows:
        if phase in {"PUBLISHED", "RECOVERED"}:
            expected.add((path, sha))
        if stage not in latest or sequence > latest[stage][0]:
            latest[stage] = (sequence, phase, path, sha)
    for sequence, phase, path, sha in latest.values():
        if phase == "PREPARED":
            expected.add((path, sha))
    return expected


def legacy_conversion_allowed(pending, preserved_names):
    return not pending or pending.startswith("MERGE|") or not any(name.lower().endswith(".tsv") for name in preserved_names)


class JournalMergeCases(unittest.TestCase):
    def test_newer_committed_expected_digest_survives_old_snapshot(self):
        a, b = index_row(), index_row("revision-b", body=b"changed")
        merged, state, bad = merge_journal(encoded([a]), encoded([a, b]))
        self.assertEqual(decoded(merged), [a, b])
        self.assertEqual(len(state.proofs), 2)
        self.assertFalse(bad)

    def test_r_alias_can_depend_on_restored_base(self):
        a = index_row()
        alias = index_row("alias", kind="R", reference="M|" + text_hash("revision-a"))
        merged, state, _ = merge_journal(encoded([a]), encoded([a, alias]))
        self.assertEqual(len(state.proofs), 2)
        self.assertEqual(decoded(merged), [a, alias])

    def test_r_alias_can_depend_on_earlier_new_overlay_row(self):
        a, b = index_row(), index_row("new-original", body=b"new")
        alias = index_row("alias", kind="R", body=b"new", reference="M|" + text_hash("new-original"))
        _, state, _ = merge_journal(encoded([a]), encoded([a, b, alias]))
        self.assertEqual(len(state.proofs), 3)

    def test_orphan_r_stops_instead_of_being_dropped_as_corrupt(self):
        orphan = index_row("alias", kind="R", reference="M|" + text_hash("missing"))
        with self.assertRaises(Review):
            merge_journal(encoded([]), encoded([orphan]))

    def test_same_key_different_hash_cannot_replace_old_expected_bytes(self):
        with self.assertRaises(Review):
            merge_journal(encoded([index_row()]), encoded([index_row(body=b"edited bytes")]))

    def test_same_key_different_source_cannot_rebind_proof(self):
        with self.assertRaises(Review):
            merge_journal(encoded([index_row()]), encoded([index_row(entry="another-entry")]))

    def test_foreign_store_is_not_published_even_with_valid_source_checksum(self):
        with self.assertRaises(Review):
            merge_journal(encoded([]), encoded([index_row(store="foreign-store")]))

    def test_later_duplicate_row_is_a_real_later_observation(self):
        a, b = index_row(), index_row("revision-b", body=b"changed")
        merged, state, _ = merge_journal(encoded([a, b]), encoded([a, b, a]))
        self.assertEqual(decoded(merged), [a, b, a])
        self.assertEqual(state.heads[source_key(BINDING[2], "entry-one")], "M|" + text_hash("revision-a"))

    def test_shorter_true_prefix_does_not_roll_back_known_base_tail(self):
        a, b = index_row(), index_row("revision-b", body=b"changed")
        merged, state, _ = merge_journal(encoded([a, b]), encoded([a]))
        self.assertEqual(decoded(merged), [a, b])
        self.assertEqual(state.heads[source_key(BINDING[2], "entry-one")], "M|" + text_hash("revision-b"))

    def test_only_bad_framing_can_be_excluded_with_raw_retained(self):
        a, b = index_row(), index_row("revision-b", body=b"changed")
        broken = a[:-3]
        raw = encoded([a, broken, b])
        merged, state, rejected = merge_journal(encoded([a]), raw)
        self.assertEqual(decoded(merged), [a, b])
        self.assertEqual(rejected, [broken])
        self.assertIn(broken, decoded(raw))

    def test_fully_checked_unknown_schema_kind_or_shape_is_fatal_review(self):
        fields = read_checked(index_row())
        for modification in (lambda f: ["MBS4"] + f[1:], lambda f: f[:1] + ["X"] + f[2:], lambda f: f[:-1]):
            with self.subTest(modification=modification):
                row = checked(modification(fields.copy()))
                with self.assertRaises(Review):
                    merge_journal(encoded([]), encoded([row]))

    def test_odd_bytes_are_not_silently_truncated(self):
        with self.assertRaises(Review):
            merge_journal(encoded([]), encoded([index_row()]) + b"\x00")

    def test_malformed_restored_base_is_not_salvaged_silently(self):
        with self.assertRaises(Review):
            merge_journal(encoded(["incomplete base"]), encoded([index_row()]))


class DurableRecoveryCases(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary.name)
        self.recovery = Recovery(self.root)
        self.a, self.b = index_row(), index_row("revision-b", body=b"changed")
        self.base = self.recovery.save_snapshot({"journal.tsv": encoded([self.a]), "settings.tsv": encoded([checked(["SETTING", "old"])])})
        (self.recovery.primary / "journal.tsv").write_bytes(encoded([self.a, self.b]))
        self.mail = self.root / "archive.msg"
        self.mail.write_bytes(b"immutable message payload")

    def tearDown(self):
        self.temporary.cleanup()

    def test_ancillary_missing_setting_does_not_erase_newer_journal(self):
        self.recovery.restore(self.base)
        self.assertEqual(decoded((self.recovery.primary / "journal.tsv").read_bytes()), [self.a, self.b])
        self.assertTrue((self.recovery.primary / "settings.tsv").exists())
        self.assertEqual(self.mail.read_bytes(), b"immutable message payload")
        self.assertEqual(self.recovery.needs_path.read_text(), "NEEDS")

    def test_every_replacement_cut_restarts_from_original_pointer(self):
        cuts = ["started", "base:journal.tsv:preserved", "base:journal.tsv:before_publish", "base:journal.tsv:published",
                "base:settings.tsv:before_publish", "merge:journal.tsv:preserved", "merge:journal.tsv:before_publish",
                "merge:journal.tsv:published", "merged"]
        for cut in cuts:
            with self.subTest(cut=cut), tempfile.TemporaryDirectory() as folder:
                r = Recovery(Path(folder))
                base = r.save_snapshot({"journal.tsv": encoded([self.a]), "settings.tsv": encoded([checked(["SETTING", "old"])])})
                (r.primary / "journal.tsv").write_bytes(encoded([self.a, self.b]))
                with self.assertRaises(Crash):
                    r.restore(base, fail_at=cut)
                pointer = r.pending()
                self.assertIsNotNone(pointer)
                plan_count = len(list(r.overlays.glob("plan_*.tsv")))
                resumed = Recovery(Path(folder))
                # A newer snapshot discovered later must not substitute the base.
                later = resumed.save_snapshot({"journal.tsv": encoded([])}, "snapshot_999999999999.tsv")
                self.assertEqual(resumed.restore(later), pointer)
                self.assertEqual(len(list(resumed.overlays.glob("plan_*.tsv"))), plan_count)
                self.assertEqual(decoded((resumed.primary / "journal.tsv").read_bytes()), [self.a, self.b])

    def test_capture_failure_before_start_does_not_replace_primary(self):
        before = (self.recovery.primary / "journal.tsv").read_bytes()
        with self.assertRaises(Crash):
            self.recovery.restore(self.base, fail_at="captured")
        self.assertIsNone(self.recovery.pending())
        self.assertEqual((self.recovery.primary / "journal.tsv").read_bytes(), before)

    def test_changed_plan_digest_blocks_resume(self):
        with self.assertRaises(Crash):
            self.recovery.restore(self.base, fail_at="started")
        pointer = self.recovery.pending()
        plan = self.recovery.overlays / pointer.split("|")[1]
        plan.write_bytes(plan.read_bytes() + b"\x00\x00")
        with self.assertRaises(Review):
            self.recovery.restore(self.base)
        self.assertEqual(self.recovery.pending(), pointer)

    def test_changed_raw_copy_blocks_done_and_keeps_original_intent(self):
        with self.assertRaises(Crash):
            self.recovery.restore(self.base, fail_at="started")
        pointer = self.recovery.pending()
        raw = next(self.recovery.overlays.glob("raw_*.tsv"))
        raw.write_bytes(encoded([]))
        with self.assertRaises(Review):
            self.recovery.restore(self.base)
        self.assertEqual(self.recovery.pending(), pointer)

    def test_snapshot_manifest_replacement_cannot_change_bound_base(self):
        with self.assertRaises(Crash):
            self.recovery.restore(self.base, fail_at="started")
        self.base.write_text("changed base")
        with self.assertRaises(Review):
            self.recovery.restore(self.base)

    def test_manifest_requires_exact_binding_and_safe_leaf_names(self):
        pointer = self.recovery.capture(self.base)
        with self.assertRaises(Review):
            self.recovery.read_plan(pointer, binding=(BINDING[0], BINDING[1], "other-store"))
        for unsafe in ("../bad.tsv", "x\\bad.tsv", "C:bad.tsv", ".hidden.tsv", "~mb_owned.tsv"):
            self.assertFalse(leaf_ok(unsafe))

    def test_checked_incomplete_manifest_is_not_completed_plan(self):
        pointer = self.recovery.capture(self.base)
        plan = self.recovery.overlays / pointer.split("|")[1]
        rows = decoded(plan.read_bytes())[:-1]
        changed = encoded(rows)
        plan.write_bytes(changed)
        forged_pointer = f"MERGE|{plan.name}|{len(changed)}|{digest(changed)}"
        with self.assertRaises(Review):
            self.recovery.read_plan(forged_pointer)

    def test_generic_merge_failure_after_journal_publish_never_writes_done(self):
        (self.recovery.primary / "stages.tsv").write_bytes(encoded([checked(["stage-proof", "newer"])]))
        with self.assertRaises(Review):
            self.recovery.restore(self.base, fail_generic="stages.tsv")
        pointer = self.recovery.pending()
        self.assertIsNotNone(pointer)
        self.assertEqual(decoded((self.recovery.primary / "journal.tsv").read_bytes()), [self.a, self.b])
        self.assertEqual(self.recovery.restore(self.base), pointer)
        self.assertIsNone(self.recovery.pending())
        self.assertTrue((self.recovery.primary / "stages.tsv").exists())

    def test_cache_object_identity_survives_success_and_conflict(self):
        cache = self.recovery.cache
        sentinel = cache["sentinel"]
        self.recovery.restore(self.base)
        self.assertIs(self.recovery.cache, cache)
        self.assertIs(self.recovery.cache["sentinel"], sentinel)
        self.assertEqual(set(self.recovery.cache), {"sentinel"})
        (self.recovery.primary / "journal.tsv").write_bytes(encoded([index_row(body=b"conflict")]))
        with self.assertRaises(Review):
            self.recovery.restore(self.base)
        self.assertIs(self.recovery.cache, cache)
        self.assertEqual(set(self.recovery.cache), {"sentinel"})

    def test_same_key_conflict_preserves_both_proof_sources_without_done(self):
        conflicting = encoded([index_row(body=b"conflict")])
        (self.recovery.primary / "journal.tsv").write_bytes(conflicting)
        with self.assertRaises(Review):
            self.recovery.restore(self.base)
        self.assertIsNotNone(self.recovery.pending())
        self.assertIn(conflicting, [path.read_bytes() for path in self.recovery.overlays.glob("raw_*.tsv")])
        self.assertIn(encoded([self.a]), [path.read_bytes() for path in self.recovery.primary.glob("journal.tsv")])

    def test_duplicate_later_head_observation_survives_repeat_restore(self):
        raw = encoded([self.a, self.b, self.a])
        (self.recovery.primary / "journal.tsv").write_bytes(raw)
        with self.assertRaises(Crash):
            self.recovery.restore(self.base, fail_at="merged")
        self.recovery.restore(self.base)
        merged = (self.recovery.primary / "journal.tsv").read_bytes()
        self.assertEqual(decoded(merged), [self.a, self.b, self.a])


class HistoricalStageCases(unittest.TestCase):
    def test_published_then_cleanup_aborted_retains_expected_digest(self):
        rows = [("stage", 1, "RESERVE", "mail.msg", ""), ("stage", 2, "PREPARED", "mail.msg", "digest"),
                ("stage", 3, "PUBLISHED", "mail.msg", "digest"), ("stage", 4, "ABORTED", "mail.msg", "")]
        self.assertEqual(historical_stage_expectations(rows), {("mail.msg", "digest")})

    def test_recovered_then_later_empty_phase_retains_expected_digest(self):
        self.assertEqual(historical_stage_expectations([("stage", 3, "RECOVERED", "mail.msg", "digest"),
                                                      ("stage", 4, "ABORTED", "mail.msg", "")]), {("mail.msg", "digest")})

    def test_never_published_prepared_then_aborted_is_not_immortal_missing_output(self):
        self.assertEqual(historical_stage_expectations([("stage", 2, "PREPARED", "mail.msg", "digest"),
                                                      ("stage", 3, "ABORTED", "mail.msg", "")]), set())

    def test_latest_prepared_still_retains_intent_for_recovery(self):
        self.assertEqual(historical_stage_expectations([("stage", 2, "PREPARED", "mail.msg", "digest")]), {("mail.msg", "digest")})

    def test_legacy_pending_restore_with_preserved_prior_proof_cannot_convert_silently(self):
        for pending in ("snapshot_older.tsv", "MSG_ONLY", "SALVAGE|old.tsv|2|digest|0"):
            self.assertFalse(legacy_conversion_allowed(pending, ["old_state/preserved/newer_journal.tsv"]))

    def test_new_merge_resume_is_not_blocked_by_its_own_preserved_partial_outputs(self):
        self.assertTrue(legacy_conversion_allowed("MERGE|plan.tsv|123|hash", ["preserved/old_journal.tsv"]))


class SourceContracts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.merge = (ROOT / "mb_recovery_merge.basfrag").read_text(encoding="ascii")
        cls.recovery = (ROOT / "mb_recovery.basfrag").read_text(encoding="ascii")
        cls.operating = (ROOT / "mb_operating.basfrag").read_text(encoding="ascii")

    def test_overlay_capture_is_immutable_verified_and_bound_before_start(self):
        capture = self.merge.split("Private Function MBR_CaptureOverlay", 1)[1].split("End Function", 1)[0]
        self.assertIn("gFSO.CopyFile source, copyPath, False", capture)
        self.assertIn("MBM_FlushPath copyPath", capture)
        self.assertIn("MBH_HashFile(copyPath) <> digest", capture)
        self.assertIn("MBH_CommitStage stage, manifest, digest, manifestBytes", capture)
        restore = self.recovery.split("Private Sub MBR_RestoreState", 1)[1].split("End Sub", 1)[0]
        self.assertLess(restore.index("MBR_CaptureOverlay"), restore.index('intent, "START"'))
        self.assertLess(restore.index('intent, "START"'), restore.index("MBR_PreserveExtraState"))

    def test_restart_reuses_pending_merge_pointer(self):
        restore = self.recovery.split("Private Sub MBR_RestoreState", 1)[1].split("End Sub", 1)[0]
        self.assertIn('If Left$(pending, 6) = "MERGE|" Then', restore)
        self.assertIn("intent = pending", restore)
        self.assertIn("MBR_ReadOverlay recoveryRoot, accountRoot, storeId, intent", restore)

    def test_done_follows_all_overlay_merges(self):
        restore = self.recovery.split("Private Sub MBR_RestoreState", 1)[1].split("End Sub", 1)[0]
        self.assertLess(restore.index("MBR_ApplyOverlay"), restore.index('intent, "DONE"'))
        self.assertNotIn("On Error Resume Next", restore)

    def test_conflict_guard_precedes_real_index_cache_mutation(self):
        merge_row = self.merge.split("Private Function MBR_TryMergeIndexRow", 1)[1].split("End Function", 1)[0]
        self.assertLess(merge_row.index("MBR_RejectIndexConflict"), merge_row.index("MBR_TryIndexRow"))
        self.assertIn("Err.Raise MB_ERR + 1154", merge_row)
        self.assertIn("n >= MB_ERR + 15 And n <= MB_ERR + 17", merge_row)

    def test_schema_shape_is_checked_after_actual_checked_framing(self):
        parser = self.merge.split("Private Function MBR_ParseIndexFields", 1)[1].split("End Function", 1)[0]
        self.assertIn("Split(line, vbTab)", parser)
        self.assertIn("UBound(", parser)
        self.assertIn("MB_ReadCheckedLine", parser)
        self.assertIn("Err.Raise MB_ERR + 1154", parser)
        self.assertLess(parser.index("MB_ReadCheckedLine"), parser.index("1154"))

    def test_only_ordered_prefix_is_skipped_not_global_duplicate_rows(self):
        merge = self.merge.split("Private Sub MBR_MergeJournal", 1)[1].split("End Sub", 1)[0]
        self.assertIn("matchingPrefix", merge)
        self.assertIn("position <= baseRows.Count", merge)
        self.assertIn("If Not skipPrefixRow Then matchingPrefix = False", merge)
        self.assertNotRegex(merge, r"(?i)\b(?:seenRows|uniqueRows)\b")

    def test_revision_and_identity_caches_are_restored_even_after_failure(self):
        merge = self.merge.split("Private Sub MBR_MergeJournal", 1)[1].split("End Sub", 1)[0]
        for current, saved in (("gIndex", "oldIndex"), ("gExternalRoots", "oldExternal"), ("gRecordMeta", "oldMeta"),
                               ("gStoreTokens", "oldStores"), ("gLegacyBindings", "oldBindings"), ("gLegacySources", "oldSources"),
                               ("gLegacyOwners", "oldOwners"), ("gIdentityBuckets", "oldBuckets"), ("gIdentitySeen", "oldSeen"),
                               ("gWindowKnown", "oldKnown"), ("gRevisionHeads", "oldHeads")):
            self.assertIn(f"Set {saved} = {current}", merge)
            self.assertIn(f"Set {current} = {saved}", merge.split("Cleanup:", 1)[1])

    def test_merge_only_rebuilds_metadata_and_does_not_export_mail(self):
        code = re.sub(r'"(?:[^"\r\n]|"")*"', '""', self.merge)
        code = "\n".join(line.split("'", 1)[0] for line in code.splitlines())
        self.assertNotRegex(code, r"(?i)\.\s*(?:SaveAs|Save|Send|OpenSharedItem|GetItemFromID)\b")

    def test_loaded_historical_published_rows_add_expected_proof_before_latest_selection(self):
        loader = self.operating.split("Private Sub MBS_LoadStages", 1)[1].split("End Sub", 1)[0]
        self.assertIn("gStageExpected", loader)
        expected = loader.index("gStageExpected")
        self.assertIn('"PUBLISHED"', loader[:expected])
        self.assertIn('"RECOVERED"', loader[:expected])
        self.assertLess(expected, loader.index("If gStageRecords.Exists(stage) Then"))

    def test_current_store_binding_is_checked_before_any_merge_cache_mutation(self):
        merge_row = self.merge.split("Private Function MBR_TryMergeIndexRow", 1)[1].split("End Function", 1)[0]
        self.assertIn("MBS_RequireStore CStr(f(7)), storeId", merge_row)
        self.assertLess(merge_row.index("MBS_RequireStore"), merge_row.index("MBR_RejectIndexConflict"))

    def test_original_base_manifest_is_guarded_through_exact_read(self):
        reader = self.merge.split("Private Sub MBR_LoadOverlayBase", 1)[1].split("End Sub", 1)[0]
        self.assertLess(reader.index("MBH_GuardFile(manifest)"), reader.index("MBH_HashFile(manifest) <> expectedDigest"))
        self.assertLess(reader.index("MBH_HashFile(manifest) <> expectedDigest"), reader.index("MBR_ReadSnapshot"))
        self.assertLess(reader.index("MBR_ReadSnapshot"), reader.index("MBH_CloseGuard guard"))

    def test_independent_stage_mirror_is_immutable_overlay_input_not_republished_as_a_setting(self):
        capture = self.merge.split("Private Function MBR_CaptureOverlay", 1)[1].split("End Function", 1)[0]
        self.assertIn('MBH_JoinPath(recoveryRoot, "stage-intents.tsv")', capture)
        self.assertIn("paths.Add Array(source, MBR_MIRROR_LEAF)", capture)
        apply = self.merge.split("Private Sub MBR_ApplyOverlay(", 1)[1].split("End Sub", 1)[0]
        self.assertIn("mirrorGuard = MBH_GuardFile(mirror)", apply)
        self.assertIn("leaf <> MBR_MIRROR_LEAF", apply)
        self.assertIn('MBR_MergeProofFile "stages.tsv", mirror, accountRoot, storeId', apply)

    def test_legacy_interrupted_restore_guard_precedes_new_overlay_capture(self):
        restore = self.recovery.split("Private Sub MBR_RestoreState", 1)[1].split("End Sub", 1)[0]
        self.assertIn("MBR_HasPreservedState()", restore)
        self.assertIn("Err.Raise MB_ERR + 1157", restore)
        self.assertLess(restore.index("MBR_HasPreservedState()"), restore.index("MBR_CaptureOverlay"))


if __name__ == "__main__":
    suite = unittest.TestSuite()
    for test in (JournalMergeCases, DurableRecoveryCases, HistoricalStageCases, SourceContracts):
        suite.addTests(unittest.TestLoader().loadTestsFromTestCase(test))
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    report = {
        "kind": "filesystem fault/behavior models and source contracts; NOT Outlook/VBA execution",
        "tests": result.testsRun, "failures": len(result.failures), "errors": len(result.errors),
        "pass": result.wasSuccessful(), "native_outlook_runtime": "NOT_RUN",
        "windows_atomic_durability_runtime": "NOT_RUN",
    }
    (ROOT / "vba_parser_review/v34_recovery_merge_models.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps(report, indent=2))
    raise SystemExit(not result.wasSuccessful())
