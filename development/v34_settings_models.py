"""v3.4 deletion/location/ownership fault models and source contracts.

Behavioral models exercise data-preservation invariants independently of VBA.
The source checks connect those invariants to the edited procedures. Neither
these tests nor the ANTLR grammar run execute Outlook, HKCU, FSO, or Windows APIs.
"""
from __future__ import annotations

from dataclasses import dataclass, field
import hashlib
import itertools
import json
import ntpath
from pathlib import Path
import re
import unittest

HERE = Path(__file__).resolve().parent
BASE = r"D:\MailArchive"
ROOT = BASE + r"\SafeV2"
ACCOUNT = r"E:\HistoricalMail\CustomAccount"
OTHER_BASE = r"C:\UnrelatedDefault"
OTHER_ROOT = OTHER_BASE + r"\SafeV2"
STORE = "store-a"


def source(name):
    return (HERE / name).read_text(encoding="ascii")


def proc(text, name):
    start = re.search(r"(?m)^(?:Private|Public) (?:Function|Sub) " + re.escape(name) + r"\b", text)
    if not start:
        raise AssertionError(name)
    end = re.search(r"(?m)^End (?:Sub|Function)", text[start.end():])
    return text[start.start():start.end() + end.end()]


def norm(path):
    if not ntpath.isabs(path) or path.startswith(("\\\\?\\", "\\\\.\\")):
        raise ValueError("absolute ordinary path required")
    return ntpath.normpath(path).casefold()


def within(path, base):
    path, base = norm(path), norm(base)
    return path == base or path.startswith(base.rstrip("\\") + "\\")


def escape(value):
    value = value.replace("%", "%25").replace("\t", "%09").replace("\r", "%0D").replace("\n", "%0A")
    if value[:1] in {"=", "+", "-", "@"}:
        value = {"=": "%3D", "+": "%2B", "-": "%2D", "@": "%40"}[value[0]] + value[1:]
    return value


def unescape(value):
    for encoded, plain in [("%09", "\t"), ("%0D", "\r"), ("%0A", "\n"), ("%3D", "="),
                           ("%2B", "+"), ("%2D", "-"), ("%40", "@"), ("%25", "%")]:
        value = value.replace(encoded, plain)
    return value


def checked(fields):
    payload = "\t".join(escape(v) for v in fields)
    digest = hashlib.sha256(payload.encode("utf-16-le")).hexdigest()
    return (payload + "\t" + digest + "\tEND\r\n").encode("utf-16")


def read_checked(data, count):
    if len(data) % 2:
        raise ValueError("torn UTF-16")
    text = data.decode("utf-16")
    lines = text.splitlines()
    if len(lines) != 1:
        raise ValueError("extra/empty rows")
    raw = lines[0].split("\t")
    if len(raw) != count + 2 or raw[-1] != "END":
        raise ValueError("framing")
    payload = "\t".join(raw[:-2])
    if hashlib.sha256(payload.encode("utf-16-le")).hexdigest() != raw[-2]:
        raise ValueError("checksum")
    return [unescape(v) for v in raw[:-2]]


@dataclass(frozen=True)
class Location:
    store: str = STORE
    base: str = BASE
    managed: str = ROOT
    account: str = ACCOUNT

    def encode(self):
        return checked(["MBBASELOCATION1", self.store, self.base, self.managed, self.account])

    @classmethod
    def decode(cls, data, expected_store=STORE):
        row = read_checked(data, 5)
        if row[0] != "MBBASELOCATION1" or row[1].casefold() != expected_store.casefold():
            raise ValueError("schema/StoreID")
        if norm(row[3]) not in {norm(row[2]), norm(ntpath.join(row[2], "SafeV2"))}:
            raise ValueError("base/root mismatch")
        if row[4]:
            norm(row[4])
        return cls(*row[1:])


@dataclass
class RootFact:
    base: str
    managed: str
    store: str = STORE
    available: bool = True
    corrupt_binding: bool = False
    reparse: bool = False
    durable_binding: bool = True
    old_evidence: tuple[str | None, ...] = ()


class LocationResolver:
    def __init__(self, facts):
        self.facts = {norm(f.base): f for f in facts}
        self.probes = []

    def usable(self, path, store=STORE):
        fact = self.facts.get(norm(path))
        if fact is None or not fact.available:
            return None, "OFFLINE"
        if fact.reparse or fact.corrupt_binding:
            return None, "REVIEW"
        if fact.durable_binding:
            if fact.store.casefold() != store.casefold():
                return None, "FOREIGN_STORE"
        else:
            for identity in fact.old_evidence:
                if identity is None:
                    return None, "REVIEW"
                if identity.casefold() != store.casefold():
                    return None, "FOREIGN_STORE"
        self.probes.append(fact.managed)
        return fact.managed, ""

    def resolve(self, registry=None, locator=None, store=STORE):
        # Missing settings never create evidence from a default path.
        location = None
        if locator is not None:
            try:
                location = Location.decode(locator, store)
            except (ValueError, UnicodeError):
                pass
        found = None
        if location:
            found, _ = self.usable(location.base, store)
            if found and norm(found) != norm(location.managed):
                found = None
        saved = None
        conflict = False
        if registry:
            saved, reason = self.usable(registry, store)
            if location:
                if saved:
                    conflict = norm(saved) != norm(location.managed)
                elif reason != "FOREIGN_STORE":
                    conflict = norm(registry) not in {norm(location.base), norm(location.managed)}
        if conflict:
            return "PICK", None
        if found:
            return "REUSE", location.base
        if saved and location is None:
            return "REUSE", registry
        return "PICK", None


def digest(data):
    return hashlib.sha256(data).hexdigest() if data is not None else ""


@dataclass
class LocatorVault:
    current: bytes | None = None
    preserved: dict[str, bytes] = field(default_factory=dict)
    stage: bytes | None = None
    locked: bool = False
    events: list[str] = field(default_factory=list)
    msg: bytes = b"existing MSG sentinel"

    def save(self, observed, new, crash=None, preservation_corrupt=False, publish_corrupt=False):
        if self.locked:
            raise OSError("busy")
        self.locked = True
        try:
            self.events.append("LOCK")
            if digest(self.current) != observed:
                raise ValueError("changed by another run")
            self.stage = new
            self.events.append("FLUSH_STAGE")
            if crash == "STAGE":
                raise InterruptedError("stage")
            if self.current == self.stage:
                self.stage = None
                return
            if self.current is not None:
                key = digest(self.current)
                self.preserved.setdefault(key, self.current)
                self.events.append("PRESERVE_OLD")
                if preservation_corrupt:
                    self.preserved[key] = b"torn preservation"
                if digest(self.preserved[key]) != key:
                    raise ValueError("preservation hash differs")
                self.events.append("VERIFY_PRESERVED")
            if crash == "PRESERVED":
                raise InterruptedError("preserved")
            self.current = self.stage
            self.stage = None
            self.events.append("PUBLISH")
            if publish_corrupt:
                self.current = b"damaged publication"
            if crash == "PUBLISHED":
                raise InterruptedError("published")
            if digest(self.current) != digest(new):
                raise ValueError("published hash differs")
            self.events.append("VERIFY_PUBLISHED")
        finally:
            self.locked = False


def migrate(active, old, activation_file=False, activation_hkcu=False):
    # Any pre-existing active evidence is authoritative against old migration.
    prior = activation_file or activation_hkcu or bool(active)
    result = dict(active)
    if not prior:
        for name, data in old.items():
            if name.lower().endswith(".tsv"):
                result.setdefault(name, data)
    return result, not prior and bool(old)


@dataclass(frozen=True)
class AccountFact:
    path: str
    marker_store: str | None = STORE
    marker_root: str = ROOT
    recovery_store: str | None = STORE
    recovery_root: str = ROOT
    available: bool = True
    corrupt: bool = False
    reparse: bool = False

    def usable(self):
        if not self.available or self.corrupt or self.reparse or within(self.path, ROOT):
            return False
        if self.marker_store is not None:
            if self.marker_store.casefold() != STORE or norm(self.marker_root) != norm(ROOT):
                return False
        if self.recovery_store is not None:
            if self.recovery_store.casefold() != STORE or norm(self.recovery_root) != norm(ROOT):
                return False
        return True

    def exact(self):
        return self.marker_store is not None or self.recovery_store is not None


def select_account(remembered, discoverable, facts, name_candidates=()):
    by_path = {norm(f.path): f for f in facts}
    candidates = {norm(path) for path in remembered if norm(path) in by_path and by_path[norm(path)].usable()}
    if not candidates:
        candidates = {norm(path) for path in discoverable if norm(path) in by_path
                      and by_path[norm(path)].usable() and by_path[norm(path)].exact()}
    if len(candidates) == 1:
        return "REUSE", next(iter(candidates))
    if not candidates and not remembered:
        names = {norm(p) for p in name_candidates if norm(p) in by_path and by_path[norm(p)].usable()}
        if len(names) == 1:
            return "REUSE", next(iter(names))
    return "PICK", None


class SettingsBehaviorModels(unittest.TestCase):
    def test_deleted_registry_reuses_independent_nondefault_location(self):
        resolver = LocationResolver([RootFact(BASE, ROOT), RootFact(OTHER_BASE, OTHER_ROOT)])
        self.assertEqual(resolver.resolve(locator=Location().encode()), ("REUSE", BASE))

    def test_all_base_references_deleted_never_choose_existing_hardcoded_default(self):
        resolver = LocationResolver([RootFact(OTHER_BASE, OTHER_ROOT)])
        self.assertEqual(resolver.resolve(), ("PICK", None))
        self.assertEqual(resolver.probes, [])

    def test_pre34_registry_only_migrates_to_locator_without_a_chooser(self):
        resolver = LocationResolver([RootFact(BASE, ROOT)])
        self.assertEqual(resolver.resolve(registry=BASE), ("REUSE", BASE))

    def test_deleted_locator_reuses_registry_and_regenerates_checked_copy(self):
        resolver = LocationResolver([RootFact(BASE, ROOT)])
        self.assertEqual(resolver.resolve(registry=BASE), ("REUSE", BASE))
        vault = LocatorVault()
        vault.save("", Location().encode())
        self.assertEqual(Location.decode(vault.current), Location())

    def test_conflicting_valid_roots_ask_once_instead_of_choosing_arbitrarily(self):
        resolver = LocationResolver([RootFact(BASE, ROOT), RootFact(OTHER_BASE, OTHER_ROOT)])
        self.assertEqual(resolver.resolve(OTHER_BASE, Location().encode()), ("PICK", None))
        vault = LocatorVault(Location().encode())
        selected = Location(base=OTHER_BASE, managed=OTHER_ROOT, account="")
        vault.save(digest(vault.current), selected.encode())
        self.assertEqual(resolver.resolve(OTHER_BASE, vault.current), ("REUSE", OTHER_BASE))

    def test_offline_conflicting_reference_is_not_silently_forgotten(self):
        resolver = LocationResolver([RootFact(BASE, ROOT), RootFact(OTHER_BASE, OTHER_ROOT, available=False)])
        self.assertEqual(resolver.resolve(OTHER_BASE, Location().encode()), ("PICK", None))

    def test_foreign_global_registry_does_not_override_exact_per_store_locator(self):
        resolver = LocationResolver([RootFact(BASE, ROOT), RootFact(OTHER_BASE, OTHER_ROOT, store="store-b")])
        self.assertEqual(resolver.resolve(OTHER_BASE, Location().encode()), ("REUSE", BASE))

    def test_foreign_root_is_rejected_before_write_probe(self):
        resolver = LocationResolver([RootFact(OTHER_BASE, OTHER_ROOT, store="store-b")])
        self.assertEqual(resolver.resolve(OTHER_BASE), ("PICK", None))
        self.assertEqual(resolver.probes, [])

    def test_foreign_old_root_without_store_markers_returns_chooser_before_probe(self):
        resolver = LocationResolver([RootFact(BASE, ROOT, durable_binding=False, old_evidence=(STORE, "store-b"))])
        self.assertEqual(resolver.resolve(BASE), ("PICK", None))
        self.assertEqual(resolver.probes, [])

    def test_unreadable_owned_old_evidence_cannot_authorize_a_fresh_store_binding(self):
        for rows in ((None,), (STORE, None), (None, "store-b")):
            with self.subTest(rows=rows):
                resolver = LocationResolver([RootFact(BASE, ROOT, durable_binding=False, old_evidence=rows)])
                self.assertEqual(resolver.resolve(BASE), ("PICK", None))
                self.assertEqual(resolver.probes, [])

    def test_valid_exact_old_store_evidence_allows_upgrade_without_a_folder_prompt(self):
        resolver = LocationResolver([RootFact(BASE, ROOT, durable_binding=False, old_evidence=(STORE, STORE.upper()))])
        self.assertEqual(resolver.resolve(BASE), ("REUSE", BASE))

    def test_surviving_current_store_binding_keeps_fast_path_over_stale_old_history(self):
        resolver = LocationResolver([RootFact(BASE, ROOT, old_evidence=(None, "store-b"))])
        self.assertEqual(resolver.resolve(BASE), ("REUSE", BASE))

    def test_corrupt_locator_is_preserved_then_repaired_from_unique_registry_binding(self):
        damaged = Location().encode()[:-1]
        resolver = LocationResolver([RootFact(BASE, ROOT)])
        self.assertEqual(resolver.resolve(BASE, damaged), ("REUSE", BASE))
        vault = LocatorVault(damaged)
        vault.save(digest(damaged), Location().encode())
        self.assertEqual(vault.preserved[digest(damaged)], damaged)
        self.assertEqual(Location.decode(vault.current), Location())

    def test_corrupt_locator_without_any_other_binding_requires_picker(self):
        resolver = LocationResolver([RootFact(BASE, ROOT), RootFact(OTHER_BASE, OTHER_ROOT)])
        self.assertEqual(resolver.resolve(locator=b"corrupt"), ("PICK", None))

    def test_locator_checks_exact_store_root_relation_and_single_checked_row(self):
        valid = Location().encode()
        cases = [Location(store="store-b").encode(), Location(managed=OTHER_ROOT).encode(), valid[:-1],
                 valid + "\r\n".encode("utf-16-le"), valid.replace(b"E\x00N\x00D\x00", b"X\x00N\x00D\x00")]
        for data in cases:
            with self.subTest(data=data[:30]), self.assertRaises((ValueError, UnicodeError)):
                Location.decode(data)

    def test_locator_unicode_and_percent_paths_roundtrip_without_locale_dates(self):
        location = Location(base=r"D:\메일%09백업", managed=r"D:\메일%09백업\SafeV2", account=r"E:\保存\Compte")
        self.assertEqual(Location.decode(location.encode()), location)

    def test_locator_compare_before_replace_detects_concurrent_other_root_choice(self):
        vault = LocatorVault(Location().encode())
        observed = digest(vault.current)
        other = Location(base=OTHER_BASE, managed=OTHER_ROOT, account="").encode()
        vault.save(observed, other)
        preserved_before = dict(vault.preserved)
        with self.assertRaisesRegex(ValueError, "another run"):
            vault.save(observed, Location().encode())
        self.assertEqual(vault.current, other)
        self.assertEqual(vault.preserved, preserved_before)

    def test_locator_exclusive_lock_blocks_second_writer(self):
        old = Location().encode()
        vault = LocatorVault(old, locked=True)
        with self.assertRaisesRegex(OSError, "busy"):
            vault.save(digest(old), Location(account="").encode())
        self.assertEqual(vault.current, old)

    def test_locator_crash_points_keep_an_exact_old_or_new_copy(self):
        old = Location().encode()
        new = Location(account=r"E:\AnotherAccount").encode()
        for point in ("STAGE", "PRESERVED", "PUBLISHED"):
            with self.subTest(point=point):
                vault = LocatorVault(old)
                with self.assertRaises(InterruptedError):
                    vault.save(digest(old), new, crash=point)
                self.assertIn(vault.current, (old, new))
                if vault.current == new:
                    self.assertEqual(vault.preserved[digest(old)], old)
                self.assertEqual(vault.msg, b"existing MSG sentinel")

    def test_failed_preservation_never_publishes_new_locator(self):
        old = b"malformed but historically retained bytes"
        vault = LocatorVault(old)
        with self.assertRaisesRegex(ValueError, "preservation"):
            vault.save(digest(old), Location().encode(), preservation_corrupt=True)
        self.assertEqual(vault.current, old)
        self.assertNotIn("PUBLISH", vault.events)

    def test_failed_publication_check_keeps_original_bytes_in_preserved(self):
        old = Location().encode()
        vault = LocatorVault(old)
        with self.assertRaisesRegex(ValueError, "published"):
            vault.save(digest(old), Location(account="").encode(), publish_corrupt=True)
        self.assertEqual(vault.preserved[digest(old)], old)

    def test_unchanged_locator_run_leaves_no_extra_preserved_or_stage(self):
        old = Location().encode()
        vault = LocatorVault(old)
        vault.save(digest(old), old)
        self.assertEqual(vault.current, old)
        self.assertEqual(vault.preserved, {})
        self.assertIsNone(vault.stage)

    def test_any_surviving_active_evidence_blocks_stale_migration_across_settings_deletions(self):
        variants = [{"journal.tsv": b"newer verified L/R rows"}, {"runs": b"directory evidence"},
                    {"README.txt": b"prior activation"}, {"coverage.tsv": b"newer coverage"}]
        for active, anchor, hkcu, old_present in itertools.product(variants, (False, True), (False, True), (False, True)):
            old = {"journal.tsv": b"stale history", "moves.tsv": b"stale moves"} if old_present else {}
            result, copied = migrate(active, old, anchor, hkcu)
            with self.subTest(active=active, anchor=anchor, hkcu=hkcu, old=old_present):
                self.assertEqual(result, active)
                self.assertFalse(copied)

    def test_first_legacy_upgrade_without_active_evidence_copies_exact_tsv_only(self):
        old = {"journal.tsv": b"retained old proof", "notes.txt": b"unrelated"}
        result, copied = migrate({}, old)
        self.assertTrue(copied)
        self.assertEqual(result, {"journal.tsv": old["journal.tsv"]})
        self.assertEqual(old["journal.tsv"], b"retained old proof")

    def test_partial_migration_is_preserved_for_recovery_instead_of_reset_from_old(self):
        active = {"journal.tsv": b"partial copied file"}
        old = {"journal.tsv": b"old complete history", "coverage.tsv": b"old coverage"}
        self.assertEqual(migrate(active, old), (active, False))

    def test_external_account_survives_registry_and_map_deletion_via_locator(self):
        location = Location.decode(Location().encode())
        self.assertEqual(select_account([location.account], [], [AccountFact(ACCOUNT)]), ("REUSE", norm(ACCOUNT)))

    def test_unique_custom_sibling_account_discovered_from_exact_marker(self):
        custom = BASE + r"\ArchiveWithRenamedDisplayName"
        self.assertEqual(select_account([], [custom], [AccountFact(custom)]), ("REUSE", norm(custom)))

    def test_recovery_owner_alone_can_identify_exact_account(self):
        facts = [AccountFact(ACCOUNT, marker_store=None)]
        self.assertEqual(select_account([], [ACCOUNT], facts), ("REUSE", norm(ACCOUNT)))

    def test_foreign_or_corrupt_owner_never_overridden_by_locator(self):
        variants = [AccountFact(ACCOUNT, marker_store="store-b"), AccountFact(ACCOUNT, marker_root=OTHER_ROOT),
                    AccountFact(ACCOUNT, recovery_store="store-b"), AccountFact(ACCOUNT, recovery_root=OTHER_ROOT),
                    AccountFact(ACCOUNT, corrupt=True), AccountFact(ACCOUNT, reparse=True)]
        for fact in variants:
            with self.subTest(fact=fact):
                self.assertEqual(select_account([ACCOUNT], [ACCOUNT], [fact]), ("PICK", None))

    def test_unrelated_malformed_sibling_does_not_hide_unique_exact_account(self):
        bad = BASE + r"\BadSibling"
        self.assertEqual(select_account([], [bad, ACCOUNT], [AccountFact(bad, corrupt=True), AccountFact(ACCOUNT)]),
                         ("REUSE", norm(ACCOUNT)))

    def test_two_surviving_equal_tier_account_bindings_require_picker(self):
        second = BASE + r"\SecondAccount"
        facts = [AccountFact(ACCOUNT), AccountFact(second)]
        self.assertEqual(select_account([], [ACCOUNT, second], facts), ("PICK", None))
        self.assertEqual(select_account([ACCOUNT, second], [], facts), ("PICK", None))

    def test_agreed_current_account_outranks_older_leftover_marker(self):
        old = BASE + r"\PreviouslySelectedAccount"
        facts = [AccountFact(ACCOUNT), AccountFact(old)]
        self.assertEqual(select_account([ACCOUNT, ACCOUNT], [old, ACCOUNT], facts), ("REUSE", norm(ACCOUNT)))

    def test_deleted_both_account_markers_can_reuse_checked_remembered_path(self):
        fact = AccountFact(ACCOUNT, marker_store=None, recovery_store=None)
        self.assertEqual(select_account([ACCOUNT], [], [fact]), ("REUSE", norm(ACCOUNT)))
        self.assertEqual(select_account([], [ACCOUNT], [fact]), ("PICK", None))

    def test_offline_remembered_account_does_not_fall_back_to_unrelated_display_name(self):
        name_guess = BASE + r"\Display_RAW"
        facts = [AccountFact(ACCOUNT, available=False), AccountFact(name_guess, marker_store=None, recovery_store=None)]
        self.assertEqual(select_account([ACCOUNT], [], facts, [name_guess]), ("PICK", None))


class SettingsSourceContracts(unittest.TestCase):
    def test_migration_prior_guard_precedes_copy_and_has_no_active_delete(self):
        body = proc(source("mb_operating.basfrag"), "MBS_Initialize")
        self.assertLess(body.index(".Files.Count > 0"), body.index("If Not prior And gFSO.FolderExists(oldDir) Then"))
        self.assertIn(".SubFolders.Count > 0", body)
        self.assertNotIn("DeleteFile target", body)
        self.assertIn("gFSO.CopyFile CStr(f.Path), target, False", body)
        self.assertIn("MBH_HashFile(target) <> oldHash", body)

    def test_resolver_uses_no_hardcoded_fallback_and_handles_foreign_root(self):
        body = proc(source("mb_paths.basfrag"), "MBP_ResolveBackupRoot")
        self.assertIn('MBP_GetSettingText(MBP_BASE_KEY, "")', body)
        self.assertNotIn("MB_OUTPUT_ROOT", body)
        self.assertIn("MBP_ReadLocator", body)
        self.assertIn("[FOREIGN_STORE]", body)
        self.assertIn("If Not conflict Then", body)

    def test_locator_atomic_order_and_cas_preserve_metadata_before_publish(self):
        body = proc(source("mb_paths.basfrag"), "MBP_PersistLocator")
        order = ["handle = MBH_CreateFileW", "currentHash <> mbpLocatorHash", "MBM_FlushPath stage",
                 "fso.CopyFile mbpLocatorPath, preserved, False", "MBM_FlushPath preserved",
                 "MBH_HashFile(preserved) <> currentHash", "MBM_MoveFileExW", "MBM_FlushPath mbpLocatorPath"]
        self.assertEqual([body.index(x) for x in order], sorted(body.index(x) for x in order))
        self.assertIn("MBH_CloseHandle(handle)", body)
        self.assertNotIn("DeleteFile mbpLocatorPath", body)
        self.assertNotIn(".msg", body)

    def test_locator_schema_and_external_account_binding_are_checked_before_reuse(self):
        paths = source("mb_paths.basfrag")
        read = proc(paths, "MBP_ReadLocator")
        self.assertIn("MB_ReadCheckedLine(MBP_ReadOneLine(mbpLocatorPath, fso), 5)", read)
        self.assertIn("StrComp(CStr(row(1)), storeId", read)
        self.assertIn('MBH_JoinPath(base, MBP_MANAGED_LEAF)', read)
        self.assertIn("mbpLocatorAccount = MBH_NormalizePath(CStr(row(4)))", read)
        remember = proc(paths, "MBP_RememberLocatedAccount")
        self.assertLess(remember.index("MBY_CheckMarker account, storeId, True"), remember.index("MBP_PersistLocator"))
        account = proc(source("mb_operating.basfrag"), "MBS_EarlyAccount")
        self.assertIn("MBP_LocatedAccount(storeId, gBaseRoot)", account)
        self.assertIn("MBY_CheckRecoveryAccountBinding candidate, storeId", account)
        self.assertLess(account.index("MBY_CheckMarker candidate"), account.index("MBM_AcquireAccountLock candidate"))

    def test_discovery_prefers_exact_bindings_and_does_not_override_ambiguity(self):
        year = source("mb_yearmonth.basfrag")
        infer = proc(year, "MBY_InferAccountRoot")
        self.assertLess(infer.index("MBY_DiscoverBoundAccounts"), infer.index("folderName = MBL_OldSanitize"))
        self.assertIn("If candidates.Count > 1 Then Exit Function", infer)
        early = proc(source("mb_operating.basfrag"), "MBS_EarlyAccount")
        self.assertIn("If candidates.Count = 0 Then MBY_DiscoverBoundAccounts", early)
        self.assertIn("candidates.Count = 0 And remembered.Count = 0", early)

    def test_owner_recovery_accepts_checked_map_registry_and_external_marker_evidence(self):
        body = proc(source("mb_operating.basfrag"), "MBS_CanRecoverOwner")
        self.assertIn("MBS_OwnerEvidenceMap", body)
        self.assertIn('"OperatingStore_"', body)
        self.assertIn("MBP_OwnerAccountHint", body)
        hint = proc(source("mb_paths.basfrag"), "MBP_OwnerAccountHint")
        self.assertIn("MBP_LocatedAccount(storeId, managed)", hint)
        self.assertIn("MBS_AccountSettingKey(storeId, managed)", hint)
        self.assertIn('"ACCOUNT", fso, CStr(account)', body)
        self.assertIn('MBS_OwnerRecoveryEvidence(CStr(account), rootPath, fso)', body)
        recovery = proc(source("mb_operating.basfrag"), "MBS_OwnerRecoveryEvidence")
        self.assertIn('MBS_OwnerEvidenceFile(marker, rootPath, "RECOVERY", fso, accountPath)', recovery)
        self.assertIn('Array(MBR_FOLDER, MBR_FOLDER_LEGACY)', recovery)
        owner = proc(source("mb_paths.basfrag"), "MBP_ExactOwner")
        self.assertIn("MB_OWNER_PREVIOUS", owner)
        self.assertIn("An invalid owner marker already exists", owner)

    def test_root_store_preflight_precedes_probe_and_has_no_writes(self):
        gate = proc(source("mb_store_binding.basfrag"), "MBS_CheckCandidateStore")
        self.assertIn("rootPath", gate)
        self.assertIn("[FOREIGN_STORE]", gate)
        self.assertNotIn("SaveSetting", gate)
        self.assertNotIn("WriteAtomic", gate)
        body = proc(source("mb_paths.basfrag"), "MBP_TryBackupBase")
        self.assertLess(body.index("MBS_CheckCandidateStore"), body.index("MBP_ProbeWritable"))

    def test_old_root_preflight_uses_explicit_root_and_strict_owned_evidence(self):
        store = source("mb_store_binding.basfrag")
        gate = proc(store, "MBS_CheckCandidateStore")
        self.assertIn("If Not bound Then MBS_PreflightCandidateEvidence rootPath, storeId, fso", gate)
        preflight = proc(store, "MBS_PreflightCandidateEvidence")
        self.assertIn('Array(MBS_DATA_LEAF, MBS_DATA_LEAF_LEGACY, "_state")', preflight)
        self.assertIn('"JOURNAL", storeId, rootPath, fso, True', preflight)
        self.assertIn('"ACCOUNT", storeId, rootPath, fso, True', preflight)
        self.assertNotIn("gRoot =", preflight)
        self.assertNotIn("Write", preflight)
        row = proc(store, "MBS_CheckStoreRow")
        self.assertIn("If strictAsOwned Then Err.Raise", row)
        self.assertIn("GoTo Unproven", row)

    def test_main_persists_only_after_root_store_binding_and_resets_locator(self):
        main = source("mb_main.basfrag")
        body = proc(main, "MB_Run")
        self.assertLess(body.index("MBS_BindStore Application.Session.DefaultStore.StoreID"),
                        body.index("MBP_PersistBackupRoot gSelectedBase"))
        self.assertIn("MBP_Reset", proc(main, "MB_Reset"))


if __name__ == "__main__":
    suite = unittest.defaultTestLoader.loadTestsFromModule(__import__(__name__))
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    report = {"kind": __doc__, "tests": result.testsRun, "failures": len(result.failures), "errors": len(result.errors),
              "migration_deletion_combinations": 32,
              "not_run": ["VBA project compiler", "Windows API execution", "HKCU/FSO runtime", "Outlook COM"]}
    target = HERE / "vba_parser_review" / "v34_settings_models.json"
    target.write_text(json.dumps(report, indent=2), encoding="utf-8")
    print(json.dumps(report, indent=2))
    raise SystemExit(not result.wasSuccessful())
