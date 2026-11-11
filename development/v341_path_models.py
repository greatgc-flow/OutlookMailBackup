"""Managed-path boundary fault models and source contracts for v3.4.1.

The schedules below check that a rejected path causes no later probe, lock, or
owner write. They model filesystem observations; they do not execute VBA,
Windows junctions, or a native Outlook run. Point-in-time path checks do not
claim an atomic defense against hostile namespace changes after a check.
"""

from dataclasses import dataclass, field
from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parent


class RejectedPath(RuntimeError):
    pass


@dataclass
class Namespace:
    exists: bool = True
    root_is_file: bool = False
    root_is_link: bool = False
    ancestor_is_link: bool = False
    lock_is_link: bool = False
    owner_exists: bool = False
    writes: list[str] = field(default_factory=list)

    def check(self, *, include_lock: bool = False) -> None:
        if self.root_is_link or self.ancestor_is_link:
            raise RejectedPath("Managed root traverses a reparse point")
        if include_lock and self.lock_is_link:
            raise RejectedPath("Run lock is a reparse point")

    def preflight(self) -> None:
        if self.root_is_file:
            raise RejectedPath("A file occupies the managed directory")
        # An unproven owner probe is not an authority to waive this check.
        self.check()
        self.writes.append("owned_probe_created_and_removed")

    def prepare(self, change=None) -> None:
        change = change or (lambda phase, namespace: None)
        change("after_selection", self)
        self.check()
        if not self.exists:
            self.writes.append("create_directory")
            self.exists = True
        change("after_directory", self)
        self.check(include_lock=True)
        self.writes.append("acquire_lock")
        change("after_lock", self)
        self.check()
        if not self.owner_exists:
            self.writes.append("create_owner")
            self.owner_exists = True


def replace_root_at(phase: str):
    def change(current: str, namespace: Namespace) -> None:
        if current == phase:
            namespace.root_is_link = True
    return change


def procedure(filename: str, name: str) -> str:
    source = (ROOT / filename).read_text(encoding="ascii")
    match = re.search(
        rf"(?ms)^(?:Private|Public) (?:Sub|Function) {re.escape(name)}\b.*?^End (?:Sub|Function)\s*$",
        source,
    )
    if match is None:
        raise AssertionError(f"Missing procedure: {filename}:{name}")
    # Source contracts must match executable statements, not explanatory comments.
    return "\n".join(line for line in match.group().splitlines()
                     if not line.lstrip().startswith("'"))


class PathBoundarySchedules(unittest.TestCase):
    def test_empty_managed_junction_cannot_trigger_write_probe(self):
        namespace = Namespace(root_is_link=True)
        with self.assertRaises(RejectedPath):
            namespace.preflight()
        self.assertEqual(namespace.writes, [])

    def test_owner_marker_does_not_authorize_managed_junction(self):
        namespace = Namespace(root_is_link=True, owner_exists=True)
        with self.assertRaises(RejectedPath):
            namespace.preflight()
        self.assertEqual(namespace.writes, [])

    def test_missing_child_below_linked_ancestor_is_rejected(self):
        namespace = Namespace(exists=False, ancestor_is_link=True)
        with self.assertRaises(RejectedPath):
            namespace.prepare()
        self.assertEqual(namespace.writes, [])

    def test_normal_absent_managed_root_can_be_initialized(self):
        namespace = Namespace(exists=False)
        namespace.prepare()
        self.assertEqual(namespace.writes, ["create_directory", "acquire_lock", "create_owner"])

    def test_normal_existing_empty_root_can_be_initialized(self):
        namespace = Namespace()
        namespace.prepare()
        self.assertEqual(namespace.writes, ["acquire_lock", "create_owner"])

    def test_existing_owned_root_is_not_recreated(self):
        namespace = Namespace(owner_exists=True)
        namespace.prepare()
        self.assertEqual(namespace.writes, ["acquire_lock"])

    def test_root_changed_after_preflight_is_rejected_before_initialization(self):
        namespace = Namespace()
        namespace.preflight()
        namespace.writes.clear()
        with self.assertRaises(RejectedPath):
            namespace.prepare(replace_root_at("after_selection"))
        self.assertEqual(namespace.writes, [])

    def test_root_changed_after_directory_creation_is_rejected_before_lock(self):
        namespace = Namespace(exists=False)
        with self.assertRaises(RejectedPath):
            namespace.prepare(replace_root_at("after_directory"))
        self.assertEqual(namespace.writes, ["create_directory"])
        self.assertFalse(namespace.owner_exists)

    def test_root_changed_after_lock_is_rejected_before_owner_write(self):
        namespace = Namespace()
        with self.assertRaises(RejectedPath):
            namespace.prepare(replace_root_at("after_lock"))
        self.assertEqual(namespace.writes, ["acquire_lock"])
        self.assertFalse(namespace.owner_exists)

    def test_lock_file_link_is_not_opened_as_the_run_lock(self):
        namespace = Namespace(lock_is_link=True)
        with self.assertRaises(RejectedPath):
            namespace.prepare()
        self.assertEqual(namespace.writes, [])

    def test_file_occupying_managed_root_is_rejected_before_probe(self):
        namespace = Namespace(root_is_file=True)
        with self.assertRaises(RejectedPath):
            namespace.preflight()
        self.assertEqual(namespace.writes, [])

    def test_normal_preflight_only_performs_its_owned_temporary_probe(self):
        namespace = Namespace()
        namespace.preflight()
        self.assertEqual(namespace.writes, ["owned_probe_created_and_removed"])
        self.assertFalse(namespace.owner_exists)


class PathSourceContracts(unittest.TestCase):
    def test_derived_managed_root_is_checked_before_store_probe_and_write_probe(self):
        body = procedure("mb_paths.basfrag", "MBP_TryBackupBase")
        derived = body.index("managed = MBH_JoinPath(base, MBP_MANAGED_LEAF)")
        guard = body.index("MBP_NoReparse managed, fso", derived)
        self.assertLess(guard, body.index("MBS_CheckCandidateStore managed"))
        self.assertLess(guard, body.index("MBP_ProbeWritable probeParent, fso"))

    def test_run_rechecks_selection_before_creating_managed_directory(self):
        body = procedure("mb_main.basfrag", "MB_Run")
        selected = body.index("baseRoot = MBP_ResolveBackupRoot(gSelectedBase)")
        guard = body.index("MBP_NoReparse baseRoot, gFSO", selected)
        self.assertLess(guard, body.index("MBH_EnsureDirectory baseRoot", selected))

    def test_run_rechecks_lock_path_after_directory_creation_before_opening_lock(self):
        body = procedure("mb_main.basfrag", "MB_Run")
        created = body.index("MBH_EnsureDirectory baseRoot")
        guard = body.index('MBP_NoReparse MBH_JoinPath(baseRoot, "backup.lock"), gFSO', created)
        self.assertLess(guard, body.index("MBH_AcquireRunLock", created))

    def test_run_rechecks_managed_root_under_lock_before_ownership_write(self):
        body = procedure("mb_main.basfrag", "MB_Run")
        locked = body.index("lockHeld = True")
        guard = body.index("MBP_NoReparse baseRoot, gFSO", locked)
        self.assertLess(guard, body.index("MB_CheckOwnership gRoot", locked))


if __name__ == "__main__":
    unittest.main(verbosity=2)
