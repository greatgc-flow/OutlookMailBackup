#!/usr/bin/env python3
"""Run all repository verification gates; no Outlook automation is performed."""
from pathlib import Path
import subprocess
import sys

sys.dont_write_bytecode = True
from test import suite_passed

ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    gates = (
        ("tools/build.py", "--check"),
        ("tools/validate.py",),
        ("tools/release_evidence.py",),
        ("-m", "unittest", "discover", "-s", "development", "-p", "test_validation_tools.py"),
        ("tools/test.py",),
    )
    for gate in gates:
        print("VERIFY: " + " ".join(gate), flush=True)
        if gate[0] == "-m":
            completed = subprocess.run([sys.executable, "-B", *gate], cwd=ROOT,
                                       stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                                       stderr=subprocess.STDOUT, text=True, encoding="utf-8",
                                       errors="replace")
            print(completed.stdout, end="", flush=True)
            passed, count = suite_passed(completed.returncode, completed.stdout, 47)
            if not passed:
                print(f"FAIL: validator mutation suite must pass all 47 tests; observed {count}.")
                return 1
        else:
            completed = subprocess.run([sys.executable, "-B", *gate], cwd=ROOT,
                                       stdin=subprocess.DEVNULL)
        if completed.returncode:
            return completed.returncode
    print("Repository gates: PASS. Actual Outlook compilation/runtime: NOT_RUN.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
