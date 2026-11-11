#!/usr/bin/env python3
"""Run the seven release model/source-contract suites in a temporary directory.

These are Python tests, not Outlook/VBA execution. Requires Python 3.10+ and no
third-party packages. Detailed output is printed on failure; success is summarized.
"""

from __future__ import annotations

from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile

# Keep importing the local build helper from creating checkout-side cache files.
sys.dont_write_bytecode = True
from build import ARTIFACT, DEVELOPMENT, assemble


SUITES = (
    ("v34_settings_models.py", (), 43),
    ("v34_revision_models.py", (), 35),
    ("v34_metrics_models.py", ("--integration",), 27),
    ("v34_fidelity_models.py", ("--integration",), 28),
    ("v34_recovery_merge_models.py", (), 44),
    ("v34_recovery_proof_models.py", (), 41),
    ("v34_policy_upgrade_models.py", (), 25),
)


def main() -> int:
    total = 0
    failed = 0
    with tempfile.TemporaryDirectory(prefix="outlookmailbackup-test-") as temporary:
        work = Path(temporary)
        generated = assemble(work)
        if not ARTIFACT.is_file() or generated != ARTIFACT.read_bytes():
            print("FAIL: shipped BAS does not match its source fragments.",
                  file=sys.stderr)
            return 1
        (work / "vba_parser_review").mkdir()
        for filename, arguments, expected in SUITES:
            shutil.copyfile(DEVELOPMENT / filename, work / filename)
            completed = subprocess.run(
                [sys.executable, filename, *arguments], cwd=work,
                stdin=subprocess.DEVNULL, stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, text=True, encoding="utf-8", errors="replace",
            )
            match = re.search(r"(?m)^Ran (\d+) tests? in ", completed.stdout)
            count = int(match.group(1)) if match else 0
            passed = completed.returncode == 0 and count == expected
            if passed:
                total += count
                print(f"PASS: {filename}: {count}/{expected}", flush=True)
            else:
                failed += 1
                print(f"FAIL: {filename}: exit={completed.returncode}, "
                      f"tests={count}, expected={expected}", file=sys.stderr)
                print(completed.stdout, file=sys.stderr)
    expected_total = sum(count for _, _, count in SUITES)
    print(f"Release models: {total}/{expected_total} passed; {failed} failed suites.")
    print("Outlook VBA compilation/runtime and real 100,000-mail benchmark: NOT_RUN.")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
