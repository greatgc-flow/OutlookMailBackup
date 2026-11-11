#!/usr/bin/env python3
"""Run release model/source-contract suites in a temporary directory.

These are Python tests, not Outlook/VBA execution. Requires Python 3.10+ and no
third-party packages. Detailed output is printed on failure; success is summarized.
"""

from __future__ import annotations

import argparse
import hashlib
import json
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
    ("v341_path_models.py", (), 16),
    ("v341_index_alias_models.py", (), 41),
    ("v35_layout_models.py", (), 32),
    ("v35_relative_models.py", (), 48),
    ("v351_recovery_continuity_models.py", (), 39),
    ("v351_layout_binding_models.py", (), 57),
    ("v351_review_guidance_models.py", ("--integration",), 46),
    ("v351_selection_models.py", (), 36),
)


def suite_passed(returncode: int, output: str, expected: int) -> tuple[bool, int]:
    """An exit code/count alone can falsely count skipped tests as passes."""
    matches = re.findall(r"(?m)^Ran (\d+) tests? in ", output)
    count = int(matches[0]) if len(matches) == 1 else 0
    clean_success = len(re.findall(r"(?m)^OK\s*$", output)) == 1
    return returncode == 0 and count == expected and clean_success, count


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--json", action="store_true", help="emit a machine-readable result")
    args = parser.parse_args()
    total = 0
    failed = 0
    outcomes = []
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
            passed, count = suite_passed(completed.returncode, completed.stdout, expected)
            outcomes.append(dict(suite=filename, tests=count, expected=expected,
                                 passed=passed, exit_code=completed.returncode))
            if passed:
                total += count
                if not args.json:
                    print(f"PASS: {filename}: {count}/{expected}", flush=True)
            else:
                failed += 1
                print(f"FAIL: {filename}: exit={completed.returncode}, "
                      f"tests={count}, expected={expected}", file=sys.stderr)
                print(completed.stdout, file=sys.stderr)
    expected_total = sum(count for _, _, count in SUITES)
    if args.json:
        print(json.dumps(dict(check="Python models/source contracts, not VBA execution",
                              sha256=hashlib.sha256(generated).hexdigest(), suites=outcomes,
                              passed=total, expected=expected_total, failed_suites=failed,
                              outlook_runtime="NOT_RUN", real_100000_benchmark="NOT_RUN"), indent=2))
    else:
        print(f"Release models: {total}/{expected_total} passed; {failed} failed suites.")
        print("Outlook VBA compilation/runtime and real 100,000-mail benchmark: NOT_RUN.")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
