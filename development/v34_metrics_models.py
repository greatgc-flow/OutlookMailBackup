"""v3.4 numerical/behavior models and source contracts; NOT Outlook/VBA execution.

Run --integration after the root-owned queue/reset/summary hooks are installed.
The bounded histogram model is compared with independent exact sorted samples.
No synthetic time generated here is a performance measurement of Outlook.
"""
from __future__ import annotations

from dataclasses import dataclass, field
import json
import math
from pathlib import Path
import random
import re
import sys
import unittest

ROOT = Path(__file__).resolve().parent
SOURCE = (ROOT / "mb_metrics.basfrag").read_text(encoding="ascii")
DECLARATIONS = (ROOT / "mb_metrics_declarations.basfrag").read_text(encoding="ascii")


def constant(name: str) -> float:
    match = re.search(rf"Const {name} As \w+ = ([0-9.]+)", DECLARATIONS)
    if not match:
        raise AssertionError(f"Missing VBA metric constant: {name}")
    return float(match.group(1))


GROUPS = int(constant("MBF_GROUP_COUNT"))
BINS = int(constant("MBF_BIN_COUNT"))
MIN_FORECAST = int(constant("MBF_MIN_FORECAST_COUNT"))
FIRST_UPPER = constant("MBF_FIRST_UPPER_SECONDS")
UPPERS = tuple(FIRST_UPPER * 2**index for index in range(BINS - 1))


def group_for(outcome: int, limited: bool = False) -> int:
    if outcome == 1 and limited:
        return 9
    if outcome == 2 and limited:
        return 10
    if 1 <= outcome <= 8:
        return outcome
    return 11 if outcome == 0 else 12


@dataclass
class Statistic:
    observed: int = 0
    timed: int = 0
    total: float = 0.0
    minimum: float = math.inf
    maximum: float = -math.inf
    zeros: int = 0
    bins: list[int] = field(default_factory=lambda: [0] * BINS)


class Metrics:
    """A bounded aggregate with no retained per-item durations."""

    def __init__(self) -> None:
        self.groups = {index: Statistic() for index in range(1, GROUPS + 1)}
        self.fault = False

    def record(self, mode, outcome, elapsed, limited=False):
        if mode != "AUTO":
            return
        stats = self.groups[group_for(outcome, limited)]
        stats.observed += 1
        if elapsed < 0:
            return
        stats.timed += 1
        stats.total += elapsed
        stats.minimum = min(stats.minimum, elapsed)
        stats.maximum = max(stats.maximum, elapsed)
        stats.zeros += elapsed == 0
        for index, upper in enumerate(UPPERS):
            if elapsed <= upper:
                stats.bins[index] += 1
                break
        else:
            stats.bins[-1] += 1

    def interval(self, group, fraction):
        stats = self.groups[group]
        if not stats.timed:
            return None
        target = math.ceil(stats.timed * fraction)
        cumulative = 0
        for index, count in enumerate(stats.bins):
            cumulative += count
            if cumulative >= target:
                lower = 0 if index == 0 else UPPERS[index - 1]
                upper = UPPERS[index] if index < len(UPPERS) else math.inf
                return lower, upper, index == 0
        raise AssertionError("Histogram lost a timed observation")

    def forecast(self, group):
        stats = self.groups[group]
        if self.fault:
            return "METRICS_ERROR", None
        if stats.observed != stats.timed:
            return "UNTIMED_RESULTS", None
        if not stats.timed:
            return "NOT_MEASURED", None
        if stats.timed < MIN_FORECAST:
            return "NOT_ESTIMATED", None
        if stats.total <= 0:
            return "CLOCK_RESOLUTION", None
        return "INDICATIVE", stats.total / stats.timed * 100000 / 3600


class NumericalBehavior(unittest.TestCase):
    def test_every_result_has_one_exclusive_group(self):
        groups = [group_for(code) for code in range(9)]
        groups += [group_for(1, True), group_for(2, True), group_for(999)]
        self.assertEqual(sorted(groups), list(range(1, GROUPS + 1)))

    def test_special_attachment_success_is_not_ordinary_new_or_reuse(self):
        m = Metrics()
        for code, duration in [(1, 50), (2, 5)]:
            m.record("AUTO", code, duration, True)
        self.assertEqual(m.groups[1].observed, 0)
        self.assertEqual(m.groups[2].observed, 0)
        self.assertEqual(m.groups[9].total, 50)
        self.assertEqual(m.groups[10].total, 5)
        self.assertEqual(m.forecast(1), ("NOT_MEASURED", None))

    def test_failed_and_review_costs_do_not_reduce_or_inflate_new_mean(self):
        m = Metrics()
        for _ in range(MIN_FORECAST):
            m.record("AUTO", 1, 0.5)
            m.record("AUTO", 2, 0.01)
            m.record("AUTO", 3, 50)
            m.record("AUTO", 7, 100)
        self.assertAlmostEqual(m.forecast(1)[1], 0.5 * 100000 / 3600)
        self.assertAlmostEqual(m.forecast(2)[1], 0.01 * 100000 / 3600)
        self.assertEqual(m.groups[3].total, 50 * MIN_FORECAST)
        self.assertEqual(m.groups[7].total, 100 * MIN_FORECAST)

    def test_private_benchmark_and_native_test_cannot_pollute_actual_metrics(self):
        m = Metrics()
        for mode in ("BENCHMARK", "SELFTEST", "REGISTER_LEGACY", "VERIFY", ""):
            for _ in range(50):
                m.record(mode, 1, 123)
        self.assertEqual(sum(s.observed for s in m.groups.values()), 0)
        self.assertEqual(m.forecast(1), ("NOT_MEASURED", None))

    def test_cancelled_and_aborted_are_elapsed_attempts_not_completed_successes(self):
        m = Metrics()
        m.record("AUTO", 5, 17)
        m.record("AUTO", 0, 29)
        self.assertEqual(m.groups[5].observed, 1)
        self.assertEqual(m.groups[11].observed, 1)
        self.assertEqual(m.groups[5].total, 17)
        self.assertEqual(m.groups[11].total, 29)
        self.assertEqual(m.groups[1].timed, 0)

    def test_negative_clock_keeps_outcome_but_suppresses_group_forecast(self):
        m = Metrics()
        for _ in range(MIN_FORECAST):
            m.record("AUTO", 1, 1)
            m.record("AUTO", 2, 0.1)
        m.record("AUTO", 1, -1)
        self.assertEqual(m.groups[1].observed, MIN_FORECAST + 1)
        self.assertEqual(m.groups[1].timed, MIN_FORECAST)
        self.assertEqual(m.forecast(1), ("UNTIMED_RESULTS", None))
        self.assertEqual(m.forecast(2)[0], "INDICATIVE")

    def test_zero_duration_is_explicit_and_not_a_zero_hour_prediction(self):
        m = Metrics()
        for _ in range(MIN_FORECAST):
            m.record("AUTO", 1, 0)
        self.assertEqual(m.groups[1].zeros, MIN_FORECAST)
        self.assertEqual(m.interval(1, 0.95), (0, FIRST_UPPER, True))
        self.assertEqual(m.forecast(1), ("CLOCK_RESOLUTION", None))

    def test_few_new_observations_do_not_suggest_a_certified_100k_runtime(self):
        m = Metrics()
        for _ in range(MIN_FORECAST - 1):
            m.record("AUTO", 1, 1)
        self.assertEqual(m.forecast(1), ("NOT_ESTIMATED", None))
        m.record("AUTO", 1, 1)
        self.assertEqual(m.forecast(1)[0], "INDICATIVE")

    def test_any_metrics_fault_suppresses_both_forecasts(self):
        m = Metrics()
        for code in (1, 2):
            for _ in range(MIN_FORECAST):
                m.record("AUTO", code, 1)
        m.fault = True
        self.assertEqual(m.forecast(1), ("METRICS_ERROR", None))
        self.assertEqual(m.forecast(2), ("METRICS_ERROR", None))

    def test_all_100000_observations_count_with_constant_memory(self):
        m = Metrics()
        # A much slower tail must remain visible, unlike a first-300 sample.
        for index in range(100000):
            m.record("AUTO", 1, 0.01 if index < 300 else 2)
        stats = m.groups[1]
        expected_total = 300 * 0.01 + 99700 * 2
        self.assertEqual(stats.timed, 100000)
        self.assertEqual(sum(stats.bins), 100000)
        self.assertEqual(stats.total, expected_total)
        self.assertEqual(len(stats.bins), BINS)
        self.assertEqual(sum(len(s.bins) for s in m.groups.values()), GROUPS * BINS)
        self.assertAlmostEqual(m.forecast(1)[1], expected_total / 3600)
        self.assertEqual(stats.minimum, 0.01)
        self.assertEqual(stats.maximum, 2)

    def test_nearest_rank_p50_p95_always_inside_reported_intervals(self):
        rng = random.Random(34)
        for count in (1, 2, 3, 19, 20, 21, 30, 100, 101, 999):
            observations = [0.0] + [10 ** rng.uniform(-5, 6) for _ in range(count - 1)]
            m = Metrics()
            for duration in observations:
                m.record("AUTO", 1, duration)
            ordered = sorted(observations)
            for fraction in (0.5, 0.95):
                exact = ordered[math.ceil(count * fraction) - 1]
                lower, upper, includes_lower = m.interval(1, fraction)
                if includes_lower:
                    self.assertGreaterEqual(exact, lower)
                else:
                    self.assertGreater(exact, lower)
                self.assertLessEqual(exact, upper)

    def test_bin_boundary_values_land_on_the_inclusive_upper_side(self):
        for index, upper in enumerate(UPPERS):
            m = Metrics()
            m.record("AUTO", 1, upper)
            self.assertEqual(m.groups[1].bins[index], 1)
            m.record("AUTO", 1, math.nextafter(upper, math.inf))
            self.assertEqual(m.groups[1].bins[index + 1], 1)

    def test_large_wait_overflow_bin_retains_observation_and_real_mean(self):
        m = Metrics()
        for _ in range(MIN_FORECAST):
            m.record("AUTO", 1, 3 * 86400)
        self.assertEqual(m.groups[1].bins[-1], MIN_FORECAST)
        self.assertTrue(math.isinf(m.interval(1, 0.95)[1]))
        self.assertEqual(m.forecast(1)[1], 3 * 86400 * 100000 / 3600)

    def test_reset_discards_previous_run_and_keeps_all_groups_available(self):
        m = Metrics()
        m.record("AUTO", 1, 42)
        m = Metrics()
        self.assertEqual(sum(s.observed for s in m.groups.values()), 0)
        self.assertEqual(len(m.groups), GROUPS)

    def test_non_success_outcomes_are_conserved_individually(self):
        m = Metrics()
        for code in (3, 4, 5, 6, 7, 8, 0, 17):
            m.record("AUTO", code, code + 100)
        self.assertEqual(sum(s.observed for s in m.groups.values()), 8)
        for code in (3, 4, 5, 6, 7, 8, 0, 17):
            self.assertEqual(m.groups[group_for(code)].total, code + 100)


class SourceContracts(unittest.TestCase):
    def test_metrics_has_no_extra_export_mail_reads_or_file_writes(self):
        without_strings = re.sub(r'"(?:[^"\r\n]|"")*"', '""', SOURCE)
        code = "\n".join(line.split("'", 1)[0] for line in without_strings.splitlines())
        self.assertNotRegex(code, r"(?i)\b(?:SaveAs|Save|Send|OpenSharedItem|GetItemFromID|CreateObject|OpenTextFile|CreateTextFile|Kill|Randomize|Rnd)\b")

    def test_auto_filter_precedes_all_counter_updates(self):
        record = SOURCE.split("Private Sub MBF_Record", 1)[1].split("End Sub", 1)[0]
        self.assertLess(record.index('If mode <> "AUTO" Then Exit Sub'), record.index("gMetricObserved(group) ="))

    def test_no_first_n_reservoir_or_external_sampling_state(self):
        self.assertNotRegex(SOURCE, r"(?i)\bReDim\b")
        self.assertIn("gMetricHistogram(group, bin) = gMetricHistogram(group, bin) + 1#", SOURCE)
        self.assertNotRegex(DECLARATIONS, r"\([ ]*\)[ ]+As")

    def test_native_data_is_not_represented_as_100k_workload_pass(self):
        self.assertIn("These observations do not replace a real 100,000-item test", SOURCE)
        self.assertIn("No whole-run duration is predicted.", SOURCE)
        self.assertIn("not a statistical confidence guarantee", SOURCE)

    def test_new_and_reuse_forecasts_use_separate_success_groups(self):
        self.assertIn('MBF_Forecast(1, "NEW")', SOURCE)
        self.assertIn('MBF_Forecast(2, "REUSE")', SOURCE)
        self.assertNotRegex(SOURCE, r"MBF_Forecast\((?:3|4|5|6|7|8|9|10|11|12),")

    def test_fault_path_cannot_stop_or_modify_backup(self):
        self.assertNotIn("Err.Raise", SOURCE)
        self.assertNotRegex(SOURCE, r"g(?:Fatal|Cancel|MailNew|MailReused|ContentReview)\s*=")
        self.assertIn("UNAVAILABLE_METRICS_ERROR", SOURCE)

    def test_negative_time_is_not_clamped_to_fake_zero(self):
        self.assertIn("If elapsedSeconds < 0# Then Exit Sub", SOURCE)
        self.assertIn("UNAVAILABLE_UNTIMED_RESULTS", SOURCE)
        self.assertNotRegex(SOURCE, r"(?m)^\s*elapsedSeconds\s*=\s*0")

    def test_module_adds_no_public_macro(self):
        self.assertNotRegex(SOURCE, r"(?m)^Public\s+(?:Sub|Function)\b")


class IntegrationContracts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.main = (ROOT / "mb_main.basfrag").read_text(encoding="ascii")

    def test_each_normal_run_resets_metrics(self):
        reset = self.main.split("Private Sub MB_Reset()", 1)[1].split("End Sub", 1)[0]
        self.assertIn("MBF_Reset", reset)

    def test_cancelled_item_is_recorded_before_early_exit(self):
        queue = self.main.split("Private Sub MB_ProcessQueue", 1)[1].split("End Sub", 1)[0]
        self.assertLess(queue.index("MBF_Record"), queue.index("If result = 5 Then Exit Do"))

    def test_aborted_in_flight_item_is_observed_without_double_counting(self):
        queue = self.main.split("Private Sub MB_ProcessQueue", 1)[1].split("End Sub", 1)[0]
        ordinary, failed = queue.split("Failed:", 1)
        self.assertIn("metricInFlight = True", ordinary)
        self.assertIn("metricInFlight = False", ordinary)
        self.assertIn("result = 0", ordinary)
        self.assertIn("If metricInFlight Then", failed)
        self.assertIn("MBF_Record", failed)
        self.assertLess(failed.index("errNo = Err.Number"), failed.index("MBF_Record"))

    def test_auto_summary_and_assembly_include_actual_metrics(self):
        self.assertIn("MBF_Summary(gMode)", self.main)
        assembly = (ROOT / "mb_assemble.py").read_text(encoding="ascii")
        self.assertIn("mb_metrics_declarations.basfrag", assembly)
        self.assertIn("mb_metrics.basfrag", assembly)


if __name__ == "__main__":
    integrated = "--integration" in sys.argv
    if integrated:
        sys.argv.remove("--integration")
    suite = unittest.TestSuite()
    loader = unittest.TestLoader()
    for test in (NumericalBehavior, SourceContracts):
        suite.addTests(loader.loadTestsFromTestCase(test))
    if integrated:
        suite.addTests(loader.loadTestsFromTestCase(IntegrationContracts))
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    report = {
        "kind": "numerical/behavior models and source contracts; NOT Outlook/VBA execution",
        "integration_hooks_checked": integrated,
        "tests": result.testsRun,
        "failures": len(result.failures),
        "errors": len(result.errors),
        "pass": result.wasSuccessful(),
        "native_outlook_runtime": "NOT_RUN",
        "real_100000_item_performance": "NOT_RUN",
    }
    report_path = ROOT / "vba_parser_review" / "v34_metrics_models.json"
    report_path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, indent=2))
    raise SystemExit(not result.wasSuccessful())
