from __future__ import annotations

import csv
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

from travel_angle_statistics.analysis import analyze_records
from travel_angle_statistics.constants import CSV_FIELDS, EXPECTED_STEPS_BY_DIRECTION
from travel_angle_statistics.errors import DataValidationError, DuplicateStepError
from travel_angle_statistics.loader import load_records_from_csv
from travel_angle_statistics.models import MeasurementRecord


def _default_angle(direction: str, step: str) -> Decimal:
    if direction == "Up" and step == "50-40":
        return Decimal("25")
    if direction == "Down" and step == "40-50":
        return Decimal("25")
    return Decimal("15")


def build_series_records(
    *,
    test_id: str = "T-001",
    motor_id: str = "M-1",
    motor_name: str = "MotorX",
    motor_variant: str = "6Nm",
    software_version: str = "A01",
    load_nm: Decimal = Decimal("0.1"),
    measurement_no: str = "1",
    missing: set[tuple[str, str]] | None = None,
    no_movement: set[tuple[str, str]] | None = None,
    angle_overrides: dict[tuple[str, str], Decimal] | None = None,
) -> list[MeasurementRecord]:
    missing = missing or set()
    no_movement = no_movement or set()
    angle_overrides = angle_overrides or {}

    records: list[MeasurementRecord] = []
    row = 2
    for direction in ("Up", "Down"):
        for step in EXPECTED_STEPS_BY_DIRECTION[direction]:
            key = (direction, step)
            if key in missing:
                continue
            status = "no_movement" if key in no_movement else "measured"
            if status == "measured":
                angle = angle_overrides.get(key, _default_angle(direction, step))
            else:
                angle = None
            records.append(
                MeasurementRecord(
                    source_file="synthetic.csv",
                    row_number=row,
                    test_id=test_id,
                    motor_id=motor_id,
                    motor_name=motor_name,
                    motor_variant=motor_variant,
                    software_version=software_version,
                    load_nm=load_nm,
                    measurement_no=measurement_no,
                    direction=direction,
                    step=step,
                    angle_deg=angle,
                    movement_status=status,
                )
            )
            row += 1
    return records


class AnalyzeRecordsTests(unittest.TestCase):
    def test_excludes_last_measured_before_trailing_no_movement(self) -> None:
        records = build_series_records(
            no_movement={("Up", "20-10"), ("Up", "10-0")},
            angle_overrides={("Up", "30-20"): Decimal("19")},
        )
        analysis = analyze_records(records)
        up_100_row = next(
            row
            for row in analysis.movement_angle_rows
            if row["movement_group"] == "Up 100ms"
        )
        self.assertEqual(up_100_row["measurement_count"], 4)

    def test_middle_gap_does_not_exclude_neighbors(self) -> None:
        records = build_series_records(
            no_movement={("Up", "30-20")},
            angle_overrides={
                ("Up", "20-10"): Decimal("16"),
                ("Up", "10-0"): Decimal("24"),
            },
        )
        analysis = analyze_records(records)
        up_100_row = next(
            row
            for row in analysis.movement_angle_rows
            if row["movement_group"] == "Up 100ms"
        )
        self.assertEqual(up_100_row["measurement_count"], 5)

    def test_missing_last_step_excludes_previous_for_movement_but_not_final_position(self) -> None:
        records = build_series_records(
            missing={("Up", "10-0")},
            angle_overrides={("Up", "20-10"): Decimal("30")},
        )
        analysis = analyze_records(records)
        up_100_row = next(
            row
            for row in analysis.movement_angle_rows
            if row["movement_group"] == "Up 100ms"
        )
        self.assertEqual(up_100_row["measurement_count"], 5)

        final_series = analysis.final_position_series_rows[0]
        self.assertEqual(final_series["partial"], "PARTIAL")
        self.assertEqual(final_series["sum_up_deg"], Decimal("145"))

    def test_missing_step_affects_percentage_denominator(self) -> None:
        records = build_series_records(missing={("Up", "20-10")})
        analysis = analyze_records(records)
        series_row = analysis.tilt_series_rows[0]
        self.assertEqual(series_row["is_complete"], "No")
        self.assertEqual(series_row["performed_up"], 8)
        self.assertEqual(series_row["performed_up_pct"], Decimal("88.88888888888888888888888889"))

    def test_combines_different_motor_id_when_configuration_matches(self) -> None:
        first = build_series_records(motor_id="M-1", measurement_no="1")
        second = build_series_records(motor_id="M-2", measurement_no="2")
        analysis = analyze_records(first + second)
        up_first_row = next(
            row
            for row in analysis.movement_angle_rows
            if row["movement_group"] == "Up 100ms - first movement"
        )
        self.assertEqual(up_first_row["measurement_count"], 2)
        self.assertEqual(up_first_row["contributing_series_count"], 2)

    def test_final_position_summary_not_evaluable_without_deviation(self) -> None:
        records = build_series_records(
            no_movement={
                ("Up", step) for step in EXPECTED_STEPS_BY_DIRECTION["Up"]
            },
            angle_overrides={("Down", "40-50"): Decimal("31")},
        )
        analysis = analyze_records(records)
        summary = analysis.final_position_summary_rows[0]
        self.assertEqual(summary["evaluated_series_count"], 0)
        self.assertEqual(summary["exceedance_pct"], "NOT EVALUABLE")


class LoadRecordsTests(unittest.TestCase):
    def test_rejects_status_angle_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            csv_path = Path(directory) / "invalid.csv"
            with csv_path.open("w", encoding="utf-8", newline="") as file_handle:
                writer = csv.DictWriter(file_handle, fieldnames=CSV_FIELDS, quoting=csv.QUOTE_ALL)
                writer.writeheader()
                writer.writerow(
                    {
                        "test_id": "T",
                        "motor_id": "1",
                        "motor_name": "M",
                        "motor_variant": "V",
                        "software_version": "S",
                        "load_nm": "0.1",
                        "measurement_no": "1",
                        "direction": "Up",
                        "step": "90-80",
                        "angle_deg": "10",
                        "movement_status": "no_movement",
                    }
                )
            with self.assertRaises(DataValidationError):
                load_records_from_csv([csv_path])

    def test_rejects_duplicate_step_across_files(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path_a = Path(directory) / "a.csv"
            path_b = Path(directory) / "b.csv"
            for index, path in enumerate((path_a, path_b), start=1):
                with path.open("w", encoding="utf-8", newline="") as file_handle:
                    writer = csv.DictWriter(
                        file_handle, fieldnames=CSV_FIELDS, quoting=csv.QUOTE_ALL
                    )
                    writer.writeheader()
                    writer.writerow(
                        {
                            "test_id": "T",
                            "motor_id": "1",
                            "motor_name": "M",
                            "motor_variant": "V",
                            "software_version": "S",
                            "load_nm": "0.1",
                            "measurement_no": "1",
                            "direction": "Up",
                            "step": "90-80",
                            "angle_deg": str(10 + index),
                            "movement_status": "measured",
                        }
                    )
            with self.assertRaises(DuplicateStepError):
                load_records_from_csv([path_a, path_b])


if __name__ == "__main__":
    unittest.main()
