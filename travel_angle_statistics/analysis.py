from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Any

from .constants import (
    ANGLE_REQUIREMENT_LOAD,
    EXPECTED_STEPS_BY_DIRECTION,
    FINAL_POSITION_REQUIREMENT_LOADS,
    MOVEMENT_GROUP_KIND,
    MOVEMENT_GROUPS,
    STATUS_FAIL,
    STATUS_NA,
    STATUS_NOT_EVALUABLE,
    STATUS_PARTIAL,
    STATUS_PASS,
)
from .models import BaseConfigKey, MeasurementRecord, SeriesKey


def _decimal_to_text(value: Decimal) -> str:
    return format(value.normalize(), "f")


def _mean(values: list[Decimal]) -> Decimal | None:
    if not values:
        return None
    return sum(values) / Decimal(len(values))


def _median(values: list[Decimal]) -> Decimal | None:
    if not values:
        return None
    sorted_values = sorted(values)
    count = len(sorted_values)
    mid = count // 2
    if count % 2 == 1:
        return sorted_values[mid]
    return (sorted_values[mid - 1] + sorted_values[mid]) / Decimal(2)


def _sample_stddev(values: list[Decimal]) -> Decimal | None:
    if len(values) < 2:
        return None
    mean_value = _mean(values)
    if mean_value is None:
        return None
    variance = sum((value - mean_value) ** 2 for value in values) / Decimal(len(values) - 1)
    return variance.sqrt()


def _ratio_percent(numerator: int, denominator: int) -> Decimal | str:
    if denominator == 0:
        return STATUS_NOT_EVALUABLE
    return (Decimal(numerator) * Decimal(100)) / Decimal(denominator)


@dataclass(frozen=True)
class AnalysisResult:
    movement_angle_rows: list[dict[str, Any]]
    tilt_overall_row: dict[str, Any]
    tilt_series_rows: list[dict[str, Any]]
    tilt_no_movement_rows: list[dict[str, Any]]
    final_position_series_rows: list[dict[str, Any]]
    final_position_summary_rows: list[dict[str, Any]]


def empty_analysis_result() -> AnalysisResult:
    return AnalysisResult(
        movement_angle_rows=[],
        tilt_overall_row={
            "performed_up": 0,
            "performed_down": 0,
            "performed_total": 0,
            "expected_up": 0,
            "expected_down": 0,
            "expected_total": 0,
            "performed_up_pct": STATUS_NOT_EVALUABLE,
            "performed_down_pct": STATUS_NOT_EVALUABLE,
            "performed_total_pct": STATUS_NOT_EVALUABLE,
        },
        tilt_series_rows=[],
        tilt_no_movement_rows=[],
        final_position_series_rows=[],
        final_position_summary_rows=[],
    )


def _series_sort_key(series_key: SeriesKey) -> tuple[str, str, str, str, str, str, str]:
    return (
        series_key.test_id,
        series_key.motor_id,
        series_key.motor_name,
        series_key.motor_variant,
        series_key.software_version,
        _decimal_to_text(series_key.load_nm),
        series_key.measurement_no,
    )


def _base_config_sort_key(config: BaseConfigKey) -> tuple[str, str, str, str]:
    return (
        config.motor_name,
        config.motor_variant,
        config.software_version,
        _decimal_to_text(config.load_nm),
    )


def _excluded_steps_for_direction(
    by_direction_step: dict[tuple[str, str], MeasurementRecord],
    direction: str,
) -> set[str]:
    expected_steps = EXPECTED_STEPS_BY_DIRECTION[direction]
    measured_indices: list[int] = []
    for index, step in enumerate(expected_steps):
        record = by_direction_step.get((direction, step))
        if record is not None and record.movement_status == "measured":
            measured_indices.append(index)
    if not measured_indices:
        return set()

    last_measured_index = measured_indices[-1]
    if last_measured_index >= len(expected_steps) - 1:
        return set()

    trailing_steps = expected_steps[last_measured_index + 1 :]
    if all(
        (
            by_direction_step.get((direction, step)) is None
            or by_direction_step[(direction, step)].movement_status == "no_movement"
        )
        for step in trailing_steps
    ):
        return {expected_steps[last_measured_index]}
    return set()


def analyze_records(records: list[MeasurementRecord]) -> AnalysisResult:
    if not records:
        return empty_analysis_result()

    by_series_records: dict[SeriesKey, list[MeasurementRecord]] = {}
    by_series_direction_step: dict[SeriesKey, dict[tuple[str, str], MeasurementRecord]] = {}
    for record in records:
        by_series_records.setdefault(record.series_key, []).append(record)
        by_series_direction_step.setdefault(record.series_key, {})[
            (record.direction, record.step)
        ] = record

    series_complete: dict[SeriesKey, bool] = {}
    excluded_steps: dict[tuple[SeriesKey, str], set[str]] = {}
    tilt_series_rows: list[dict[str, Any]] = []
    tilt_no_movement_rows: list[dict[str, Any]] = []
    final_position_series_rows: list[dict[str, Any]] = []

    for series_key in sorted(by_series_records, key=_series_sort_key):
        direction_step_map = by_series_direction_step[series_key]
        missing_any = False
        for direction in ("Up", "Down"):
            expected = EXPECTED_STEPS_BY_DIRECTION[direction]
            for step in expected:
                if (direction, step) not in direction_step_map:
                    missing_any = True
                    break
            if missing_any:
                break

        series_complete[series_key] = not missing_any
        excluded_steps[(series_key, "Up")] = _excluded_steps_for_direction(direction_step_map, "Up")
        excluded_steps[(series_key, "Down")] = _excluded_steps_for_direction(direction_step_map, "Down")

        series_records = by_series_records[series_key]
        performed_up = sum(
            1
            for record in series_records
            if record.direction == "Up" and record.movement_status == "measured"
        )
        performed_down = sum(
            1
            for record in series_records
            if record.direction == "Down" and record.movement_status == "measured"
        )
        performed_total = performed_up + performed_down

        tilt_series_rows.append(
            {
                "test_id": series_key.test_id,
                "motor_id": series_key.motor_id,
                "motor_name": series_key.motor_name,
                "motor_variant": series_key.motor_variant,
                "software_version": series_key.software_version,
                "load_nm": series_key.load_nm,
                "measurement_no": series_key.measurement_no,
                "is_complete": "Yes" if series_complete[series_key] else "No",
                "performed_up": performed_up,
                "performed_down": performed_down,
                "performed_total": performed_total,
                "up_requirement": STATUS_PASS if performed_up == 9 else STATUS_FAIL,
                "down_requirement": STATUS_PASS if performed_down == 9 else STATUS_FAIL,
                "performed_up_pct": _ratio_percent(performed_up, 9),
                "performed_down_pct": _ratio_percent(performed_down, 9),
                "performed_total_pct": _ratio_percent(performed_total, 18),
            }
        )

        for record in sorted(
            (item for item in series_records if item.movement_status == "no_movement"),
            key=lambda item: (item.direction, item.step),
        ):
            tilt_no_movement_rows.append(
                {
                    "test_id": record.test_id,
                    "motor_id": record.motor_id,
                    "motor_name": record.motor_name,
                    "motor_variant": record.motor_variant,
                    "software_version": record.software_version,
                    "load_nm": record.load_nm,
                    "measurement_no": record.measurement_no,
                    "direction": record.direction,
                    "step": record.step,
                }
            )

        measured_up = [
            record.angle_deg
            for record in series_records
            if record.direction == "Up"
            and record.movement_status == "measured"
            and record.angle_deg is not None
        ]
        measured_down = [
            record.angle_deg
            for record in series_records
            if record.direction == "Down"
            and record.movement_status == "measured"
            and record.angle_deg is not None
        ]

        sum_up = sum(measured_up) if measured_up else None
        sum_down = sum(measured_down) if measured_down else None
        deviation = (sum_up - sum_down) if (sum_up is not None and sum_down is not None) else None
        load_applicable = series_key.load_nm in FINAL_POSITION_REQUIREMENT_LOADS

        if deviation is None:
            deviation_requirement = STATUS_NOT_EVALUABLE if load_applicable else STATUS_NA
        elif not load_applicable:
            deviation_requirement = STATUS_NA
        else:
            deviation_requirement = (
                STATUS_PASS if abs(deviation) <= Decimal("10") else STATUS_FAIL
            )

        partial_flag = (
            STATUS_PARTIAL
            if (not series_complete[series_key] and deviation is not None)
            else ""
        )
        final_position_series_rows.append(
            {
                "test_id": series_key.test_id,
                "motor_id": series_key.motor_id,
                "motor_name": series_key.motor_name,
                "motor_variant": series_key.motor_variant,
                "software_version": series_key.software_version,
                "load_nm": series_key.load_nm,
                "measurement_no": series_key.measurement_no,
                "is_complete": "Yes" if series_complete[series_key] else "No",
                "sum_up_deg": sum_up if sum_up is not None else STATUS_NOT_EVALUABLE,
                "sum_down_deg": sum_down if sum_down is not None else STATUS_NOT_EVALUABLE,
                "deviation_deg": deviation if deviation is not None else STATUS_NOT_EVALUABLE,
                "deviation_requirement": deviation_requirement,
                "partial": partial_flag,
            }
        )

    performed_up_total = sum(
        1 for record in records if record.direction == "Up" and record.movement_status == "measured"
    )
    performed_down_total = sum(
        1
        for record in records
        if record.direction == "Down" and record.movement_status == "measured"
    )
    series_count = len(by_series_records)
    expected_up_total = series_count * 9
    expected_down_total = series_count * 9
    expected_total = series_count * 18
    tilt_overall_row = {
        "performed_up": performed_up_total,
        "performed_down": performed_down_total,
        "performed_total": performed_up_total + performed_down_total,
        "expected_up": expected_up_total,
        "expected_down": expected_down_total,
        "expected_total": expected_total,
        "performed_up_pct": _ratio_percent(performed_up_total, expected_up_total),
        "performed_down_pct": _ratio_percent(performed_down_total, expected_down_total),
        "performed_total_pct": _ratio_percent(
            performed_up_total + performed_down_total, expected_total
        ),
    }

    base_configs = sorted(
        {record.base_config_key for record in records},
        key=_base_config_sort_key,
    )
    included_records: list[MeasurementRecord] = []
    for record in records:
        if record.movement_status != "measured":
            continue
        excluded_for_direction = excluded_steps[(record.series_key, record.direction)]
        if record.step in excluded_for_direction:
            continue
        included_records.append(record)

    grouped_records: dict[tuple[BaseConfigKey, str], list[MeasurementRecord]] = {}
    contributing_series: dict[tuple[BaseConfigKey, str], set[SeriesKey]] = {}
    incomplete_contributing_series: dict[tuple[BaseConfigKey, str], set[SeriesKey]] = {}

    for record in included_records:
        key = (record.base_config_key, record.movement_group)
        grouped_records.setdefault(key, []).append(record)
        contributing_series.setdefault(key, set()).add(record.series_key)
        if not series_complete[record.series_key]:
            incomplete_contributing_series.setdefault(key, set()).add(record.series_key)

    movement_angle_rows: list[dict[str, Any]] = []
    for base_config in base_configs:
        for movement_group in MOVEMENT_GROUPS:
            key = (base_config, movement_group)
            rows = grouped_records.get(key, [])
            angles = [
                record.angle_deg
                for record in rows
                if record.movement_status == "measured" and record.angle_deg is not None
            ]
            mean_value = _mean(angles)
            max_value = max(angles) if angles else None

            group_kind = MOVEMENT_GROUP_KIND[movement_group]
            if base_config.load_nm == ANGLE_REQUIREMENT_LOAD:
                if group_kind == "100ms":
                    if mean_value is None:
                        mean_requirement = STATUS_NOT_EVALUABLE
                    else:
                        mean_requirement = (
                            STATUS_PASS
                            if Decimal("14") <= mean_value <= Decimal("18")
                            else STATUS_FAIL
                        )
                    maximum_limit = Decimal("20")
                else:
                    mean_requirement = STATUS_NA
                    maximum_limit = Decimal("30")

                if max_value is None:
                    maximum_requirement = STATUS_NOT_EVALUABLE
                    exceedance_count: int | str = STATUS_NOT_EVALUABLE
                else:
                    maximum_requirement = (
                        STATUS_PASS if max_value <= maximum_limit else STATUS_FAIL
                    )
                    exceedance_count = sum(1 for angle in angles if angle > maximum_limit)
            else:
                mean_requirement = STATUS_NA
                maximum_requirement = STATUS_NA
                exceedance_count = STATUS_NA

            movement_angle_rows.append(
                {
                    "motor_name": base_config.motor_name,
                    "motor_variant": base_config.motor_variant,
                    "software_version": base_config.software_version,
                    "load_nm": base_config.load_nm,
                    "movement_group": movement_group,
                    "measurement_count": len(angles),
                    "contributing_series_count": len(contributing_series.get(key, set())),
                    "incomplete_contributing_series_count": len(
                        incomplete_contributing_series.get(key, set())
                    ),
                    "mean_angle_deg": mean_value if mean_value is not None else STATUS_NOT_EVALUABLE,
                    "median_angle_deg": _median(angles)
                    if angles
                    else STATUS_NOT_EVALUABLE,
                    "min_angle_deg": min(angles) if angles else STATUS_NOT_EVALUABLE,
                    "max_angle_deg": max_value if max_value is not None else STATUS_NOT_EVALUABLE,
                    "sample_stddev_deg": _sample_stddev(angles)
                    if len(angles) >= 2
                    else STATUS_NOT_EVALUABLE,
                    "mean_requirement": mean_requirement,
                    "max_requirement": maximum_requirement,
                    "exceedance_count": exceedance_count,
                }
            )

    final_position_summary_rows: list[dict[str, Any]] = []
    summary_keys = sorted(
        {
            (
                row["motor_name"],
                row["motor_variant"],
                row["software_version"],
                row["load_nm"],
            )
            for row in final_position_series_rows
        },
        key=lambda item: (item[0], item[1], item[2], _decimal_to_text(item[3])),
    )
    for motor_name, motor_variant, software_version, load_nm in summary_keys:
        rows = [
            row
            for row in final_position_series_rows
            if row["motor_name"] == motor_name
            and row["motor_variant"] == motor_variant
            and row["software_version"] == software_version
            and row["load_nm"] == load_nm
        ]
        load_applicable = load_nm in FINAL_POSITION_REQUIREMENT_LOADS
        evaluated_rows = [
            row
            for row in rows
            if row["deviation_requirement"] in {STATUS_PASS, STATUS_FAIL}
        ]
        evaluated_count = len(evaluated_rows)
        incomplete_evaluated_count = sum(1 for row in evaluated_rows if row["partial"] == STATUS_PARTIAL)
        exceedance_count = sum(1 for row in evaluated_rows if row["deviation_requirement"] == STATUS_FAIL)

        if not load_applicable:
            exceedance_percentage: Decimal | str = STATUS_NA
        elif evaluated_count == 0:
            exceedance_percentage = STATUS_NOT_EVALUABLE
        else:
            exceedance_percentage = (
                Decimal(exceedance_count) * Decimal(100) / Decimal(evaluated_count)
            )

        final_position_summary_rows.append(
            {
                "motor_name": motor_name,
                "motor_variant": motor_variant,
                "software_version": software_version,
                "load_nm": load_nm,
                "evaluated_series_count": evaluated_count,
                "incomplete_evaluated_series_count": incomplete_evaluated_count,
                "exceedance_count": exceedance_count,
                "exceedance_pct": exceedance_percentage,
            }
        )

    return AnalysisResult(
        movement_angle_rows=movement_angle_rows,
        tilt_overall_row=tilt_overall_row,
        tilt_series_rows=tilt_series_rows,
        tilt_no_movement_rows=tilt_no_movement_rows,
        final_position_series_rows=final_position_series_rows,
        final_position_summary_rows=final_position_summary_rows,
    )
