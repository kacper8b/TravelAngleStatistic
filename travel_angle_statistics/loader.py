from __future__ import annotations

import csv
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path

from .constants import CSV_FIELDS, EXPECTED_STEPS_BY_DIRECTION
from .errors import DataValidationError, DuplicateStepError
from .models import MeasurementRecord


def _parse_decimal(value: str, source_file: str, row_number: int, field_name: str) -> Decimal:
    try:
        parsed = Decimal(value.replace(",", "."))
    except InvalidOperation as error:
        raise DataValidationError(
            f"{source_file}:{row_number}: field '{field_name}' has invalid decimal value {value!r}"
        ) from error
    if not parsed.is_finite():
        raise DataValidationError(
            f"{source_file}:{row_number}: field '{field_name}' has non-finite value {value!r}"
        )
    return parsed


def _require_text_field(
    row: dict[str, str],
    field_name: str,
    source_file: str,
    row_number: int,
) -> str:
    value = row.get(field_name)
    if value is None:
        raise DataValidationError(
            f"{source_file}:{row_number}: missing required field '{field_name}'"
        )
    text = value.strip()
    if not text:
        raise DataValidationError(
            f"{source_file}:{row_number}: field '{field_name}' is required"
        )
    return text


def _validate_and_build_record(
    row: dict[str, str],
    source_file: str,
    row_number: int,
) -> MeasurementRecord:
    test_id = _require_text_field(row, "test_id", source_file, row_number)
    motor_id = _require_text_field(row, "motor_id", source_file, row_number)
    motor_name = _require_text_field(row, "motor_name", source_file, row_number)
    motor_variant = _require_text_field(row, "motor_variant", source_file, row_number)
    software_version = _require_text_field(row, "software_version", source_file, row_number)
    load_nm_text = _require_text_field(row, "load_nm", source_file, row_number)
    measurement_no = _require_text_field(row, "measurement_no", source_file, row_number)
    direction = _require_text_field(row, "direction", source_file, row_number)
    step = _require_text_field(row, "step", source_file, row_number)
    movement_status = _require_text_field(row, "movement_status", source_file, row_number)

    if direction not in EXPECTED_STEPS_BY_DIRECTION:
        raise DataValidationError(
            f"{source_file}:{row_number}: field 'direction' has unknown value {direction!r}"
        )
    if step not in EXPECTED_STEPS_BY_DIRECTION[direction]:
        raise DataValidationError(
            f"{source_file}:{row_number}: field 'step' has unknown value {step!r} for direction {direction!r}"
        )

    if movement_status not in {"measured", "no_movement"}:
        raise DataValidationError(
            f"{source_file}:{row_number}: field 'movement_status' has unknown value {movement_status!r}"
        )

    load_nm = _parse_decimal(load_nm_text, source_file, row_number, "load_nm")
    angle_text = (row.get("angle_deg") or "").strip()
    angle_deg: Decimal | None

    if movement_status == "measured":
        if not angle_text:
            raise DataValidationError(
                f"{source_file}:{row_number}: field 'angle_deg' must be present when movement_status='measured'"
            )
        angle_deg = _parse_decimal(angle_text, source_file, row_number, "angle_deg")
        if angle_deg == 0:
            raise DataValidationError(
                f"{source_file}:{row_number}: field 'angle_deg' cannot be zero when movement_status='measured'"
            )
    else:
        if angle_text:
            raise DataValidationError(
                f"{source_file}:{row_number}: field 'angle_deg' must be empty when movement_status='no_movement'"
            )
        angle_deg = None

    return MeasurementRecord(
        source_file=source_file,
        row_number=row_number,
        test_id=test_id,
        motor_id=motor_id,
        motor_name=motor_name,
        motor_variant=motor_variant,
        software_version=software_version,
        load_nm=load_nm,
        measurement_no=measurement_no,
        direction=direction,
        step=step,
        angle_deg=angle_deg,
        movement_status=movement_status,
    )


def _detect_duplicate_steps(records: list[MeasurementRecord]) -> None:
    grouped: dict[
        tuple[str, str, str, str, str, Decimal, str, str, str],
        list[MeasurementRecord],
    ] = defaultdict(list)

    for record in records:
        key = (
            record.test_id,
            record.motor_id,
            record.motor_name,
            record.motor_variant,
            record.software_version,
            record.load_nm,
            record.measurement_no,
            record.direction,
            record.step,
        )
        grouped[key].append(record)

    conflicts: list[str] = []
    for key, conflict_records in grouped.items():
        if len(conflict_records) < 2:
            continue
        (
            test_id,
            motor_id,
            motor_name,
            motor_variant,
            software_version,
            load_nm,
            measurement_no,
            direction,
            step,
        ) = key
        header = (
            "duplicate step for "
            f"test_id={test_id}, motor_id={motor_id}, motor_name={motor_name}, "
            f"motor_variant={motor_variant}, software_version={software_version}, "
            f"load_nm={load_nm}, measurement_no={measurement_no}, direction={direction}, step={step}"
        )
        lines = [header]
        for conflict in conflict_records:
            angle_value = "" if conflict.angle_deg is None else str(conflict.angle_deg)
            lines.append(
                f"  - {conflict.source_file}:{conflict.row_number} "
                f"(angle_deg={angle_value!r}, movement_status={conflict.movement_status!r})"
            )
        conflicts.append("\n".join(lines))

    if conflicts:
        message = "Duplicate steps detected:\n" + "\n".join(conflicts)
        raise DuplicateStepError(message)


def load_records_from_csv(csv_paths: list[Path]) -> list[MeasurementRecord]:
    records: list[MeasurementRecord] = []
    for csv_path in csv_paths:
        if not csv_path.exists():
            raise DataValidationError(f"CSV source does not exist: {csv_path}")

        with csv_path.open("r", encoding="utf-8-sig", newline="") as csv_file:
            reader = csv.DictReader(csv_file)
            if reader.fieldnames != CSV_FIELDS:
                raise DataValidationError(
                    f"{csv_path.name}: invalid header. Expected {CSV_FIELDS}, got {reader.fieldnames}"
                )
            for row_number, row in enumerate(reader, start=2):
                if row is None or any(key is None for key in row.keys()):
                    raise DataValidationError(f"{csv_path.name}:{row_number}: malformed row")
                records.append(
                    _validate_and_build_record(
                        row=row,
                        source_file=csv_path.name,
                        row_number=row_number,
                    )
                )

    _detect_duplicate_steps(records)
    return records
