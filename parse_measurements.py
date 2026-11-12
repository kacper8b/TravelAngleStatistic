from __future__ import annotations

import csv
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


DATA_DIR = Path(__file__).resolve().parent / "data"
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "measurements"
MASTER_CSV = DATA_DIR / "measurements.csv"
CSV_FIELDS = [
    "test_id",
    "motor_id",
    "motor_name",
    "motor_variant",
    "software_version",
    "load_nm",
    "measurement_no",
    "direction",
    "step",
    "angle_deg",
    "movement_status",
]
METADATA_FIELDS = CSV_FIELDS[:5]
STEP_PATTERN = re.compile(r"^\s*(\d+(?:[.,]\d+)?)\s*-\s*(\d+(?:[.,]\d+)?)\s*$")
LOAD_PATTERN = re.compile(r"^\s*(\d+(?:[.,]\d+)?)\s*Nm\s*$", re.IGNORECASE)
NO_MOVEMENT_VALUES = {"-", "–", "—"}


@dataclass
class SheetData:
    workbook_name: str
    sheet_name: str
    records: list[dict[str, str]]


@dataclass
class OutputGroup:
    metadata: dict[str, str]
    sheets: list[SheetData] = field(default_factory=list)


def normalize_text(value: object) -> str:
    return str(value).strip().casefold() if value is not None else ""


def format_scalar(value: object) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def format_decimal(value: Decimal) -> str:
    return format(value.normalize(), "f")


def find_measurement_number_row(rows: list[tuple[object, ...]]) -> int:
    for row_index, row in enumerate(rows):
        if any(normalize_text(value).rstrip(".") == "measurement no" for value in row):
            return row_index + 1
    raise ValueError("could not find the 'Measurement no.' header")


def find_direction(
    rows: list[tuple[object, ...]], header_row: int, step_column: int
) -> str:
    for row_index in range(header_row - 1, max(-1, header_row - 5), -1):
        if row_index >= len(rows) or step_column >= len(rows[row_index]):
            continue
        value = normalize_text(rows[row_index][step_column])
        if value in {"up direction", "down direction"}:
            return "Up" if value.startswith("up") else "Down"
    raise ValueError(
        f"could not find a direction above row {header_row + 1}, "
        f"column {get_column_letter(step_column + 1)}"
    )


def find_load(rows: list[tuple[object, ...]], before_row: int, step_column: int) -> str:
    for row_index in range(before_row - 1, -1, -1):
        if step_column >= len(rows[row_index]):
            continue
        value = rows[row_index][step_column]
        match = LOAD_PATTERN.fullmatch(str(value)) if value is not None else None
        if match:
            return format_decimal(Decimal(match.group(1).replace(",", ".")))
    raise ValueError(
        f"could not find a load above row {before_row + 1}, "
        f"column {get_column_letter(step_column + 1)}"
    )


def parse_angle(value: object, cell_address: str) -> tuple[str, str]:
    if value is None or (isinstance(value, str) and not value.strip()):
        raise ValueError(f"blank angle at {cell_address}; no movement marker is required")
    if isinstance(value, str) and value.strip() in NO_MOVEMENT_VALUES:
        return "", "no_movement"
    try:
        angle = Decimal(str(value).replace(",", "."))
    except InvalidOperation as error:
        raise ValueError(f"invalid angle {value!r} at {cell_address}") from error
    if not angle.is_finite():
        raise ValueError(f"non-finite angle {value!r} at {cell_address}")
    if angle == 0:
        return "", "no_movement"
    return format(float(angle), ".12g"), "measured"


def parse_sheet(workbook_name: str, worksheet: object) -> SheetData:
    rows = [tuple(row) for row in worksheet.iter_rows(values_only=True)]
    if not rows:
        raise ValueError("worksheet is empty")

    measurement_number_row = find_measurement_number_row(rows)
    header_positions: list[tuple[int, int]] = []
    for row_index, row in enumerate(rows):
        for column_index in range(len(row) - 1):
            if (
                normalize_text(row[column_index]) == "step"
                and normalize_text(row[column_index + 1]) in {"angle [°]", "angle [deg]"}
            ):
                header_positions.append((row_index, column_index))
    if not header_positions:
        raise ValueError("could not find any Step / Angle [°] column pairs")

    section_rows = sorted({row_index for row_index, _ in header_positions})
    records: list[dict[str, str]] = []
    for header_row, step_column in header_positions:
        if measurement_number_row >= len(rows) or step_column >= len(rows[measurement_number_row]):
            raise ValueError("measurement number is missing for a Step / Angle block")
        measurement_no = rows[measurement_number_row][step_column]
        if measurement_no is None or not str(measurement_no).strip():
            raise ValueError(
                f"measurement number is missing in column {get_column_letter(step_column + 1)}"
            )

        direction = find_direction(rows, header_row, step_column)
        load_nm = find_load(rows, header_row, step_column)
        section_end = next(
            (row_index for row_index in section_rows if row_index > header_row), len(rows)
        )

        for row_index in range(header_row + 1, section_end):
            row = rows[row_index]
            if step_column + 1 >= len(row):
                continue
            step_value = row[step_column]
            match = STEP_PATTERN.fullmatch(str(step_value)) if step_value is not None else None
            if not match:
                continue
            step = f"{match.group(1).replace(',', '.')}-{match.group(2).replace(',', '.')}"
            angle_cell = f"{get_column_letter(step_column + 2)}{row_index + 1}"
            angle_deg, movement_status = parse_angle(row[step_column + 1], angle_cell)
            records.append(
                {
                    "load_nm": load_nm,
                    "measurement_no": format_scalar(measurement_no),
                    "direction": direction,
                    "step": step,
                    "angle_deg": angle_deg,
                    "movement_status": movement_status,
                }
            )

    if not records:
        raise ValueError("no measurement rows were found")
    return SheetData(workbook_name, worksheet.title, records)


def select_workbooks(workbooks: list[Path]) -> list[Path]:
    print("Available Excel workbooks:")
    for index, workbook_path in enumerate(workbooks, start=1):
        print(f"  {index}. {workbook_path.name}")
    while True:
        answer = input("Select workbook numbers separated by commas, 'all', or 'q' to quit: ").strip()
        if answer.casefold() == "q":
            return []
        if answer.casefold() == "all":
            return workbooks
        try:
            indices = [int(item.strip()) for item in answer.split(",")]
        except ValueError:
            print("Enter valid numbers, 'all', or 'q'.")
            continue
        if not indices or len(set(indices)) != len(indices) or any(
            index < 1 or index > len(workbooks) for index in indices
        ):
            print("Select each workbook once using the displayed numbers.")
            continue
        return [workbooks[index - 1] for index in indices]


def prompt_metadata() -> dict[str, str]:
    metadata: dict[str, str] = {}
    for field_name in METADATA_FIELDS:
        while True:
            value = input(f"  {field_name}: ").strip()
            if value:
                metadata[field_name] = value
                break
            print("  This field is required.")
    return metadata


def safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip()).strip("._-")
    return cleaned or "unknown"


def rows_for_group(group: OutputGroup) -> list[dict[str, str]]:
    output_rows: list[dict[str, str]] = []
    for sheet in group.sheets:
        for measurement in sheet.records:
            output_rows.append({**group.metadata, **measurement})
    return output_rows


def write_new_csv(path: Path, rows: list[dict[str, str]]) -> None:
    with path.open("x", encoding="utf-8", newline="") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=CSV_FIELDS, quoting=csv.QUOTE_ALL)
        writer.writeheader()
        writer.writerows(rows)


def validate_master_csv(path: Path) -> None:
    if not path.exists() or path.stat().st_size == 0:
        return
    with path.open("r", encoding="utf-8-sig", newline="") as csv_file:
        header = next(csv.reader(csv_file), None)
    if header != CSV_FIELDS:
        raise ValueError(f"existing CSV has an unexpected header: {path}")


def append_csv(path: Path, rows: list[dict[str, str]]) -> None:
    needs_header = not path.exists() or path.stat().st_size == 0
    with path.open("a", encoding="utf-8", newline="") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=CSV_FIELDS, quoting=csv.QUOTE_ALL)
        if needs_header:
            writer.writeheader()
        writer.writerows(rows)


def unique_output_path(group: OutputGroup, group_index: int) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = (
        f"{timestamp}_measurements_{safe_filename(group.metadata['test_id'])}_"
        f"{safe_filename(group.metadata['motor_id'])}"
    )
    path = OUTPUT_DIR / f"{base_name}.csv"
    suffix = 1
    while path.exists():
        path = OUTPUT_DIR / f"{base_name}_{suffix}.csv"
        suffix += 1
    return path


def main() -> int:
    workbooks = sorted(
        path for path in INPUT_DIR.glob("*.xlsx") if not path.name.startswith("~$")
    )
    if not workbooks:
        print(f"No Excel workbooks found in {INPUT_DIR}")
        return 1

    selected_workbooks = select_workbooks(workbooks)
    if not selected_workbooks:
        print("No workbooks selected.")
        return 0

    sheets: list[SheetData] = []
    try:
        for workbook_path in selected_workbooks:
            workbook = load_workbook(workbook_path, data_only=True, read_only=True)
            try:
                for worksheet in workbook.worksheets:
                    sheet_data = parse_sheet(workbook_path.name, worksheet)
                    sheets.append(sheet_data)
                    print(
                        f"Read {len(sheet_data.records)} measurement rows from "
                        f"{workbook_path.name} / {worksheet.title}"
                    )
            finally:
                workbook.close()
    except (OSError, ValueError, KeyError) as error:
        print(f"Could not parse workbook data: {error}", file=sys.stderr)
        return 1

    groups: list[OutputGroup] = []
    previous_group: OutputGroup | None = None
    for sheet in sheets:
        print(f"\nWorksheet: {sheet.workbook_name} / {sheet.sheet_name}")
        reuse_metadata = False
        if previous_group is not None:
            answer = input("Reuse metadata from the previous worksheet? [y/N]: ").strip().casefold()
            reuse_metadata = answer in {"y", "yes"}
        if reuse_metadata and previous_group is not None:
            group = previous_group
        else:
            group = OutputGroup(prompt_metadata())
            groups.append(group)
        group.sheets.append(sheet)
        previous_group = group

    aggregate_rows = [row for group in groups for row in rows_for_group(group)]
    try:
        validate_master_csv(MASTER_CSV)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        new_output_paths: list[tuple[Path, list[dict[str, str]]]] = []
        for group_index, group in enumerate(groups, start=1):
            group_rows = rows_for_group(group)
            output_path = unique_output_path(group, group_index)
            write_new_csv(output_path, group_rows)
            new_output_paths.append((output_path, group_rows))
        append_csv(MASTER_CSV, aggregate_rows)
    except (OSError, ValueError, csv.Error) as error:
        print(f"Could not update CSV outputs: {error}", file=sys.stderr)
        return 1

    print(f"\nCreated {len(new_output_paths)} new worksheet CSV files:")
    for path, rows in new_output_paths:
        print(f"  {path.name}: {len(rows)} rows")
    print(f"Appended {len(aggregate_rows)} rows to {MASTER_CSV}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())