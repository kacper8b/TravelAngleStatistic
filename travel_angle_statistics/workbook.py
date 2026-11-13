from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from .analysis import AnalysisResult
from .constants import STATUS_NOT_EVALUABLE


def _cell_value(value: Any) -> Any:
    if isinstance(value, Decimal):
        return float(value)
    return value


def _write_table(
    worksheet: Any,
    start_row: int,
    start_column: int,
    headers: list[str],
    rows: list[dict[str, Any]],
    table_name: str | None = None,
) -> int:
    header_font = Font(bold=True)
    for index, header in enumerate(headers):
        cell = worksheet.cell(row=start_row, column=start_column + index, value=header)
        cell.font = header_font

    current_row = start_row + 1
    for row in rows:
        for index, header in enumerate(headers):
            value = row.get(header, "")
            cell = worksheet.cell(
                row=current_row,
                column=start_column + index,
                value=_cell_value(value),
            )
            if isinstance(value, Decimal):
                cell.number_format = "0.00"
            elif header.endswith("_pct") and isinstance(value, (int, float)):
                cell.number_format = "0.00"
        current_row += 1

    end_row = max(start_row + 1, current_row - 1)
    end_column = start_column + len(headers) - 1
    table_range = (
        f"{get_column_letter(start_column)}{start_row}:"
        f"{get_column_letter(end_column)}{end_row}"
    )
    if table_name is not None and rows:
        table = Table(displayName=table_name, ref=table_range)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        worksheet.add_table(table)
    for index, header in enumerate(headers):
        worksheet.column_dimensions[get_column_letter(start_column + index)].width = max(
            len(header) + 2,
            16,
        )
    return current_row


def _set_warning_banner(worksheet: Any, warning: str | None) -> None:
    if warning:
        worksheet["A1"] = warning
        worksheet["A1"].font = Font(color="9C0006", bold=True)
        worksheet["A1"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    else:
        worksheet["A1"] = ""
        worksheet["A1"].font = Font(bold=False)
        worksheet["A1"].fill = PatternFill(fill_type=None)


def _write_data_sources_sheet(
    workbook: Workbook,
    sources: list[Path],
    last_successful_refresh: datetime | None,
    stale_warning: str | None,
) -> None:
    sheet = workbook.create_sheet(title="Data sources")
    sheet["A1"] = "Active CSV sources"
    sheet["A1"].font = Font(bold=True)
    sheet["A2"] = f"Count: {len(sources)}"
    sheet["A3"] = "Last successful refresh"
    sheet["B3"] = (
        last_successful_refresh.strftime("%Y-%m-%d %H:%M:%S")
        if last_successful_refresh is not None
        else STATUS_NOT_EVALUABLE
    )

    if stale_warning is not None:
        sheet["A5"] = stale_warning
        sheet["A5"].font = Font(color="9C0006", bold=True)
        sheet["A5"].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    row = 7
    for source in sources:
        sheet.cell(row=row, column=1, value=str(source))
        row += 1
    sheet.column_dimensions["A"].width = 120
    sheet.column_dimensions["B"].width = 30


def _write_movement_angle_sheet(workbook: Workbook, analysis: AnalysisResult) -> None:
    sheet = workbook.create_sheet(title="Movement angle")
    _set_warning_banner(sheet, None)
    headers = [
        "motor_name",
        "motor_variant",
        "software_version",
        "load_nm",
        "movement_group",
        "measurement_count",
        "contributing_series_count",
        "incomplete_contributing_series_count",
        "mean_angle_deg",
        "median_angle_deg",
        "min_angle_deg",
        "max_angle_deg",
        "sample_stddev_deg",
        "mean_requirement",
        "max_requirement",
        "exceedance_count",
    ]
    _write_table(
        sheet,
        3,
        1,
        headers,
        analysis.movement_angle_rows,
        table_name="MovementAngleTable",
    )


def _write_tilt_counts_sheet(workbook: Workbook, analysis: AnalysisResult) -> None:
    sheet = workbook.create_sheet(title="Tilt counts")
    _set_warning_banner(sheet, None)
    overall_headers = [
        "performed_up",
        "performed_down",
        "performed_total",
        "expected_up",
        "expected_down",
        "expected_total",
        "performed_up_pct",
        "performed_down_pct",
        "performed_total_pct",
    ]
    sheet["A3"] = "Overall"
    sheet["A3"].font = Font(bold=True)
    _write_table(
        sheet,
        4,
        1,
        overall_headers,
        [analysis.tilt_overall_row],
        table_name="TiltOverallTable",
    )

    series_headers = [
        "test_id",
        "motor_id",
        "motor_name",
        "motor_variant",
        "software_version",
        "load_nm",
        "measurement_no",
        "is_complete",
        "performed_up",
        "performed_down",
        "performed_total",
        "up_requirement",
        "down_requirement",
        "performed_up_pct",
        "performed_down_pct",
        "performed_total_pct",
    ]
    sheet["A8"] = "Series"
    sheet["A8"].font = Font(bold=True)
    next_row = _write_table(
        sheet,
        9,
        1,
        series_headers,
        analysis.tilt_series_rows,
        table_name="TiltSeriesTable",
    )

    no_movement_headers = [
        "test_id",
        "motor_id",
        "motor_name",
        "motor_variant",
        "software_version",
        "load_nm",
        "measurement_no",
        "direction",
        "step",
    ]
    section_row = next_row + 1
    sheet.cell(row=section_row, column=1, value="No movement").font = Font(bold=True)
    _write_table(
        sheet,
        section_row + 1,
        1,
        no_movement_headers,
        analysis.tilt_no_movement_rows,
        table_name="TiltNoMovementTable",
    )


def _write_final_position_sheet(workbook: Workbook, analysis: AnalysisResult) -> None:
    sheet = workbook.create_sheet(title="Final position")
    _set_warning_banner(sheet, None)
    series_headers = [
        "test_id",
        "motor_id",
        "motor_name",
        "motor_variant",
        "software_version",
        "load_nm",
        "measurement_no",
        "is_complete",
        "sum_up_deg",
        "sum_down_deg",
        "deviation_deg",
        "deviation_requirement",
        "partial",
    ]
    sheet["A3"] = "Series"
    sheet["A3"].font = Font(bold=True)
    next_row = _write_table(
        sheet,
        4,
        1,
        series_headers,
        analysis.final_position_series_rows,
        table_name="FinalPositionSeriesTable",
    )

    summary_headers = [
        "motor_name",
        "motor_variant",
        "software_version",
        "load_nm",
        "evaluated_series_count",
        "incomplete_evaluated_series_count",
        "exceedance_count",
        "exceedance_pct",
    ]
    section_row = next_row + 1
    sheet.cell(row=section_row, column=1, value="Summary").font = Font(bold=True)
    _write_table(
        sheet,
        section_row + 1,
        1,
        summary_headers,
        analysis.final_position_summary_rows,
        table_name="FinalPositionSummaryTable",
    )


def write_statistics_workbook(
    output_path: Path,
    analysis: AnalysisResult,
    sources: list[Path],
    last_successful_refresh: datetime | None,
    stale_warning: str | None = None,
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    _write_data_sources_sheet(
        workbook=workbook,
        sources=sources,
        last_successful_refresh=last_successful_refresh,
        stale_warning=stale_warning,
    )
    _write_movement_angle_sheet(workbook=workbook, analysis=analysis)
    _write_tilt_counts_sheet(workbook=workbook, analysis=analysis)
    _write_final_position_sheet(workbook=workbook, analysis=analysis)

    if stale_warning is not None:
        for sheet_name in ("Movement angle", "Tilt counts", "Final position"):
            _set_warning_banner(workbook[sheet_name], stale_warning)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
