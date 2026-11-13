from __future__ import annotations

import argparse
from pathlib import Path

import win32com.client as win32


XL_OPEN_XML_MACRO_ENABLED = 52


def _load_module_code(module_path: Path) -> str:
    return module_path.read_text(encoding="utf-8")


def _create_data_sources_layout(sheet, script_path: str, state_path: str, generated_path: str) -> None:
    def set_cell(row: int, column: int, value: str) -> None:
        cell = sheet.Cells(row, column)
        cell.Value = value

    sheet.Name = "Data sources"
    set_cell(1, 1, "TravelAngleStatistic - Data sources")
    set_cell(2, 1, "Python command")
    set_cell(2, 2, "python")
    set_cell(3, 1, "Engine script path")
    set_cell(3, 2, script_path)
    set_cell(4, 1, "State file path")
    set_cell(4, 2, state_path)
    set_cell(5, 1, "Generated workbook path")
    set_cell(5, 2, generated_path)
    set_cell(6, 1, "Last command status")
    set_cell(6, 2, "")
    set_cell(7, 1, "Active source summary")
    set_cell(8, 1, "Last successful refresh")
    set_cell(9, 1, "Warnings")
    set_cell(10, 1, "Active sources list")
    sheet.Columns("A").ColumnWidth = 34
    sheet.Columns("B").ColumnWidth = 110

    def add_action_shape(left: int, top: int, width: int, height: int, caption: str, action: str) -> None:
        shape = sheet.Shapes.AddShape(1, left, top, width, height)
        shape.TextFrame.Characters().Text = caption
        shape.OnAction = action

    add_action_shape(640, 24, 110, 28, "Add CSV", "AddSourcesButton")
    add_action_shape(760, 24, 110, 28, "Replace CSV", "ReplaceSourcesButton")
    add_action_shape(640, 58, 110, 28, "Remove selected", "RemoveSelectedSourceButton")
    add_action_shape(760, 58, 110, 28, "Refresh", "RefreshSourcesButton")


def build_workbook(output_path: Path, script_path: Path, state_path: Path, generated_path: Path) -> None:
    module_path = Path(__file__).resolve().parent / "excel_vba" / "statistics_ui.bas"
    module_code = _load_module_code(module_path)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    workbook = None
    try:
        excel.Workbooks.Add()
        workbook = excel.ActiveWorkbook
        if workbook is None:
            raise RuntimeError("Excel did not return an active workbook after Workbooks.Add().")
        while workbook.Worksheets.Count > 3:
            workbook.Worksheets(workbook.Worksheets.Count).Delete()
        while workbook.Worksheets.Count < 3:
            workbook.Worksheets.Add(None, workbook.Worksheets.Item(workbook.Worksheets.Count))

        data_sheet = workbook.Worksheets.Item(1)
        _create_data_sources_layout(
            data_sheet,
            script_path=str(script_path),
            state_path=str(state_path),
            generated_path=str(generated_path),
        )

        workbook.Worksheets.Item(2).Name = "Movement angle"
        workbook.Worksheets.Item(3).Name = "Tilt counts"
        workbook.Worksheets.Add(None, workbook.Worksheets.Item(workbook.Worksheets.Count))
        workbook.Worksheets.Item(workbook.Worksheets.Count).Name = "Final position"

        standard_module = workbook.VBProject.VBComponents.Add(1)
        standard_module.Name = "StatisticsUI"
        standard_module.CodeModule.AddFromString(module_code)

        this_workbook_code = (
            "Option Explicit\n\n"
            "Private Sub Workbook_Open()\n"
            "    StatisticsWorkbook_Open\n"
            "End Sub\n\n"
            "Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)\n"
            "    StatisticsWorkbook_BeforeSave\n"
            "End Sub\n"
        )
        workbook.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString(this_workbook_code)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.SaveAs(str(output_path), FileFormat=XL_OPEN_XML_MACRO_ENABLED)
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()


def main() -> int:
    parser = argparse.ArgumentParser(description="Build macro-enabled statistics workbook template.")
    parser.add_argument(
        "--output",
        default="data/statistics.xlsm",
        help="Path to output .xlsm workbook.",
    )
    parser.add_argument(
        "--script",
        default="create_statistics.py",
        help="Path to create_statistics.py used by workbook macros.",
    )
    parser.add_argument(
        "--state",
        default="data/statistics_state.json",
        help="Path to state file used by workbook macros.",
    )
    parser.add_argument(
        "--generated",
        default="data/statistics_generated.xlsx",
        help="Path to generated analysis workbook used by workbook macros.",
    )
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    output_path = (base_dir / args.output).resolve()
    script_path = (base_dir / args.script).resolve()
    state_path = (base_dir / args.state).resolve()
    generated_path = (base_dir / args.generated).resolve()

    build_workbook(
        output_path=output_path,
        script_path=script_path,
        state_path=state_path,
        generated_path=generated_path,
    )
    print(f"Workbook created: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
