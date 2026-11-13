from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from travel_angle_statistics import (
    DataValidationError,
    DuplicateStepError,
    analyze_records,
    empty_analysis_result,
    load_records_from_csv,
    write_statistics_workbook,
)


@dataclass
class StatisticsState:
    sources: list[Path]
    last_successful_refresh: datetime | None


def _resolve_paths(values: list[str]) -> list[Path]:
    return [Path(value).expanduser().resolve() for value in values]


def _load_state(state_path: Path) -> StatisticsState:
    if not state_path.exists():
        return StatisticsState(sources=[], last_successful_refresh=None)

    with state_path.open("r", encoding="utf-8") as state_file:
        payload = json.load(state_file)

    source_values = payload.get("sources", [])
    if not isinstance(source_values, list):
        raise ValueError(f"Invalid state file format in {state_path}")
    sources = [Path(value).expanduser().resolve() for value in source_values]

    timestamp_value = payload.get("last_successful_refresh")
    if timestamp_value is None:
        refresh_time = None
    else:
        refresh_time = datetime.fromisoformat(str(timestamp_value))

    return StatisticsState(sources=sources, last_successful_refresh=refresh_time)


def _save_state(state_path: Path, state: StatisticsState) -> None:
    state_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "sources": [str(source) for source in state.sources],
        "last_successful_refresh": (
            state.last_successful_refresh.isoformat()
            if state.last_successful_refresh is not None
            else None
        ),
    }
    with state_path.open("w", encoding="utf-8") as state_file:
        json.dump(payload, state_file, indent=2)


def _refresh_and_write(
    workbook_path: Path,
    sources: list[Path],
    stale_warning: str | None = None,
) -> datetime:
    records = load_records_from_csv(sources)
    analysis = analyze_records(records)
    refreshed_at = datetime.now()
    write_statistics_workbook(
        output_path=workbook_path,
        analysis=analysis,
        sources=sources,
        last_successful_refresh=refreshed_at,
        stale_warning=stale_warning,
    )
    return refreshed_at


def _write_empty_workbook(workbook_path: Path, sources: list[Path], warning: str | None) -> None:
    write_statistics_workbook(
        output_path=workbook_path,
        analysis=empty_analysis_result(),
        sources=sources,
        last_successful_refresh=None,
        stale_warning=warning,
    )


def _attempt_operation_with_recovery(
    operation_name: str,
    workbook_path: Path,
    state_path: Path,
    current_state: StatisticsState,
    proposed_sources: list[Path],
) -> int:
    try:
        refreshed_at = _refresh_and_write(workbook_path=workbook_path, sources=proposed_sources)
    except (DataValidationError, DuplicateStepError, OSError, ValueError) as error:
        print(f"{operation_name} failed: {error}", file=sys.stderr)
        if current_state.sources:
            try:
                _refresh_and_write(
                    workbook_path=workbook_path,
                    sources=current_state.sources,
                    stale_warning=(
                        "STALE — refresh failed for requested source change; restored previous sources"
                    ),
                )
            except (DataValidationError, DuplicateStepError, OSError, ValueError) as recovery_error:
                print(
                    f"Recovery failed: {recovery_error}. Previous workbook remains the last known result.",
                    file=sys.stderr,
                )
                return 1

            print("Requested change failed. Previous sources were restored and refreshed.")
            return 1

        _write_empty_workbook(
            workbook_path=workbook_path,
            sources=[],
            warning=f"STALE — {operation_name.lower()} failed; no active sources",
        )
        current_state.sources = []
        current_state.last_successful_refresh = None
        _save_state(state_path, current_state)
        return 1

    current_state.sources = proposed_sources
    current_state.last_successful_refresh = refreshed_at
    _save_state(state_path, current_state)
    print(
        f"{operation_name} succeeded. Refreshed {len(proposed_sources)} source file(s) into {workbook_path}."
    )
    return 0


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Create and refresh travel-angle statistics workbook from one or more parser CSV files."
        )
    )
    parser.add_argument(
        "--workbook",
        default="statistics_generated.xlsx",
        help="Path to output workbook (.xlsx or .xlsm).",
    )
    parser.add_argument(
        "--state",
        default="statistics_state.json",
        help="Path to JSON file storing active CSV sources and refresh metadata.",
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("refresh", help="Refresh using currently active source list.")

    add_parser = subparsers.add_parser("add", help="Add source CSV files and refresh.")
    add_parser.add_argument("files", nargs="+", help="CSV file paths to add.")

    replace_parser = subparsers.add_parser(
        "replace", help="Replace source list with provided CSV files and refresh."
    )
    replace_parser.add_argument("files", nargs="+", help="CSV file paths to use.")

    remove_parser = subparsers.add_parser("remove", help="Remove source CSV files and refresh.")
    remove_parser.add_argument("files", nargs="+", help="CSV file paths to remove.")

    subparsers.add_parser("list-sources", help="Show active source list.")
    return parser


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    workbook_path = Path(args.workbook).expanduser().resolve()
    state_path = Path(args.state).expanduser().resolve()
    state = _load_state(state_path)

    if args.command == "list-sources":
        if not state.sources:
            print("No active CSV sources.")
        else:
            print("Active CSV sources:")
            for source in state.sources:
                print(f"  - {source}")
        return 0

    if args.command == "refresh":
        if not state.sources:
            print("No active CSV sources. Add files first.")
            _write_empty_workbook(
                workbook_path=workbook_path,
                sources=[],
                warning="No active CSV sources. Add files to start analysis.",
            )
            state.last_successful_refresh = None
            _save_state(state_path, state)
            return 0
        return _attempt_operation_with_recovery(
            operation_name="Refresh",
            workbook_path=workbook_path,
            state_path=state_path,
            current_state=state,
            proposed_sources=state.sources,
        )

    if args.command == "replace":
        proposed_sources = _resolve_paths(args.files)
        return _attempt_operation_with_recovery(
            operation_name="Replace",
            workbook_path=workbook_path,
            state_path=state_path,
            current_state=state,
            proposed_sources=proposed_sources,
        )

    if args.command == "add":
        incoming = _resolve_paths(args.files)
        existing = set(state.sources)
        skipped = [path for path in incoming if path in existing]
        new_sources = [path for path in incoming if path not in existing]
        if skipped:
            print("Skipped already active files:")
            for path in skipped:
                print(f"  - {path}")
        if not new_sources:
            print("No new files to add.")
            return 0
        proposed_sources = [*state.sources, *new_sources]
        return _attempt_operation_with_recovery(
            operation_name="Add",
            workbook_path=workbook_path,
            state_path=state_path,
            current_state=state,
            proposed_sources=proposed_sources,
        )

    if args.command == "remove":
        to_remove = set(_resolve_paths(args.files))
        proposed_sources = [path for path in state.sources if path not in to_remove]
        removed_count = len(state.sources) - len(proposed_sources)
        if removed_count == 0:
            print("No matching active files to remove.")
            return 0
        if not proposed_sources:
            _write_empty_workbook(
                workbook_path=workbook_path,
                sources=[],
                warning="No active CSV sources. Results were cleared.",
            )
            state.sources = []
            state.last_successful_refresh = None
            _save_state(state_path, state)
            print("Removed last active source. Cleared results and refresh timestamp.")
            return 0
        return _attempt_operation_with_recovery(
            operation_name="Remove",
            workbook_path=workbook_path,
            state_path=state_path,
            current_state=state,
            proposed_sources=proposed_sources,
        )

    parser.error(f"Unsupported command: {args.command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
