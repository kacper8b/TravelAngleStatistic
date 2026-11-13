from __future__ import annotations

from dataclasses import dataclass


class StatisticsError(Exception):
    pass


class DataValidationError(StatisticsError):
    pass


@dataclass(frozen=True)
class DuplicateConflictRecord:
    source_file: str
    row_number: int
    angle_deg: str
    movement_status: str


class DuplicateStepError(DataValidationError):
    def __init__(self, message: str) -> None:
        super().__init__(message)
