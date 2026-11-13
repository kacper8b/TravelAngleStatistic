from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal

from .constants import MOVEMENT_GROUP_BY_DIRECTION_STEP


@dataclass(frozen=True)
class SeriesKey:
    test_id: str
    motor_id: str
    motor_name: str
    motor_variant: str
    software_version: str
    load_nm: Decimal
    measurement_no: str


@dataclass(frozen=True)
class BaseConfigKey:
    motor_name: str
    motor_variant: str
    software_version: str
    load_nm: Decimal


@dataclass(frozen=True)
class MeasurementRecord:
    source_file: str
    row_number: int
    test_id: str
    motor_id: str
    motor_name: str
    motor_variant: str
    software_version: str
    load_nm: Decimal
    measurement_no: str
    direction: str
    step: str
    angle_deg: Decimal | None
    movement_status: str

    @property
    def series_key(self) -> SeriesKey:
        return SeriesKey(
            test_id=self.test_id,
            motor_id=self.motor_id,
            motor_name=self.motor_name,
            motor_variant=self.motor_variant,
            software_version=self.software_version,
            load_nm=self.load_nm,
            measurement_no=self.measurement_no,
        )

    @property
    def base_config_key(self) -> BaseConfigKey:
        return BaseConfigKey(
            motor_name=self.motor_name,
            motor_variant=self.motor_variant,
            software_version=self.software_version,
            load_nm=self.load_nm,
        )

    @property
    def movement_group(self) -> str:
        return MOVEMENT_GROUP_BY_DIRECTION_STEP[(self.direction, self.step)]
