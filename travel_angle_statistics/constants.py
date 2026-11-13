from __future__ import annotations

from decimal import Decimal


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

EXPECTED_STEPS_BY_DIRECTION = {
    "Up": [
        "90-80",
        "80-70",
        "70-60",
        "60-50",
        "50-40",
        "40-30",
        "30-20",
        "20-10",
        "10-0",
    ],
    "Down": [
        "0-10",
        "10-20",
        "20-30",
        "30-40",
        "40-50",
        "50-60",
        "60-70",
        "70-80",
        "80-90",
    ],
}

MOVEMENT_GROUP_UP_100_FIRST = "Up 100ms - first movement"
MOVEMENT_GROUP_UP_100 = "Up 100ms"
MOVEMENT_GROUP_UP_150 = "Up 150ms"
MOVEMENT_GROUP_DOWN_100_FIRST = "Down 100ms - first movement"
MOVEMENT_GROUP_DOWN_100 = "Down 100ms"
MOVEMENT_GROUP_DOWN_150 = "Down 150ms"
MOVEMENT_GROUP_DOWN_150_LAST = "Down 150ms - last movement"

MOVEMENT_GROUPS = [
    MOVEMENT_GROUP_UP_100_FIRST,
    MOVEMENT_GROUP_UP_100,
    MOVEMENT_GROUP_UP_150,
    MOVEMENT_GROUP_DOWN_100_FIRST,
    MOVEMENT_GROUP_DOWN_100,
    MOVEMENT_GROUP_DOWN_150,
    MOVEMENT_GROUP_DOWN_150_LAST,
]

MOVEMENT_GROUP_BY_DIRECTION_STEP = {
    ("Up", "90-80"): MOVEMENT_GROUP_UP_100_FIRST,
    ("Up", "80-70"): MOVEMENT_GROUP_UP_100,
    ("Up", "70-60"): MOVEMENT_GROUP_UP_100,
    ("Up", "60-50"): MOVEMENT_GROUP_UP_100,
    ("Up", "50-40"): MOVEMENT_GROUP_UP_150,
    ("Up", "40-30"): MOVEMENT_GROUP_UP_100,
    ("Up", "30-20"): MOVEMENT_GROUP_UP_100,
    ("Up", "20-10"): MOVEMENT_GROUP_UP_100,
    ("Up", "10-0"): MOVEMENT_GROUP_UP_150,
    ("Down", "0-10"): MOVEMENT_GROUP_DOWN_100_FIRST,
    ("Down", "10-20"): MOVEMENT_GROUP_DOWN_100,
    ("Down", "20-30"): MOVEMENT_GROUP_DOWN_100,
    ("Down", "30-40"): MOVEMENT_GROUP_DOWN_100,
    ("Down", "40-50"): MOVEMENT_GROUP_DOWN_150,
    ("Down", "50-60"): MOVEMENT_GROUP_DOWN_100,
    ("Down", "60-70"): MOVEMENT_GROUP_DOWN_100,
    ("Down", "70-80"): MOVEMENT_GROUP_DOWN_100,
    ("Down", "80-90"): MOVEMENT_GROUP_DOWN_150_LAST,
}

MOVEMENT_GROUP_KIND = {
    MOVEMENT_GROUP_UP_100_FIRST: "100ms",
    MOVEMENT_GROUP_UP_100: "100ms",
    MOVEMENT_GROUP_UP_150: "150ms",
    MOVEMENT_GROUP_DOWN_100_FIRST: "100ms",
    MOVEMENT_GROUP_DOWN_100: "100ms",
    MOVEMENT_GROUP_DOWN_150: "150ms",
    MOVEMENT_GROUP_DOWN_150_LAST: "150ms",
}

ANGLE_REQUIREMENT_LOAD = Decimal("0.1")
FINAL_POSITION_REQUIREMENT_LOADS = {
    Decimal("0.1"),
    Decimal("2"),
    Decimal("4"),
    Decimal("8"),
}

STATUS_PASS = "PASS"
STATUS_FAIL = "FAIL"
STATUS_NA = "N/A"
STATUS_NOT_EVALUABLE = "NOT EVALUABLE"
STATUS_PARTIAL = "PARTIAL"
