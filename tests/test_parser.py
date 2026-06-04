from datetime import datetime

from app.parser import parse_workbook
from app.parser.constants import ROLE_OTHER, ROLE_PSA, ROLE_RN, SHIFT_DAY, SHIFT_NIGHT
from tests.fixtures.source_workbooks import build_source_workbook


def test_parser_detects_repeated_sections_and_unknown_codes():
    data = build_source_workbook(
        [
            {
                "schedule": [
                    {
                        "job": "3178-Staff RN Placeholder",
                        "employee": "Employee 001",
                        "start": datetime(2026, 4, 28, 7, 0),
                        "end": datetime(2026, 4, 28, 19, 30),
                    },
                    {
                        "job": "9999-Other Placeholder",
                        "employee": "Employee 002",
                        "start": datetime(2026, 4, 28, 11, 0),
                        "end": datetime(2026, 4, 28, 15, 0),
                    },
                ]
            },
            {
                "schedule": [
                    {
                        "job": "5056-Support Placeholder",
                        "employee": "Employee 003",
                        "start": datetime(2026, 4, 28, 19, 0),
                        "end": datetime(2026, 4, 29, 7, 30),
                    }
                ]
            },
        ]
    )

    parsed = parse_workbook(data)

    assert len(parsed.records) == 3
    by_employee = {record.employee_display_name: record for record in parsed.records}
    assert by_employee["Employee 001"].role_group == ROLE_RN
    assert by_employee["Employee 001"].shift_kind == SHIFT_DAY
    assert by_employee["Employee 002"].role_group == ROLE_OTHER
    assert by_employee["Employee 003"].role_group == ROLE_PSA
    assert by_employee["Employee 003"].shift_kind == SHIFT_NIGHT


def test_parser_excludes_staff_lpn_rows():
    data = build_source_workbook(
        [
            {
                "schedule": [
                    {
                        "job": "3178-Staff RN Placeholder",
                        "employee": "Employee 001",
                        "start": datetime(2026, 4, 28, 7, 0),
                        "end": datetime(2026, 4, 28, 19, 30),
                    },
                    {
                        "job": "4120-Staff LPN",
                        "employee": "Employee 002",
                        "start": datetime(2026, 4, 28, 7, 0),
                        "end": datetime(2026, 4, 28, 19, 30),
                    },
                ]
            }
        ]
    )

    parsed = parse_workbook(data)

    assert [record.employee_display_name for record in parsed.records] == ["Employee 001"]


def test_parser_includes_lead_pay_only_rows_with_nearby_context():
    data = build_source_workbook(
        [
            {
                "schedule": [
                    {
                        "job": "3178-Staff RN Placeholder",
                        "employee": "Employee 001",
                        "start": datetime(2026, 4, 28, 19, 0),
                        "end": datetime(2026, 4, 29, 7, 30),
                    }
                ],
                "pay_codes": [
                    {
                        "job": "3178-Staff RN Placeholder",
                        "employee": "Employee 002",
                        "pay_code": "LP - LEAD PAY",
                    }
                ],
            }
        ]
    )

    parsed = parse_workbook(data)

    assert len(parsed.records) == 2
    lead_pay_record = next(record for record in parsed.records if record.employee_display_name == "Employee 002")
    assert lead_pay_record.has_lead_pay is True
    assert lead_pay_record.work_date.isoformat() == "2026-04-28"
    assert lead_pay_record.shift_kind == SHIFT_NIGHT
    assert lead_pay_record.warnings == ["Lead-pay-only row was included using nearby schedule context."]
