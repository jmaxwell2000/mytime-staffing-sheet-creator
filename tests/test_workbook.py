from datetime import date, time
from io import BytesIO

from openpyxl import load_workbook

from app.parser.constants import ROLE_PSA, ROLE_RN, SHIFT_DAY, SHIFT_NIGHT
from app.parser.models import ScheduleRecord
from app.workbook.build_workbook import (
    BACK_LINE_ROWS,
    BACK_PENCIL_REMINDER,
    COLUMN_WIDTHS,
    FALL_PREVENTION_TEXT,
    FRONT_PENCIL_REMINDER,
    build_staffing_workbook,
    resolve_lead_display,
)

EXPECTED_FALL_PREVENTION_TEXT = "Post-fall measures (circle):              Alarm           Nonskid socks             Fall sign             Stay with pt toileting             Fall risk score             Other:____________"


def record(
    name: str,
    job_code: str,
    *,
    role_group: str = ROLE_RN,
    work_date: date = date(2026, 4, 28),
    shift_kind: str = SHIFT_NIGHT,
    start: time = time(19, 0),
    end: time = time(7, 30),
    clinical_leader: bool = False,
    lead_pay: bool = False,
) -> ScheduleRecord:
    return ScheduleRecord(
        employee_display_name=name,
        work_date=work_date,
        shift_kind=shift_kind,
        start_time=start,
        end_time=end,
        job_code=job_code,
        role_group=role_group,
        is_clinical_leader=clinical_leader,
        has_lead_pay=lead_pay,
    )


def test_clinical_leader_overrides_lead_pay_regression():
    records = [
        record("Employee 001", "3178", lead_pay=True),
        record("Employee 002", "8536", clinical_leader=True),
    ]

    output = build_staffing_workbook(records)
    workbook = load_workbook(BytesIO(output))
    sheet = workbook["Tue Apr28 NIGHT"]

    lead_cells = [cell.coordinate for row in sheet.iter_rows() for cell in row if cell.value == "LEAD"]
    assert lead_cells == ["F3"]
    assert sheet["C3"].value == "Employee 002"
    assert sheet["C4"].value == "Employee 001"
    assert sheet["F3"].alignment.horizontal == "center"


def test_names_are_normalized_and_sorted_with_lead_rn_first():
    records = [
        record("bETAUSER, bRAVO", "3178"),
        record("zETAUSER, aLPHA", "8536", clinical_leader=True),
        record("aLPHAUSER, cHARLIE", "3178"),
        record("pSAUSER, dELTA", "5056", role_group=ROLE_PSA),
    ]

    output = build_staffing_workbook(records)
    workbook = load_workbook(BytesIO(output))
    sheet = workbook["Tue Apr28 NIGHT"]

    assert sheet["C3"].value == "Zetauser, Alpha"
    assert sheet["F3"].value == "LEAD"
    assert sheet["C4"].value == "Alphauser, Charlie"
    assert sheet["C5"].value == "Betauser, Bravo"
    assert sheet["C19"].value == "Psauser, Delta"


def test_multiple_clinical_leaders_get_cl_suffix_and_no_lead():
    records = [
        record("Employee 001", "8536", clinical_leader=True),
        record("Employee 002", "8536", clinical_leader=True),
    ]

    display = resolve_lead_display(records)

    assert display[id(records[0])] == ("Employee 001 (CL)", False)
    assert display[id(records[1])] == ("Employee 002 (CL)", False)


def test_workbook_builder_excludes_staff_lpn_records_defensively():
    records = [
        record("Employee 001", "3178"),
        record("Employee 002", "4120"),
    ]

    output = build_staffing_workbook(records)
    workbook = load_workbook(BytesIO(output))
    sheet = workbook["Tue Apr28 NIGHT"]

    visible_values = [cell.value for row in sheet.iter_rows() for cell in row if cell.value is not None]
    assert "Employee 001" in visible_values
    assert "Employee 002" not in visible_values


def test_staff_name_font_size_uses_final_cl_display_text():
    records = [
        record("aBCDEFGHIJ, kLMNO", "8536", clinical_leader=True),
        record("pQRSTUVWXYZ, qRSTU", "8536", clinical_leader=True),
    ]

    output = build_staffing_workbook(records)
    workbook = load_workbook(BytesIO(output))
    sheet = workbook["Tue Apr28 NIGHT"]

    assert sheet["C3"].value == "Abcdefghij, Klmno (CL)"
    assert sheet["C3"].font.sz == 9
    assert sheet["C3"].alignment.shrinkToFit is True
    assert sheet["C4"].value == "Pqrstuvwxyz, Qrstu (CL)"
    assert sheet["C4"].font.sz == 9
    assert sheet["C4"].alignment.shrinkToFit is True


def test_single_cl_name_stays_normal_size_when_displayed_as_lead():
    records = [
        record("aBCDEFGHIJ, kLMNO", "8536", clinical_leader=True),
        record("pQRSTUVWXYZ, qRSTU", "3178"),
    ]

    output = build_staffing_workbook(records)
    workbook = load_workbook(BytesIO(output))
    sheet = workbook["Tue Apr28 NIGHT"]

    assert sheet["C3"].value == "Abcdefghij, Klmno"
    assert sheet["C3"].font.sz == 11
    assert sheet["F3"].value == "LEAD"


def test_workbook_front_and_back_structure_and_time_styling():
    records = [
        record(
            "Employee 001",
            "3178",
            work_date=date(2026, 4, 29),
            shift_kind=SHIFT_DAY,
            start=time(11, 0),
            end=time(15, 0),
        ),
        record(
            "Employee 002",
            "5056",
            role_group=ROLE_PSA,
            work_date=date(2026, 4, 29),
            shift_kind=SHIFT_DAY,
            start=time(7, 0),
            end=time(19, 0),
        ),
    ]

    output = build_staffing_workbook(records, include_back_pages=True, back_page_offset=0.05)
    workbook = load_workbook(BytesIO(output))

    assert workbook.sheetnames == ["Wed Apr29 DAY", "Wed Apr29 DAY BACK"]
    front = workbook["Wed Apr29 DAY"]
    back = workbook["Wed Apr29 DAY BACK"]
    assert str(front.print_area) == "'Wed Apr29 DAY'!$B$1:$H$38"
    assert front.page_setup.orientation == "portrait"
    assert front.page_setup.fitToWidth == 1
    assert front.page_setup.fitToHeight == 1
    assert front.page_margins.left == 0.75
    assert front.page_margins.right == 0.12
    assert str(front.print_area) == "'Wed Apr29 DAY'!$B$1:$H$38"
    assert front.column_dimensions["B"].width == COLUMN_WIDTHS["B"]
    assert front.column_dimensions["F"].width == COLUMN_WIDTHS["F"]
    assert back.column_dimensions["B"].width == front.column_dimensions["B"].width
    assert back.column_dimensions["F"].width == front.column_dimensions["F"].width
    assert round(sum(front.column_dimensions[column].width for column in "BCDEFGH"), 2) == 101.49
    assert front["B1"].value == FRONT_PENCIL_REMINDER
    assert front["B1"].alignment.horizontal == "left"
    assert front["G1"].value == "Wed 4/29 - DAY"
    assert front["G1"].alignment.horizontal == "right"
    assert back["B1"].value == "Wed 4/29 - DAY"
    assert back["B1"].alignment.horizontal == "left"
    assert back["G1"].value == BACK_PENCIL_REMINDER
    assert back["G1"].alignment.horizontal == "right"
    assert front["B2"].value == "Initial"
    assert front["B2"].font.sz == 7
    assert front["B2"].alignment.shrinkToFit is True
    assert front["C18"].value == "PSA"
    assert front["B18"].value == "Initial"
    assert front["B30"].value.startswith("Patient Fall Review (Wed 4/29) - Charge RN:")
    assert front["B30"].font.bold is True
    assert front["B30"].font.sz == 10
    assert front["B30"].alignment.horizontal == "center"
    assert front.row_dimensions[30].height == 26
    assert front.row_dimensions[32].height == 30
    assert front.row_dimensions[36].height == 30
    assert "Yes/No" not in front["B30"].value
    assert "answer the questions" in front["B30"].value
    assert front["B31"].value == "Question"
    assert front["B31"].font.sz == 10
    assert front["E31"].value is None
    assert front["F31"].value == "Answer / Notes"
    assert front["G31"].value is None
    assert front["F31"].font.sz == 10
    assert front["B32"].value == "\u2022 Bathroom or bedside commode?"
    assert front["B32"].font.sz == 10
    assert front["B32"].alignment.indent == 1
    assert front["F32"].value is None
    assert front["F32"].border.bottom.style == "thin"
    assert front["B33"].value == "\u2022 Staff or family with patient at time of fall?"
    assert front["B34"].value == "\u2022 If no, should someone have been present?"
    assert front["B35"].value == "\u2022 Any injury?"
    assert FALL_PREVENTION_TEXT == EXPECTED_FALL_PREVENTION_TEXT
    assert front["B36"].value == EXPECTED_FALL_PREVENTION_TEXT
    assert "\n" not in front["B36"].value
    assert "|" not in front["B36"].value
    assert front["B36"].font.sz == 10
    assert front["B36"].font.bold is False
    assert front["B36"].alignment.horizontal == "left"
    assert front["B36"].alignment.wrapText is not True
    assert front["B36"].alignment.shrinkToFit is not True
    assert front["D3"].value == 1100
    assert front["D3"].number_format == "0000"
    assert front["D3"].font.bold is False
    assert front["D3"].font.italic is False
    assert front["D3"].font.sz == 11
    assert _rgb(front["D3"].fill.fgColor) == "FFF2CC"
    assert front["E3"].value == 1500
    assert front["E3"].number_format == "0000"
    assert front["E3"].font.bold is False
    assert front["E3"].font.italic is False
    assert front["E3"].font.sz == 11
    assert _rgb(front["E3"].fill.fgColor) == "FFF2CC"
    assert front["E19"].value == 1930
    assert front["E19"].number_format == "0000"
    assert back.page_margins.left == 0.12
    assert back.page_margins.right == 0.75
    assert back.page_margins.top == 0.2
    for row in range(1, 39):
        assert back.row_dimensions[row].height == front.row_dimensions[row].height

    for row in range(2, 37):
        for column in range(2, 9):
            border = front.cell(row, column).border
            for side in (border.left, border.right, border.top, border.bottom):
                assert side.style in (None, "thin")

    front_grid_rows = [
        row
        for row in range(2, 30)
        if any(front.cell(row, column).border.bottom.style for column in range(2, 9))
    ]
    back_line_rows = [
        row
        for row in range(2, 30)
        if any(back.cell(row, column).border.bottom.style for column in range(2, 9))
    ]
    assert tuple(back_line_rows) == BACK_LINE_ROWS
    assert 2 not in back_line_rows
    assert all(row in front_grid_rows for row in back_line_rows)
    assert back["B1"].font.bold == front["G1"].font.bold
    assert back["B1"].font.sz == front["G1"].font.sz
    assert back["B38"].font.bold == front["B38"].font.bold
    assert back["B38"].font.sz == front["B38"].font.sz
    assert front["G1"].alignment.horizontal == "right"
    assert front["B38"].alignment.horizontal == "right"
    assert back["B1"].alignment.horizontal == "left"
    assert back["B38"].alignment.horizontal == "left"
    assert _rgb(back["B1"].font.color) == "808080"
    assert _rgb(back["G1"].font.color) == "808080"
    assert _rgb(back["B3"].font.color) == "808080"
    assert _rgb(back["B3"].border.bottom.color) == "B7B7B7"


def test_full_rn_and_psa_sections_keep_one_blank_writable_row():
    records = [
        record(f"Rn {index:03d}", "3178", start=time(19, 0), end=time(7, 30))
        for index in range(15)
    ]
    records.extend(
        record(f"Psa {index:03d}", "5056", role_group=ROLE_PSA, start=time(19, 0), end=time(7, 30))
        for index in range(11)
    )

    output = build_staffing_workbook(records)
    workbook = load_workbook(BytesIO(output))
    sheet = workbook["Tue Apr28 NIGHT"]

    assert sheet["C18"].value is None
    assert sheet["C18"].border.bottom.style == "thin"
    assert sheet["C19"].value == "PSA"
    assert sheet["C31"].value is None
    assert sheet["C31"].border.bottom.style == "thin"
    assert sheet["B32"].value.startswith("Patient Fall Review (Tue 4/28) - Charge RN:")
    assert str(sheet.print_area) == "'Tue Apr28 NIGHT'!$B$1:$H$40"


def _rgb(color) -> str | None:
    if color is None or color.type != "rgb" or color.rgb is None:
        return None
    return color.rgb[-6:]
