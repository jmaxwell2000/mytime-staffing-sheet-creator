from datetime import datetime

from fastapi.testclient import TestClient
from openpyxl import load_workbook

from app import __version__
from app.main import app
from tests.fixtures.source_workbooks import build_source_workbook


client = TestClient(app)


def test_healthz():
    response = client.get("/healthz")
    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_index_has_native_post_fallback_and_visible_offset_control():
    response = client.get("/")
    assert response.status_code == 200
    html = response.text
    assert app.title == "MyTime Staffing Sheet Creator"
    assert __version__ == "2.0"
    assert "MyTime Staffing Sheet Creator" in html
    assert "v2.0" in html
    assert "Updated June 3, 2026" in html
    assert "MyTime 3.0" not in html
    assert 'method="post"' in html
    assert 'action="/convert"' in html
    assert 'id="offset-control" class="offset-control"' in html
    assert 'id="offset-control" class="offset-control" hidden' not in html
    assert "viewed through paper" in html


def test_static_js_does_not_replace_native_download_with_blob_url():
    response = client.get("/static/app.js?v=4")
    assert response.status_code == 200
    script = response.text
    assert "URL.createObjectURL" not in script
    assert "event.preventDefault" not in script
    assert 'fetch("/convert"' not in script


def test_convert_returns_xlsx_download(tmp_path):
    data = build_source_workbook(
        [
            {
                "schedule": [
                    {
                        "job": "8536-RN Clinical Leader",
                        "employee": "Employee 001",
                        "start": datetime(2026, 4, 28, 19, 0),
                        "end": datetime(2026, 4, 29, 7, 30),
                    }
                ]
            }
        ]
    )

    response = client.post(
        "/convert",
        files={"workbook": ("source.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"include_back_pages": "false", "back_page_offset": "0.00"},
    )

    assert response.status_code == 200
    assert response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert 'filename="staffing-sheets-2026-04-28.xlsx"' in response.headers["content-disposition"]

    output = tmp_path / "output.xlsx"
    output.write_bytes(response.content)
    workbook = load_workbook(output)
    assert workbook.sheetnames == ["Tue Apr28 NIGHT"]
    assert workbook["Tue Apr28 NIGHT"]["F3"].value == "LEAD"


def test_convert_rejects_non_xlsx():
    response = client.post(
        "/convert",
        files={"workbook": ("source.txt", b"not a workbook", "text/plain")},
        data={"include_back_pages": "false", "back_page_offset": "0.00"},
    )

    assert response.status_code == 400
    assert response.json()["detail"] == "Upload a valid .xlsx workbook."


def test_convert_rejects_bad_offset():
    data = build_source_workbook(
        [
            {
                "schedule": [
                    {
                        "job": "3178-Staff RN Placeholder",
                        "employee": "Employee 001",
                        "start": datetime(2026, 4, 28, 7, 0),
                        "end": datetime(2026, 4, 28, 19, 30),
                    }
                ]
            }
        ]
    )

    response = client.post(
        "/convert",
        files={"workbook": ("source.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"include_back_pages": "true", "back_page_offset": "0.50"},
    )

    assert response.status_code == 400
    assert "Back page offset" in response.json()["detail"]
