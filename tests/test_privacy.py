from pathlib import Path


def test_no_real_workbook_or_pdf_fixtures_are_committed():
    root = Path(__file__).resolve().parents[1]
    blocked_suffixes = {".xlsx", ".xls", ".pdf"}
    blocked = [
        path
        for path in root.rglob("*")
        if path.is_file()
        and path.suffix.lower() in blocked_suffixes
        and ".git" not in path.parts
    ]

    assert blocked == []

