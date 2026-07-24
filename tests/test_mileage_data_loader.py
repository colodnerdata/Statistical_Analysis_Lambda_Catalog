"""Tests for the Mileage Data loader (lambda_catalog/write_sheet_mileage_data.py)."""
from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from lambda_catalog.write_sheet_mileage_data import (
    DEFAULT_XLSX_PATH,
    SHEET_NAME,
    SOURCE_TABLE_NAME,
    TABLE_NAME,
    _cell_row_col,
    load_mileage_rows,
)

_EXPECTED_HEADERS = (
    "MPG",
    "Cylinders",
    "Displacement",
    "Horsepower",
    "Weight",
    "Acceleration",
    "Model Year",
    "Origin",
    "Car Name",
    "Make",
    "Model?",
)


def test_names_follow_the_life_expectancy_naming_convention() -> None:
    assert SHEET_NAME == "Mileage Data"
    assert TABLE_NAME == "MileageData"
    assert SOURCE_TABLE_NAME == "Auto_MPG_Data"


def test_load_mileage_rows_returns_expected_shape() -> None:
    headers, rows = load_mileage_rows(DEFAULT_XLSX_PATH)

    assert tuple(headers) == _EXPECTED_HEADERS
    assert len(rows) == 406
    assert all(len(row) == len(headers) for row in rows)


def test_load_mileage_rows_parses_known_complete_row() -> None:
    headers, rows = load_mileage_rows(DEFAULT_XLSX_PATH)
    first_row = dict(zip(headers, rows[0]))

    assert first_row["MPG"] == 18
    assert first_row["Cylinders"] == 8
    assert first_row["Horsepower"] == 130
    assert first_row["Car Name"] == "chevrolet chevelle malibu"
    assert first_row["Make"] == "chevrolet"
    assert first_row["Model?"] == "chevelle malibu"


def test_load_mileage_rows_treats_na_marker_as_none() -> None:
    headers, rows = load_mileage_rows(DEFAULT_XLSX_PATH)
    mpg_idx = headers.index("MPG")
    horsepower_idx = headers.index("Horsepower")

    mpg_missing_rows = [row for row in rows if row[mpg_idx] is None]
    horsepower_missing_rows = [row for row in rows if row[horsepower_idx] is None]

    assert len(mpg_missing_rows) == 8
    assert len(horsepower_missing_rows) == 6
    # The two missing-value sets are disjoint: no row is missing both.
    assert not any(row[mpg_idx] is None for row in horsepower_missing_rows)


def _build_minimal_fixture(tmp_path: Path) -> Path:
    """Hand-roll a minimal xlsx with a 2-column, 2-row Auto_MPG_Data table."""
    xlsx_path = tmp_path / "mini_mpg.xlsx"
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        "</Relationships>"
    )
    sheet_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'
        "</Relationships>"
    )
    shared_strings = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">'
        "<si><t>MPG</t></si><si><t>Horsepower</t></si>"
        "</sst>"
    )
    # Row 2: MPG=18, Horsepower="NA" (text, missing marker).
    # Row 3: MPG blank, Horsepower=130.
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData>"
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
        '<row r="2"><c r="A2"><v>18</v></c><c r="B2" t="str"><v>NA</v></c></row>'
        '<row r="3"><c r="B3"><v>130</v></c></row>'
        "</sheetData>"
        "</worksheet>"
    )
    table_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'id="1" name="Auto_MPG_Data" displayName="Auto_MPG_Data" ref="A1:B3" totalsRowShown="0">'
        '<autoFilter ref="A1:B3"/>'
        '<tableColumns count="2">'
        '<tableColumn id="1" name="MPG"/>'
        '<tableColumn id="2" name="Horsepower"/>'
        "</tableColumns>"
        "</table>"
    )

    with zipfile.ZipFile(xlsx_path, "w") as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", sheet_rels)
        archive.writestr("xl/tables/table1.xml", table_xml)
        archive.writestr("xl/sharedStrings.xml", shared_strings)

    return xlsx_path


def test_load_mileage_rows_on_synthetic_fixture_exercises_blank_na_and_numeric(
    tmp_path: Path,
) -> None:
    xlsx_path = _build_minimal_fixture(tmp_path)

    headers, rows = load_mileage_rows(xlsx_path)

    assert headers == ["MPG", "Horsepower"]
    assert rows == [[18, None], [None, 130]]


def test_cell_row_col_parses_valid_reference() -> None:
    assert _cell_row_col("K407") == (407, 10)


def test_cell_row_col_rejects_trailing_garbage() -> None:
    with pytest.raises(ValueError):
        _cell_row_col("A1junk")


def test_load_mileage_rows_raises_when_table_ref_attribute_missing(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "no_ref.xlsx"
    with zipfile.ZipFile(xlsx_path, "w") as archive:
        archive.writestr(
            "xl/tables/table1.xml",
            '<?xml version="1.0"?>'
            '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'name="{SOURCE_TABLE_NAME}"><tableColumns count="1">'
            '<tableColumn id="1" name="X"/></tableColumns></table>',
        )

    with pytest.raises(ValueError, match="ref attribute"):
        load_mileage_rows(xlsx_path)


def test_load_mileage_rows_raises_when_table_column_unnamed(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "unnamed_column.xlsx"
    with zipfile.ZipFile(xlsx_path, "w") as archive:
        archive.writestr(
            "xl/tables/table1.xml",
            '<?xml version="1.0"?>'
            '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'name="{SOURCE_TABLE_NAME}" ref="A1:A1"><tableColumns count="1">'
            '<tableColumn id="1"/></tableColumns></table>',
        )

    with pytest.raises(ValueError, match="unnamed tableColumn"):
        load_mileage_rows(xlsx_path)


def test_load_mileage_rows_raises_on_missing_source_table(tmp_path: Path) -> None:
    xlsx_path = tmp_path / "no_table.xlsx"
    with zipfile.ZipFile(xlsx_path, "w") as archive:
        archive.writestr(
            "xl/tables/table1.xml",
            '<?xml version="1.0"?>'
            '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'name="Some_Other_Table" ref="A1:A1"><tableColumns count="1">'
            '<tableColumn id="1" name="X"/></tableColumns></table>',
        )

    with pytest.raises(ValueError):
        load_mileage_rows(xlsx_path)
