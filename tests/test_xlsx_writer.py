import os
import tempfile
import pytest
import opensheet_core


@pytest.fixture
def tmp_xlsx(tmp_path):
    return str(tmp_path / "output.xlsx")


def test_basic_write_and_read(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("Data")
    writer.write_row(["Name", "Value"])
    writer.write_row(["Alice", 42])
    writer.close()

    sheets = opensheet_core.read_xlsx(tmp_xlsx)
    assert len(sheets) == 1
    assert sheets[0]["name"] == "Data"
    assert sheets[0]["rows"][0] == ["Name", "Value"]
    assert sheets[0]["rows"][1] == ["Alice", 42]


def test_multiple_sheets(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("Sheet1")
    writer.write_row(["a", "b"])
    writer.add_sheet("Sheet2")
    writer.write_row([1, 2, 3])
    writer.close()

    names = opensheet_core.sheet_names(tmp_xlsx)
    assert names == ["Sheet1", "Sheet2"]

    rows1 = opensheet_core.read_sheet(tmp_xlsx, sheet_name="Sheet1")
    assert rows1 == [["a", "b"]]

    rows2 = opensheet_core.read_sheet(tmp_xlsx, sheet_name="Sheet2")
    assert rows2 == [[1, 2, 3]]


def test_all_types(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("Types")
    writer.write_row(["text", 42, 3.14, True, False, None])
    writer.close()

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert rows[0][0] == "text"
    assert rows[0][1] == 42
    assert rows[0][2] == 3.14
    assert rows[0][3] is True
    assert rows[0][4] is False
    # None (empty) cells are not written, so they won't appear at the end


def test_context_manager(tmp_xlsx):
    with opensheet_core.XlsxWriter(tmp_xlsx) as w:
        w.add_sheet("Auto")
        w.write_row(["closed", "automatically"])

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert rows == [["closed", "automatically"]]


def test_write_after_close_raises(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("X")
    writer.write_row(["ok"])
    writer.close()

    with pytest.raises(RuntimeError, match="already closed"):
        writer.write_row(["fail"])


def test_write_without_sheet_raises(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    with pytest.raises(Exception):
        writer.write_row(["no sheet"])
    writer.close()


def test_special_characters(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("Special")
    writer.write_row(["a & b", "<tag>", 'quote "here"', "it's fine"])
    writer.close()

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert rows[0][0] == "a & b"
    assert rows[0][1] == "<tag>"
    assert rows[0][2] == 'quote "here"'
    assert rows[0][3] == "it's fine"


def test_empty_rows(tmp_xlsx):
    writer = opensheet_core.XlsxWriter(tmp_xlsx)
    writer.add_sheet("Gaps")
    writer.write_row(["row1"])
    writer.write_row([])
    writer.write_row(["row3"])
    writer.close()

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert len(rows) == 3
    assert rows[0] == ["row1"]
    assert rows[2] == ["row3"]


def test_large_write(tmp_xlsx):
    """Write 10k rows to verify streaming doesn't blow up."""
    with opensheet_core.XlsxWriter(tmp_xlsx) as w:
        w.add_sheet("Big")
        for i in range(10000):
            w.write_row([f"row_{i}", i, i * 0.1])

    rows = opensheet_core.read_sheet(tmp_xlsx)
    assert len(rows) == 10000
    assert rows[0] == ["row_0", 0, 0.0]
    assert rows[9999][0] == "row_9999"
