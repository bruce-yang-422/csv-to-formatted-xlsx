"""CSV 解碼與結構驗證的回歸測試。

測試資料刻意涵蓋台灣常見編碼、BOM、引號內逗號／換行、異常欄名與列寬錯誤，
確保任何資料型別推測前，原始文字已被完整且可追蹤地讀入。
"""

from __future__ import annotations

from pathlib import Path

import pytest

from csv_to_formatted_xlsx.csv_reader import CsvReadError, read_csv


def test_reads_utf8_bom_and_preserves_quoted_multiline_value(tmp_path: Path) -> None:
    """UTF-8 BOM 應移除，CSV 引號內的逗號與換行則必須保留。"""

    source = tmp_path / "report.csv"
    source.write_bytes('條碼,品名\r\n02131137,"第一行,含逗號\n第二行"\r\n'.encode("utf-8-sig"))

    data = read_csv(source)

    assert data.encoding == "utf-8-sig"
    assert data.headers == ["條碼", "品名"]
    assert data.rows == [["02131137", "第一行,含逗號\n第二行"]]


def test_falls_back_to_cp950(tmp_path: Path) -> None:
    """UTF-8 解碼失敗時，常見繁體中文 CP950 檔案仍應可讀取。"""

    source = tmp_path / "report.csv"
    source.write_bytes("條碼,品名\r\n02131137,測試\r\n".encode("cp950"))

    data = read_csv(source)

    assert data.encoding == "cp950"
    assert data.rows[0][0] == "02131137"


def test_duplicate_and_blank_headers_produce_warnings(tmp_path: Path) -> None:
    """空白或重複欄名不應靜默忽略，而要留下可診斷的警告。"""

    source = tmp_path / "report.csv"
    source.write_text("條碼,,% Δ,% Δ\n1,,10%,20%\n", encoding="utf-8")

    data = read_csv(source)

    assert any("空白欄名" in warning for warning in data.warnings)
    assert any("重複欄名" in warning for warning in data.warnings)


def test_mismatched_row_width_reports_line_number(tmp_path: Path) -> None:
    """資料列欄數不符時應中止該檔，並精確指出 CSV 實體列號。"""

    source = tmp_path / "broken.csv"
    source.write_text("a,b\n1,2\n3\n", encoding="utf-8")

    with pytest.raises(CsvReadError, match="第 3 列欄位數不一致"):
        read_csv(source)
