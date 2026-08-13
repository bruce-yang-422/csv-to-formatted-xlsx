"""欄位分類與單格轉換規則的回歸測試。

這組測試保護兩項核心目標：識別碼永不被推測成數字，以及日期／數值輸出能被
Excel 樞紐分析辨識。新增關鍵字或日期格式時，應優先在此加入案例。
"""

from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal

from csv_to_formatted_xlsx.column_rules import build_rules, classify_header, convert_value
from csv_to_formatted_xlsx.models import ColumnKind


def test_identifier_headers_are_protected_text() -> None:
    """完全相符的中英文識別碼欄名必須分類為受保護文字。"""

    for header in ("條碼", " 貨號 ", "SKU", "訂單編號"):
        assert classify_header(header) is ColumnKind.PROTECTED_TEXT


def test_barcode_related_keywords_are_protected_text() -> None:
    """欄名只要包含條碼相關關鍵字，也必須整欄強制為文字。"""

    headers = (
        "商品條碼（主要）",
        "备用条码",
        "Barcode Number",
        "EAN-13",
        "UPC Code",
        "product_code",
        "Item-Code",
    )
    for header in headers:
        assert classify_header(header) is ColumnKind.PROTECTED_TEXT


def test_order_temporal_headers_are_pivot_friendly_types() -> None:
    """日期時間、年月與年/月必須選擇適合樞紐分析的語意型別。"""

    expected = {
        "訂單時間": ColumnKind.DATETIME,
        "訂單日期": ColumnKind.DATE,
        "訂單 年月": ColumnKind.YEAR_MONTH,
        "訂單年": ColumnKind.INTEGER,
        "月": ColumnKind.INTEGER,
        "Order Date": ColumnKind.DATE,
        "Order_Year": ColumnKind.INTEGER,
        "month": ColumnKind.INTEGER,
    }
    for header, kind in expected.items():
        assert classify_header(header) is kind


def test_custom_protected_alias_is_normalized() -> None:
    """使用者自訂別名應忽略前後空白、連續空白與英文大小寫。"""

    assert classify_header(" 客戶 ID ", {"客戶   id"}) is ColumnKind.PROTECTED_TEXT


def test_duplicate_percent_headers_get_index_based_rules() -> None:
    """重複欄名必須各自保有規則，不能因名稱相同而覆蓋。"""

    rules = build_rules(["% Δ", "% Δ"])
    assert [rule.index for rule in rules] == [0, 1]
    assert all(rule.kind is ColumnKind.PERCENT for rule in rules)


def test_numeric_conversion_is_strict_and_conservative() -> None:
    """合法數值才轉型；混合業務文字必須原樣保留並提出警告。"""

    assert convert_value("1,234", ColumnKind.INTEGER) == (1234, None)
    assert convert_value("12.5%", ColumnKind.PERCENT) == (Decimal("0.125"), None)
    assert convert_value("0.125", ColumnKind.PERCENT) == (Decimal("0.125"), None)

    value, warning = convert_value("12個", ColumnKind.INTEGER)
    assert value == "12個"
    assert warning is not None


def test_temporal_conversion_creates_real_excel_compatible_values() -> None:
    """無歧義日期應轉成 Excel 可辨識型別，模糊日期則不得猜測。"""

    assert convert_value("2026-08-13", ColumnKind.DATE) == (date(2026, 8, 13), None)
    assert convert_value("2026/08/13 09:30", ColumnKind.DATETIME) == (
        datetime(2026, 8, 13, 9, 30),
        None,
    )
    assert convert_value("2026-08", ColumnKind.YEAR_MONTH) == (date(2026, 8, 1), None)

    value, warning = convert_value("08/13/2026", ColumnKind.DATE)
    assert value == "08/13/2026"
    assert warning is not None


def test_text_and_identifiers_are_never_inferred() -> None:
    """超長識別碼和公式外觀文字不得在規則層被改值。"""

    long_identifier = "001234567890123456789"
    assert convert_value(long_identifier, ColumnKind.PROTECTED_TEXT) == (long_identifier, None)
    assert convert_value("=HYPERLINK(\"https://example.test\")", ColumnKind.TEXT)[0].startswith("=")
