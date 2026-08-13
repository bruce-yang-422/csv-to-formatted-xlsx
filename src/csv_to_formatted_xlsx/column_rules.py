"""欄位分類與安全型別轉換規則。

資料流為「原始欄名 → ColumnKind → 儲存格值」。條碼等識別碼的保真優先級
最高；日期與數值則以 Excel 樞紐分析可辨識為目標。任何不能完全確認的值都
保留原始文字並產生警告，禁止用猜測方式改動資料。

新增欄位規則時請同步更新 tests/test_column_rules.py 與 XLSX 回讀測試。
"""

from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from .models import ColumnKind, ColumnRule

# 數值正規表示式採完整比對，避免把「12個」之類的業務文字誤轉成數字。
_WHITESPACE_RE = re.compile(r"\s+")
_INTEGER_RE = re.compile(r"^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)$")
_DECIMAL_RE = re.compile(r"^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$")

# 識別碼規則必須先於日期／數值規則判斷，命中後整欄永遠以文字寫入。
PROTECTED_HEADERS = {
    "條碼",
    "条码",
    "貨號",
    "货号",
    "sku",
    "ean",
    "upc",
    "gtin",
    "barcode",
    "訂單編號",
    "order id",
    "order number",
}
PROTECTED_CHINESE_KEYWORDS = (
    "條碼",
    "条码",
    "貨號",
    "货号",
    "商品碼",
    "商品码",
    "識別碼",
    "识别码",
)
PROTECTED_ENGLISH_PATTERN = re.compile(
    r"(?:^|[^a-z0-9])"
    r"(?:barcode|bar\s*code|ean(?:\s*\d+)?|upc(?:\s*[ae])?|gtin(?:\s*\d+)?|sku|"
    r"product\s+code|item\s+code)"
    r"(?:$|[^a-z0-9])"
)
TEXT_HEADERS = {
    "平台",
    "品牌",
    "系列",
    "品名",
    "包裝單位",
}
DATE_HEADERS = {
    "訂單日期",
    "訂購日期",
    "下單日期",
    "date",
    "order date",
}
DATETIME_HEADERS = {
    "訂單時間",
    "訂購時間",
    "下單時間",
    "datetime",
    "timestamp",
    "order time",
    "order datetime",
    "order timestamp",
}
YEAR_MONTH_HEADERS = {
    "訂單年月",
    "訂購年月",
    "下單年月",
    "年月",
    "year month",
    "yearmonth",
    "order year month",
}
YEAR_MONTH_INTEGER_HEADERS = {
    "訂單年",
    "訂單月",
    "訂購年",
    "訂購月",
    "下單年",
    "下單月",
    "年",
    "月",
    "order year",
    "order month",
    "year",
    "month",
}
INTEGER_HEADERS = {
    *YEAR_MONTH_INTEGER_HEADERS,
    "數量",
    "quantity",
    "qty",
}
DECIMAL_HEADERS = {
    "銷售額",
    "金額",
    "包裝值",
    "sales",
    "amount",
}


def normalize_header(header: str) -> str:
    """正規化欄名供規則比對；絕對不可用此函式改動實際儲存格內容。"""

    return _WHITESPACE_RE.sub(" ", header.removeprefix("\ufeff").strip()).casefold()


def classify_header(header: str, protected_aliases: set[str] | None = None) -> ColumnKind:
    """依欄名選擇保守且適合樞紐分析的輸出型別。"""

    normalized = normalize_header(header)
    protected = PROTECTED_HEADERS | {
        normalize_header(alias) for alias in (protected_aliases or set())
    }
    english_searchable = re.sub(r"[_-]+", " ", normalized)
    # 中文使用包含關鍵字，英文使用單字邊界，避免字串片段造成錯誤命中。
    if (
        normalized in protected
        or any(keyword in normalized for keyword in PROTECTED_CHINESE_KEYWORDS)
        or PROTECTED_ENGLISH_PATTERN.search(english_searchable)
    ):
        return ColumnKind.PROTECTED_TEXT
    compact = normalized.replace(" ", "")
    is_chinese_order_column = compact.startswith(("訂單", "訂購", "下單"))
    english_words = english_searchable.split()
    is_english_order_column = "order" in english_words
    # 判斷順序由較具體到較一般：年月 → 日期時間 → 日期 → 數值／文字。
    if normalized in YEAR_MONTH_HEADERS or english_searchable in YEAR_MONTH_HEADERS or (
        is_chinese_order_column and compact.endswith("年月")
    ):
        return ColumnKind.YEAR_MONTH
    if normalized in DATETIME_HEADERS or english_searchable in DATETIME_HEADERS or (
        is_chinese_order_column and compact.endswith("時間")
    ) or (
        is_english_order_column
        and any(keyword in english_words for keyword in ("time", "datetime", "timestamp"))
    ):
        return ColumnKind.DATETIME
    if normalized in DATE_HEADERS or english_searchable in DATE_HEADERS or (
        is_chinese_order_column and compact.endswith("日期")
    ) or (
        is_english_order_column and "date" in english_words
    ):
        return ColumnKind.DATE
    if "%" in normalized or normalized in {"百分比", "percent", "percentage", "pct"}:
        return ColumnKind.PERCENT
    if normalized in INTEGER_HEADERS or english_searchable in INTEGER_HEADERS:
        return ColumnKind.INTEGER
    if normalized in DECIMAL_HEADERS:
        return ColumnKind.DECIMAL
    if normalized in TEXT_HEADERS:
        return ColumnKind.TEXT
    return ColumnKind.TEXT


def build_rules(
    headers: list[str], protected_aliases: set[str] | None = None
) -> list[ColumnRule]:
    """以欄位索引建立規則，確保重複欄名（如兩個 % Δ）仍能正確處理。"""

    return [
        ColumnRule(index=index, header=header, kind=classify_header(header, protected_aliases))
        for index, header in enumerate(headers)
    ]


def convert_value(value: str, kind: ColumnKind) -> tuple[object | None, str | None]:
    """安全轉換一格資料；失敗時回傳原字串及可寫入日誌的警告。"""

    if value == "":
        return None, None
    if kind in {ColumnKind.PROTECTED_TEXT, ColumnKind.TEXT}:
        return value, None

    stripped = value.strip()
    if kind is ColumnKind.INTEGER:
        if _INTEGER_RE.fullmatch(stripped):
            return int(stripped.replace(",", "")), None
        return value, "無法安全解析為整數，已保留文字"

    if kind is ColumnKind.DECIMAL:
        if _DECIMAL_RE.fullmatch(stripped):
            try:
                return Decimal(stripped.replace(",", "")), None
            except InvalidOperation:
                pass
        return value, "無法安全解析為數值，已保留文字"

    if kind is ColumnKind.PERCENT:
        # 含 % 的值要除以 100；不含 % 的 0.125 已是 Excel 所需的小數值。
        number = stripped
        divisor = Decimal(100) if number.endswith("%") else Decimal(1)
        if number.endswith("%"):
            number = number[:-1].strip()
        if _DECIMAL_RE.fullmatch(number):
            try:
                return Decimal(number.replace(",", "")) / divisor, None
            except InvalidOperation:
                pass
        return value, "無法安全解析為百分比，已保留文字"

    if kind is ColumnKind.DATE:
        # 僅接受無歧義的年-月-日格式，不猜測 08/09 是月/日或日/月。
        for date_format in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
            try:
                return datetime.strptime(stripped, date_format).date(), None
            except ValueError:
                continue
        return value, "無法安全解析為日期，已保留文字"

    if kind is ColumnKind.DATETIME:
        datetime_formats = (
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M",
            "%Y/%m/%d %H:%M:%S",
            "%Y/%m/%d %H:%M",
            "%Y-%m-%dT%H:%M:%S",
            "%Y-%m-%dT%H:%M",
            "%Y年%m月%d日 %H:%M:%S",
            "%Y年%m月%d日 %H:%M",
        )
        for datetime_format in datetime_formats:
            try:
                return datetime.strptime(stripped, datetime_format), None
            except ValueError:
                continue
        for time_format in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(stripped, time_format).time(), None
            except ValueError:
                continue
        return value, "無法安全解析為日期時間，已保留文字"

    if kind is ColumnKind.YEAR_MONTH:
        # Excel 沒有獨立的「年月」型別，以每月 1 日儲存才能供樞紐日期分組。
        for year_month_format in ("%Y-%m", "%Y/%m", "%Y.%m", "%Y年%m月", "%Y%m"):
            try:
                parsed = datetime.strptime(stripped, year_month_format)
                return date(parsed.year, parsed.month, 1), None
            except ValueError:
                continue
        return value, "無法安全解析為年月，已保留文字"

    return value, None
