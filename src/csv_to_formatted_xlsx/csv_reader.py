"""CSV 解碼、方言偵測與資料列結構驗證。

本模組只回傳原始字串，不進行任何數值或日期推測，避免條碼在進入欄位規則
前就遺失前導零或長數字精度。解碼必須完整成功，不可使用 errors="ignore"。
"""

from __future__ import annotations

import csv
import io
from pathlib import Path

from .models import CsvData


class CsvReadError(ValueError):
    """Raised when a CSV cannot be decoded or safely parsed."""


def _decode_csv(raw: bytes) -> tuple[str, str]:
    """依安全優先順序完整解碼位元組，並回傳實際使用的編碼名稱。"""

    # 有 BOM 時先用 utf-8-sig 自動移除 BOM；台灣舊系統常見 CP950 作最後備援。
    candidates = (
        ("utf-8-sig", "utf-8", "cp950")
        if raw.startswith(b"\xef\xbb\xbf")
        else ("utf-8", "cp950")
    )
    errors: list[str] = []
    for encoding in candidates:
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError as exc:
            errors.append(f"{encoding}: byte {exc.start}")
    raise CsvReadError("不支援的 CSV 編碼（嘗試結果：" + "; ".join(errors) + "）")


def _dialect_for(text: str) -> csv.Dialect:
    """偵測逗號、分號或 Tab 分隔；無法判斷時回退到標準逗號 CSV。"""

    # 限制樣本大小，避免大型報表只為判斷分隔符就重複掃描完整內容。
    sample = text[:65536]
    try:
        return csv.Sniffer().sniff(sample, delimiters=",;\t")
    except csv.Error:
        return csv.excel


def read_csv(path: Path) -> CsvData:
    """不推測型別地讀取 CSV，並驗證每筆資料的欄位數。"""

    try:
        raw = path.read_bytes()
    except OSError as exc:
        raise CsvReadError(f"無法讀取檔案：{exc}") from exc

    text, encoding = _decode_csv(raw)
    reader = csv.reader(io.StringIO(text, newline=""), dialect=_dialect_for(text))
    try:
        raw_headers = next(reader)
    except StopIteration as exc:
        raise CsvReadError("CSV 是空檔案，缺少標題列") from exc
    except csv.Error as exc:
        raise CsvReadError(f"無法解析標題列：{exc}") from exc

    if not raw_headers or all(not header.strip() for header in raw_headers):
        raise CsvReadError("CSV 缺少有效的標題列")

    # 只清理欄名前後空白與意外殘留 BOM；資料值本身完全不 strip。
    headers = [header.removeprefix("\ufeff").strip() for header in raw_headers]
    warnings: list[str] = []
    blank_columns = [str(index + 1) for index, header in enumerate(headers) if not header]
    if blank_columns:
        warnings.append("空白欄名位於第 " + "、".join(blank_columns) + " 欄")

    normalized = [" ".join(header.split()).casefold() for header in headers]
    duplicate_names = sorted({name for name in normalized if name and normalized.count(name) > 1})
    if duplicate_names:
        warnings.append("重複欄名：" + "、".join(duplicate_names))

    rows: list[list[str]] = []
    expected = len(headers)
    try:
        for row in reader:
            # csv.reader 對純空白實體列回傳 []，可安全略過；含分隔符的 ",," 仍是資料列。
            if not row:
                continue
            if len(row) != expected:
                raise CsvReadError(
                    f"第 {reader.line_num} 列欄位數不一致：預期 {expected} 欄，實際 {len(row)} 欄"
                )
            rows.append(row)
    except csv.Error as exc:
        raise CsvReadError(f"第 {reader.line_num} 列 CSV 格式錯誤：{exc}") from exc

    return CsvData(headers=headers, rows=rows, encoding=encoding, warnings=warnings)
