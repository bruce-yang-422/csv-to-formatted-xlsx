"""各模組共用的資料模型。

集中定義模型可避免 CSV 讀取、欄位規則、XLSX 寫入與批次流程互相依賴實作
細節。若新增 ColumnKind，必須同步更新 convert_value 與 xlsx_writer 的格式。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path


class ColumnKind(Enum):
    """CSV 欄位在 XLSX 中應採用的語意型別。"""

    PROTECTED_TEXT = auto()  # 條碼／識別碼：不得做任何型別推測。
    TEXT = auto()
    INTEGER = auto()
    DECIMAL = auto()
    PERCENT = auto()
    DATE = auto()
    DATETIME = auto()
    YEAR_MONTH = auto()  # 以當月 1 日儲存、顯示 yyyy-mm，供樞紐分組。


@dataclass(frozen=True, slots=True)
class ColumnRule:
    """依欄位索引選定的規則；索引設計可支援重複欄名。"""

    index: int
    header: str
    kind: ColumnKind


@dataclass(slots=True)
class CsvData:
    """已完整解碼且通過欄數驗證的 CSV；rows 仍全部是原始字串。"""

    headers: list[str]
    rows: list[list[str]]
    encoding: str
    warnings: list[str] = field(default_factory=list)


@dataclass(slots=True)
class ConversionResult:
    """單一 CSV 的轉換結果，供 CLI 摘要、測試與未來 GUI 使用。"""

    source: Path
    output: Path | None
    success: bool
    row_count: int = 0
    encoding: str | None = None
    warnings: list[str] = field(default_factory=list)
    error: str | None = None
