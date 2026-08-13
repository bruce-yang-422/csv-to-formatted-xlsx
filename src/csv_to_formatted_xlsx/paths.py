"""原始碼模式與 PyInstaller EXE 模式的執行路徑解析。

重要：封裝後的 in/out/logs 必須位於 EXE 旁，不可使用 PyInstaller 的臨時解壓
目錄（_MEIPASS），否則使用者會找不到輸入與輸出檔案。
"""

from __future__ import annotations

import sys
from pathlib import Path


def default_base_dir() -> Path:
    """回傳專案根目錄，或封裝後 EXE 所在目錄。"""

    # sys.frozen 是 PyInstaller 執行檔的標記；sys.executable 此時就是 EXE 路徑。
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def ensure_runtime_dirs(base_dir: Path) -> tuple[Path, Path, Path]:
    """確保預設輸入、輸出與日誌資料夾存在，並依序回傳三者。"""

    input_dir = base_dir / "in"
    output_dir = base_dir / "out"
    log_dir = base_dir / "logs"
    for directory in (input_dir, output_dir, log_dir):
        directory.mkdir(parents=True, exist_ok=True)
    return input_dir, output_dir, log_dir
