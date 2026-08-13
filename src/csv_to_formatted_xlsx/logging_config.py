"""轉換日誌設定。

日誌供雙擊 EXE 後追查問題，固定使用 UTF-8 寫入 logs/converter.log。不得記錄
完整報表內容，只記檔名、編碼、列數、警告摘要與例外堆疊。
"""

from __future__ import annotations

import logging
from pathlib import Path


def configure_logging(log_dir: Path) -> logging.Logger:
    """建立或重用指向指定目錄的檔案 logger。"""

    log_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("csv_to_formatted_xlsx")
    logger.setLevel(logging.INFO)
    # 禁止向 root logger 傳遞，避免 VS Code 或其他宿主重複輸出同一筆訊息。
    logger.propagate = False

    log_path = (log_dir / "converter.log").resolve()
    # 測試可能在同一 Python 程序重複呼叫；相同路徑不得重複加入 handler。
    for handler in logger.handlers:
        if isinstance(handler, logging.FileHandler) and Path(handler.baseFilename) == log_path:
            return logger

    # 若 log_dir 改變，先關閉舊檔案 handle，避免 Windows 鎖住日誌檔。
    for handler in list(logger.handlers):
        handler.close()
        logger.removeHandler(handler)

    handler = logging.FileHandler(log_path, encoding="utf-8")
    handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s %(name)s - %(message)s")
    )
    logger.addHandler(handler)
    return logger
