"""批次轉換流程的協調層。

本模組只負責尋找輸入檔、安排輸出名稱、逐檔呼叫讀取與寫入模組，以及彙整
成功／失敗結果。CSV 解析規則與 Excel 格式不可放在這裡，否則日後修改其中
一層時容易連帶破壞其他流程。

維護重點：每個 CSV 必須獨立處理；單檔失敗不得中止後續檔案。
"""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path

from .csv_reader import read_csv
from .models import ConversionResult
from .xlsx_writer import write_xlsx


def discover_csv_files(input_dir: Path) -> list[Path]:
    """找出輸入資料夾第一層的 CSV，排除常見暫存檔並依檔名排序。"""

    return sorted(
        (
            path
            for path in input_dir.iterdir()
            if path.is_file()
            and path.suffix.casefold() == ".csv"
            and not path.name.startswith(("~$", "."))
        ),
        key=lambda path: path.name.casefold(),
    )


def run(
    input_dir: Path,
    output_dir: Path,
    *,
    overwrite: bool = True,
    protected_aliases: set[str] | None = None,
    logger: logging.Logger | None = None,
    generated_at: datetime | None = None,
) -> tuple[int, list[ConversionResult]]:
    """逐一轉換 CSV，回傳程序結束碼與各檔結果。

    ``generated_at`` 主要供測試固定時間；正式執行時省略即可使用本機現在時間。
    """

    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_files = discover_csv_files(input_dir)
    print(f"找到 {len(csv_files)} 個 CSV")
    if not csv_files:
        print(f"請將 CSV 放入：{input_dir}")
        return 0, []

    # 同一批次共用一個時間戳，方便使用者判斷哪些 XLSX 是同次操作產生的。
    batch_timestamp = (generated_at or datetime.now()).strftime("%Y%m%d_%H%M%S")
    results: list[ConversionResult] = []
    for source in csv_files:
        output = output_dir / f"{source.stem}_{batch_timestamp}.xlsx"
        try:
            data = read_csv(source)
            warnings = write_xlsx(
                data,
                output,
                overwrite=overwrite,
                protected_aliases=protected_aliases,
            )
            result = ConversionResult(
                source=source,
                output=output,
                success=True,
                row_count=len(data.rows),
                encoding=data.encoding,
                warnings=warnings,
            )
            print(f"[成功] {source.name} → {output.name}（{len(data.rows):,} 列）")
            if logger:
                logger.info(
                    "成功 file=%s output=%s encoding=%s rows=%d warnings=%d",
                    source.name,
                    output.name,
                    data.encoding,
                    len(data.rows),
                    len(warnings),
                )
                for warning in warnings:
                    logger.warning("file=%s %s", source.name, warning)
        # 這裡刻意捕捉單檔的所有例外，確保損壞的 CSV 不會拖垮整個批次。
        except Exception as exc:
            result = ConversionResult(
                source=source,
                output=None,
                success=False,
                error=str(exc),
            )
            print(f"[失敗] {source.name}：{exc}")
            if logger:
                logger.exception("失敗 file=%s error=%s", source.name, exc)
        results.append(result)

    success_count = sum(result.success for result in results)
    failed_count = len(results) - success_count
    print(f"完成：成功 {success_count}、失敗 {failed_count}")
    # 結束碼 1 表示至少一檔失敗；全部成功或沒有輸入檔則由前面回傳 0。
    return (1 if failed_count else 0), results
