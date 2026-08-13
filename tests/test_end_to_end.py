"""從 CSV 探索到 XLSX 回讀的端到端測試。

此處不重複檢查所有格式細節，而是驗證模組串接、批次錯誤隔離、時間戳檔名與
重要識別碼確實能走完整個流程而不失真。
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from csv_to_formatted_xlsx.app import run


def test_batch_continues_after_one_file_fails(tmp_path: Path) -> None:
    """損壞 CSV 不得阻止有效檔案輸出，批次應回傳部分失敗結束碼。"""

    input_dir = tmp_path / "in"
    output_dir = tmp_path / "out"
    input_dir.mkdir()
    (input_dir / "a_good.csv").write_text(
        "條碼,品名,數量\n02131137,測試,2\n5060412219549三包,商品,3\n",
        encoding="utf-8-sig",
    )
    (input_dir / "b_bad.csv").write_text("a,b\n1\n", encoding="utf-8")

    # 固定產生時間，讓輸出檔名測試不受測試執行當下的系統時間影響。
    exit_code, results = run(
        input_dir,
        output_dir,
        generated_at=datetime(2026, 8, 13, 15, 30, 45),
    )

    assert exit_code == 1
    assert [result.success for result in results] == [True, False]
    expected_output = output_dir / "a_good_20260813_153045.xlsx"
    assert expected_output.exists()
    assert not (output_dir / "b_bad_20260813_153045.xlsx").exists()

    # 回讀正式輸出，而非只相信函式回傳成功，確認識別碼實際寫入無誤。
    workbook = load_workbook(expected_output)
    worksheet = workbook.active
    assert worksheet["A2"].value == "02131137"
    assert worksheet["A3"].value == "5060412219549三包"
    workbook.close()


def test_empty_input_is_a_normal_result(tmp_path: Path) -> None:
    """沒有 CSV 是正常操作情況，不應被誤報為程式失敗。"""

    exit_code, results = run(tmp_path / "in", tmp_path / "out")
    assert exit_code == 0
    assert results == []
