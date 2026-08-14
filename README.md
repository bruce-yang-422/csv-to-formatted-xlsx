# CSV to Formatted XLSX

[![Version](https://img.shields.io/badge/version-v0.1.0-0969da)](https://github.com/bruce-yang-422/csv-to-formatted-xlsx/releases)
[![Python](https://img.shields.io/badge/Python-3.11%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-0078D4?logo=windows&logoColor=white)](#支援格式)
[![License: MIT](https://img.shields.io/badge/license-MIT-yellow.svg)](./LICENSE)

將 Looker Studio／Data Studio 匯出的 CSV 批次轉換成適合 Microsoft Excel 與樞紐分析使用的 XLSX。

本工具會保護條碼、貨號、SKU 等識別碼，避免 Excel 自動轉成科學記號、刪除前導零，或破壞超過 15 位的長數字。同時會把日期、年月、年、月及數值轉成樞紐分析可辨識的型別。

## 一般使用者操作

將整個發布資料夾複製到任何位置，例如桌面、文件夾、網路磁碟或隨身碟。資料夾也可以自行重新命名，但請保留 EXE 與 `in`、`out`、`logs` 的相對位置：

```text
任意資料夾/
├─ CSV報表轉換工具.exe
├─ 使用說明.txt
├─ in/
├─ out/
└─ logs/
```

接著操作：

1. 將 CSV 放入 EXE 旁的 `in`。
2. 雙擊 `CSV報表轉換工具.exe`。
3. 從 EXE 旁的 `out` 取得轉換完成的 XLSX。

執行紀錄位於：

```text
logs\converter.log
```

程式沒有綁定 C 槽、D 槽或任何固定路徑；EXE 會以自己所在的資料夾為基準。原始 CSV 不會被上傳、修改、移動或刪除。本工具全程離線，也不要求電腦安裝 Microsoft Excel 或 Python。

## 輸出檔名

輸出會保留原始 CSV 檔名，並加入執行電腦的本地產生時間：

```text
原始檔名_YYYYMMDD_HHMMSS.xlsx
```

例如：

```text
銷售報表.csv
→ 銷售報表_20260813_153045.xlsx
```

同一批次處理的所有檔案會使用相同的時間尾綴。

## 欄位處理規則

### 條碼與識別碼

以下欄位會強制寫成真正的 Excel 文字儲存格：

- 中文關鍵字：條碼、条码、貨號、货号、商品碼、商品码、識別碼、识别码。
- 英文關鍵字：Barcode、Bar Code、EAN、UPC、GTIN、SKU、Product Code、Item Code。
- 明確識別碼欄名：訂單編號、Order ID、Order Number。

欄名包含關鍵字也會生效，例如 `商品條碼（主要）`、`Barcode Number`、`EAN-13`。

受保護欄位會保留：

- 前導零，例如 `02131137`。
- 18 位以上長識別碼。
- 中文註記，例如 `02131356短效`。
- 非數字內容，例如 `無資料`。

### 日期與樞紐分析

| CSV 欄位 | XLSX 儲存型別 | 顯示格式 |
|---|---|---|
| 訂單日期、Order Date | Excel 日期 | `yyyy-mm-dd` |
| 訂單時間、Order Time | Excel 日期時間 | `yyyy-mm-dd hh:mm:ss` |
| 訂單年月、年月 | Excel 日期（當月 1 日） | `yyyy-mm` |
| 訂單年、年 | 整數 | `#,##0` |
| 訂單月、月 | 整數 | `#,##0` |

年月以該月 1 日作為底層日期，例如 `2026-08` 儲存為 2026-08-01，但畫面只顯示 `2026-08`。這樣 Excel 樞紐分析才能辨識、排序及進行日期分組。

程式只解析無歧義格式，例如：

```text
2026-08-13
2026/08/13
2026-08-13 09:30
2026-08
202608
```

無法安全解析的日期會保留原文字並寫入警告，不會自行猜測月／日順序。

### 數值與百分比

- 數量等明確整數欄會寫成 Excel 數值。
- 銷售額、金額等欄位會寫成數值並套用千分位格式。
- `12.5%` 會轉成 Excel 百分比值 `0.125`。
- 不含 `%` 的 `0.125` 視為 12.5%。
- 無法安全解析的內容會保留原文字並記錄警告。

### Excel 格式與安全

- 標題列為深色底、白字、粗體。
- 凍結首列 `A2`。
- 自動篩選涵蓋完整資料。
- 自動估算欄寬。
- `=`、`+`、`-`、`@` 開頭的一般文字會以文字儲存，避免公式注入。
- 先寫入並驗證暫存 XLSX，成功後才發布正式檔案。
- 單一 CSV 失敗不會中止其他檔案。

## 支援格式

- 輸入：`.csv`，副檔名不分大小寫。
- 輸入編碼：UTF-8 BOM、UTF-8、CP950。
- 分隔符：逗號、分號或 Tab；無法判斷時使用逗號。
- 輸出：`.xlsx`。
- 目標平台：64 位元 Windows 10／11。

## VS Code／原始碼執行

原始碼模式會使用專案根目錄的 `in`、`out`、`logs`，不會使用 `release` 裡面的資料夾。

建立環境：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --editable ".[dev]"
```

執行程式：

```powershell
python run_converter.py
```

或使用套件入口：

```powershell
python -m csv_to_formatted_xlsx
```

開發或自動化時，可選擇自訂輸入與輸出資料夾。以下使用相對路徑作為範例，不限定磁碟位置：

```powershell
python -m csv_to_formatted_xlsx `
  --input ".\我的CSV" `
  --output ".\轉換結果" `
  --no-pause
```

`--input` 和 `--output` 都是選用參數，也可以指定桌面、網路磁碟或其他磁碟的完整路徑。一般主管使用發布版時不需要這些參數。

查看全部參數：

```powershell
python -m csv_to_formatted_xlsx --help
```

## 測試

```powershell
python -m pytest
python -m ruff check src tests run_converter.py
```

目前共有 17 項自動化測試，涵蓋：

- UTF-8 BOM、UTF-8、CP950。
- 引號內的逗號與換行。
- CSV 欄位數異常與列號回報。
- 條碼前導零及 18 位以上識別碼。
- 中文／英文條碼關鍵字。
- 重複百分比欄位。
- 日期、日期時間、年月及樞紐友善型別。
- Excel 公式注入防護。
- XLSX 回讀、凍結窗格與自動篩選。
- 批次錯誤隔離及時間戳輸出檔名。

真實 5,983 列報表已完成轉換驗收。

## 建置 Windows EXE

```powershell
.\scripts\build.bat
```

建置腳本會：

1. 建立或使用 `.venv`。
2. 安裝開發依賴。
3. 執行全部測試。
4. 使用 `converter.spec` 呼叫 PyInstaller。
5. 將發布成果複製至 `release`。

發布內容：

```text
release/
├─ CSV報表轉換工具.exe
├─ 使用說明.txt
├─ in/
├─ out/
└─ logs/
```

清除可重建的封裝中間產物：

```powershell
.\scripts\clean_build.bat
```

## 常見問題

### 找不到 CSV

確認 CSV 放在目前執行模式所使用的 `in`：

- 執行 `run_converter.py`：使用專案根目錄的 `in`。
- 執行發布版 EXE：使用 EXE 所在資料夾旁的 `in`，不論整個資料夾被放在哪裡。

### XLSX 無法輸出

若目標 XLSX 正在 Excel 中開啟，Windows 可能鎖定檔案。請關閉 Excel 中的檔案後重試；既有輸出不會被破壞。

### CSV 編碼不支援

程式不會忽略無法解碼的字元。請將 CSV 另存為 UTF-8 BOM、UTF-8 或 CP950。

### CSV 結構異常

若某列欄位數與標題列不同，該檔會失敗並顯示列號，不會產生看似完整的半成品。

### Windows SmartScreen 警告

未簽章的自製 EXE 可能觸發 Windows Defender SmartScreen。正式公司發布建議使用可信任的程式碼簽章憑證，並只從可信管道取得程式。

## 結束碼

| 結束碼 | 意義 |
|---|---|
| `0` | 全部成功，或輸入資料夾沒有 CSV |
| `1` | 至少一個 CSV 轉換失敗 |
| `2` | 發生不可恢復的程式錯誤 |

## 已知限制

- 目前使用一般 openpyxl Workbook；超大型資料尚未導入串流寫入。
- 欄位規則集中於程式碼，尚未提供外部 `config.toml`。
- 尚未提供拖放圖形介面。
- EXE 尚未進行程式碼簽章。

## 授權

本專案採用 [MIT License](./LICENSE)。
