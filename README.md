# CSV to Formatted XLSX

將 Looker Studio／Data Studio 匯出的 CSV 批次轉換為已套用格式的 XLSX，重點是完整保留條碼、貨號、SKU 與訂單編號等識別碼。

> 目前狀態：規格設計階段，轉換程式尚待實作。完整需求請參閱 [開發白皮書](./csv-to-formatted-xlsx_%E9%96%8B%E7%99%BC%E7%99%BD%E7%9A%AE%E6%9B%B8.md)。

## 為什麼需要這個工具？

直接使用 Excel 開啟 CSV 時，Excel 可能自動推測欄位型別，造成：

- 長條碼顯示為科學記號。
- 以 `0` 開頭的條碼遺失前導零。
- 超過 15 位的數字發生不可逆的精度損失。
- 條碼中的中文註記或混合文字格式不一致。

本工具預計直接讀取 CSV 原始文字，並將受保護欄位寫成真正的 Excel 文字儲存格，而不是已失真的數字搭配顯示格式。

## 預計操作方式

1. 將 CSV 放入程式旁的 `in` 資料夾。
2. 雙擊 `CSV報表轉換工具.exe`。
3. 從 `out` 資料夾取得同名 XLSX。

## 預計功能

- 批次處理 `in` 資料夾內的 CSV。
- 支援 UTF-8 BOM、UTF-8，並以 CP950 作為常見台灣 CSV 的備援編碼。
- 保留條碼、前導零、長識別碼及中文註記。
- 防止 CSV 內容成為 Excel 公式。
- 套用標題樣式、凍結首列、自動篩選、欄寬及基本數值格式。
- 使用暫存檔與原子性取代，安全覆蓋既有同名 XLSX。
- 單一檔案失敗時繼續處理其他檔案，並留下執行日誌。
- 全程離線；不會上傳、移動、修改或刪除原始 CSV。

## 開發環境

規格以 Python 3.11 或 3.12 為目標版本，主要使用：

- Python 標準庫 `csv`
- `openpyxl`
- `pytest`
- `PyInstaller`

程式碼完成後，預計使用以下流程建立環境：

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements-dev.txt
```

執行測試：

```powershell
pytest
```

實際執行與封裝命令會在程式實作後補充。

## 預計錯誤處理

- **輸出檔被占用**：關閉正在開啟該 XLSX 的 Excel 視窗後重試。
- **編碼不支援**：來源必須能以 UTF-8 BOM、UTF-8 或 CP950 完整解碼，工具不會忽略無法解碼的字元。
- **CSV 結構異常**：若某列欄位數與標題不一致，該檔案會失敗並回報列號，不會產生看似完整的半成品。

## Windows SmartScreen

未簽章的自製 EXE 可能觸發 Windows Defender SmartScreen 警告。正式公司發布時建議使用可信任的程式碼簽章憑證；使用者也應只執行由可信管道取得的版本。

## 授權

本專案採用 [MIT License](./LICENSE)。
