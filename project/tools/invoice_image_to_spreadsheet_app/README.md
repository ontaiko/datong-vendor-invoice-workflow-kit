# 進貨單 OCR 批次轉試算表

這是單純版進貨單 OCR 工具：一次選多張圖片，逐張輸出可覆核的 Excel 試算表。

## 使用方式

1. 雙擊 `進貨單OCR批次轉試算表.exe`。
2. 按 `選擇圖片`，可一次多選 `.jpg`、`.jpeg`、`.png`、`.webp`。
3. 按 `開始批次 OCR`。
4. 每張圖片會依序處理，Excel 會輸出在原圖片同一個資料夾。
5. 處理完成後，可選取清單列並按 `開啟選取 Excel` 或 `開啟圖片資料夾`。

## 輸出內容

每個 Excel 會包含：

- `進貨明細`：正式交接用明細表。
- `OCR測試紀錄`：OCR 原文、版本、信心分數與疑點。

本工具只做 OCR 轉試算表，不做產品比對、不產生建檔用或採購單匯入檔。

## 狀態說明

- `完成`：Excel 已產生，且 OCR 腳本未回報疑點。
- `需覆核`：Excel 已產生，但需要打開 `OCR測試紀錄` 檢查疑點。
- `失敗`：該圖片未成功產生 Excel；其他圖片仍會繼續處理。

## OCR 環境

工具沿用：

```text
C:\Users\user\.codex\skills\extract-vendor-invoice-image\scripts\local_paddleocr_invoice_to_xlsx.py
```

設定檔：

```text
C:\Users\user\Documents\大統工作助手\參考資料\OCR設定.json
```

暫存檔：

```text
C:\Users\user\Documents\大統工作助手\.codex-tmp\invoice-image-to-spreadsheet
```

## 重新打包

在 PowerShell 執行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\tools\invoice_image_to_spreadsheet_app\build_exe.ps1"
```

打包完成後會產生：

```text
tools\invoice_image_to_spreadsheet_app\進貨單OCR批次轉試算表.exe
```
