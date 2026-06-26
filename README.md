# 大統進貨單 OCR 與建檔流程工作包

這個 repo 是給同事安裝到 Codex 使用的完整工作包，包含：

- 進貨流程 skills
- 大統工作助手專案規則
- 商品與廠商參考資料
- OCR 設定與本機 PaddleOCR 安裝腳本
- 進貨單 OCR 批次工具原始碼
- 完整 OCR + 產品比對 + 建檔/採購工具原始碼

## 安裝方式

在 Windows PowerShell 執行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\install.ps1"
```

預設安裝到：

```text
%USERPROFILE%\.codex\skills
%USERPROFILE%\Documents\大統工作助手
```

安裝完成後請重開 Codex。

## Codex 內使用

上傳或提供進貨單圖片時，Codex 會依 `AGENTS.md` 與 `convert-vendor-invoice-image` 啟動流程：

1. OCR 圖片轉試算表
2. 產品資料比對
3. 產品名稱、大類、產品代號覆核
4. 建檔用與採購匯入檔輸出

## 只做圖片轉 Excel

若只想批次把圖片轉成 Excel，可在安裝後執行：

```powershell
& "$env:USERPROFILE\Documents\大統工作助手\.venv-paddleocr\Scripts\python.exe" -X utf8 "$env:USERPROFILE\Documents\大統工作助手\tools\invoice_image_to_spreadsheet_app\invoice_image_to_spreadsheet_gui.py"
```

## 更新產品資料

把新的 `產品資料輸出.CSV` 放到：

```text
%USERPROFILE%\Documents\大統工作助手\參考資料
```

產品比對流程會檢查 CSV 是否為今天建立或修改的版本。

## 本機 OCR 引擎

`install.ps1` 會建立：

```text
%USERPROFILE%\Documents\大統工作助手\.venv-paddleocr
```

並安裝 `engine\requirements-ocr.txt` 內的 PaddleOCR 相關套件。

如果 repo release 內另外提供 `official_models.zip`，可放到 `engine\official_models.zip` 後重跑 `install.ps1`，腳本會解壓到 `%USERPROFILE%\.paddlex\official_models`。
