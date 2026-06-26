# 大統進貨單 OCR 與建檔流程工作包

這個 repo 是給同事安裝到 Codex 使用的完整工作包。使用者不需要理解資料夾結構，只要把這個 GitHub 連結貼給 Codex，並說「幫我安裝這個」，Codex 就可以讀取 repo 並執行安裝。

- 進貨流程 skills
- 大統工作助手專案規則
- 商品與廠商參考資料
- OCR 設定與本機 OpenCV / PaddleOCR 安裝腳本
- RapidFuzz 本地商品相似度比對規則

本工作包完全依賴 Codex 執行流程，不包含額外桌面程式。

## 安裝方式

最簡單方式：在同事的 Codex 裡貼上：

```text
https://github.com/ontaiko/datong-vendor-invoice-workflow-kit

幫我安裝這個
```

Codex 應執行的工作：

1. 讀取這個 GitHub repo。
2. 下載或 clone repo 到本機。
3. 執行 repo 根目錄的 `install.ps1`。
4. 安裝完成後，提醒使用者重開 Codex。

手動安裝時，在 repo 根目錄執行：

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

1. OpenCV 圖片前處理後，用 PaddleOCR 圖片轉試算表
2. RapidFuzz 搭配產品資料 CSV 做本地產品資料比對
3. 產品名稱、大類、產品代號覆核
4. 建檔用與採購匯入檔輸出

原則是先用本地引擎處理穩定、可重複的判定，AI 只協助低信心項目、疑點整理與人工確認清單。

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

並安裝 `engine\requirements-ocr.txt` 內的 OpenCV、PaddleOCR、RapidFuzz 相關套件。

如果 repo release 內另外提供 `official_models.zip`，可放到 `engine\official_models.zip` 後重跑 `install.ps1`，腳本會解壓到 `%USERPROFILE%\.paddlex\official_models`。

## 驗證

更新或安裝後可執行本地規則測試：

```powershell
.\.venv-paddleocr\Scripts\python.exe -X utf8 ".\scripts\test-local-engine-rules.py"
```

測試會檢查：

- RapidFuzz 商品比對能分出 `已建檔`、`有類似產品`、`確認為新品`。
- 低可信雜訊候選不會進入人工確認清單。
- 已填產品代號但品名差異過大時，檢查表會加註解與淡黃色提醒。
