# 大統進貨單 OCR 與建檔流程工作包

這個 repo 是給同事安裝到 Codex 使用的完整工作包。使用者不需要理解資料夾結構，只要把這個 GitHub 連結貼給 Codex，並說「幫我安裝這個」，Codex 就可以讀取 repo 並執行安裝。

- 進貨流程 5 個 skills（總管＋4 個可獨立執行的子技能）
- 大統工作助手專案規則
- 商品與廠商參考資料
- OCR 設定與本機引擎自動安裝／驗證腳本

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
4. 確認 5 個技能、本機引擎與專案參考資料通過驗證。
5. 安裝完成後，提醒使用者重開 Codex。

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

若電腦已有同名技能或專案檔案，安裝程式會先備份到：

```text
%USERPROFILE%\.codex\install-backups\datong-vendor-invoice-workflow-kit
```

## 包含的技能

- `convert-vendor-invoice-image`
- `extract-vendor-invoice-image`
- `match-product-catalog`
- `review-invoice-product-check`
- `build-inventory-import-files`

## Codex 內使用

上傳或提供進貨單圖片時，Codex 會依 `AGENTS.md` 與 `convert-vendor-invoice-image` 啟動流程：

1. OCR 圖片轉試算表
2. 產品資料比對
3. 產品名稱、大類、產品代號覆核
4. 建檔用與採購匯入檔輸出

## 更新產品資料

把新的 `產品資料輸出.CSV` 放到：

```text
%USERPROFILE%\Documents\大統工作助手\參考資料
```

產品比對流程會檢查 CSV 是否為今天建立或修改的版本。

## 本機 OCR 引擎

`install.ps1` 會自動安裝 Python 3.12（缺少時透過 winget）、建立：

```text
%USERPROFILE%\Documents\大統工作助手\.venv-paddleocr
```

並安裝及驗證下列本機引擎：

- OpenCV：單據圖片前處理
- PaddleOCR／PaddlePaddle：本機 OCR
- openpyxl：中間 `.xlsx` 讀寫
- RapidFuzz：本機產品相似比對
- Pillow／NumPy：圖片與陣列處理

安裝時會預先初始化中文 OCR 模型，因此第一次安裝需要網路，時間也會比一般 skill 安裝久。

正式輸出舊版 `.xls` 還需要桌面版 Microsoft Excel。Excel 屬於授權軟體，不會由這個 repo 自動安裝；安裝驗證會檢查 Excel COM，未安裝時會清楚警告，但 OCR、比對與中間 `.xlsx` 仍可使用。

如果 repo release 內另外提供 `official_models.zip`，可放到 `engine\official_models.zip` 後重跑 `install.ps1`，腳本會解壓到 `%USERPROFILE%\.paddlex\official_models`。

## 安裝驗證

安裝程式最後會自動執行 `scripts\verify-install.ps1`，檢查：

- 5 個技能與必要腳本是否齊全
- 專案規則、OCR 設定、產品資料、廠商資料與 `.xls` 範本是否齊全
- PaddleOCR、OpenCV、openpyxl、RapidFuzz 等 Python 套件能否匯入
- Python 環境是否有相依衝突
- Excel COM 是否可用
