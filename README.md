# 大統進貨單 OCR 與建檔流程工作包

這個 repo 是給同事安裝到 Codex 使用的完整工作包。使用者不需要理解資料夾結構，只要把這個 GitHub 連結貼給 Codex，並說「幫我安裝這個」，Codex 就可以讀取 repo 並執行安裝。

- 進貨流程 5 個 skills（總管＋4 個可獨立執行的子技能）
- 大統工作助手專案規則
- 商品與廠商參考資料
- 整個大統工作助手的可攜式專案記憶
- OCR 設定、本機引擎自動安裝／驗證腳本，以及 5 組可離線使用的 OCR 模型
- 可選用的 Windows 桌面版「大統進貨助手」原始碼、測試與已驗證 EXE

原本的 Codex 安裝流程與 `install.ps1` 維持不變，不會自動安裝桌面 APP。需要不透過 AI 操作時，可另外使用 `desktop-app`。

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
5. 確認專案根目錄已建立 `PROJECT_MEMORY.md`。
6. 安裝完成後，提醒使用者重開 Codex。

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

## 專案記憶

`memory-seed/datong-project-memory.md` 保存整個大統工作助手的可攜式記憶，包括使用者偏好、進貨流程、預購檢查、Gmail 收單、封面套圖、PDF／Line 文案、停止節點與已知故障。

安裝時會複製到：

```text
%USERPROFILE%\Documents\大統工作助手\PROJECT_MEMORY.md
```

原始聊天、帳號、token、郵件識別碼與機器狀態不會放進公開 repo。這份記憶提供專案上下文，但不會讓舊聊天自動出現在另一台電腦的 Codex 側邊欄。

## Codex 內使用

詳細操作請先看：[進貨流程使用說明](docs/workflow-usage-guide.md)。

上傳或提供進貨單圖片時，Codex 會依 `AGENTS.md` 與 `convert-vendor-invoice-image` 啟動流程：

1. OCR 圖片轉試算表
2. 產品資料比對
3. 產品名稱、大類、產品代號覆核
4. 建檔用與採購匯入檔輸出

## Windows 桌面 APP

`desktop-app` 提供不依賴 ChatGPT、Codex 或雲端 AI API 的本機操作介面，使用 OpenCV、PaddleOCR、RapidFuzz、openpyxl 與 Excel COM 執行相同的 OCR、產品比對、人工覆核與正式輸出流程。

目前上傳內容包含：

- `大統進貨助手.exe`
- 完整 Python／PowerShell 原始碼
- APP 必要參考資料與舊快照 SHA256 防護
- 自動回歸測試
- `build_exe.ps1` 與 `package-manifest.json`

操作、環境需求與重新打包方式請看：[Windows APP 使用說明](desktop-app/README.md)。

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

倉庫已包含目前流程使用的 5 組 PaddleX 模型：

- `PP-OCRv6_medium_det`
- `PP-OCRv6_medium_rec`
- `PP-LCNet_x1_0_doc_ori`
- `PP-LCNet_x1_0_textline_ori`
- `UVDoc`

安裝程式會將模型複製到 `%USERPROFILE%\.paddlex\official_models`，依 `engine\model-manifest.json` 逐檔核對大小與 SHA256，再實際載入模型。模型不需要在第一次 OCR 時另外下載；Python 與套件安裝仍需要網路。

模型由 PaddlePaddle 發布並依 Apache License 2.0 提供，來源與授權資訊放在 `engine\official_models\NOTICE.md`。

正式輸出舊版 `.xls` 還需要桌面版 Microsoft Excel。Excel 屬於授權軟體，不會由這個 repo 自動安裝；安裝驗證會檢查 Excel COM，未安裝時會清楚警告，但 OCR、比對與中間 `.xlsx` 仍可使用。

## 安裝驗證

安裝程式最後會自動執行 `scripts\verify-install.ps1`，檢查：

- 5 個技能與必要腳本是否齊全
- 專案規則、OCR 設定、產品資料、廠商資料與 `.xls` 範本是否齊全
- `PROJECT_MEMORY.md` 是否已安裝
- PaddleOCR、OpenCV、openpyxl、RapidFuzz 等 Python 套件能否匯入
- 5 組離線模型的檔案大小、SHA256 與模型初始化是否通過
- Python 環境是否有相依衝突
- Excel COM 是否可用

## 自動判斷與優化檢查

安裝後，Agent 會依專案 `AGENTS.md` 與 `convert-vendor-invoice-image` 自動判斷上傳圖片是否為進貨單、銷貨憑單或送貨單。每次技能或流程執行，都會在停止節點與完成時檢查是否有可改善之處；即使沒有改善項目，也會明確回報檢查結果。
