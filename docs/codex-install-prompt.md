# 給同事貼到 Codex 的安裝指令

同事只需要在 Codex 貼下面這段：

```text
https://github.com/ontaiko/datong-vendor-invoice-workflow-kit

幫我安裝這個
```

Codex 讀到後要做的事情：

1. 使用 GitHub 讀取 `ontaiko/datong-vendor-invoice-workflow-kit`。
2. 把 repo 下載到本機。
3. 執行 repo 根目錄的 `install.ps1`。
4. 等待 Python 套件安裝與中文 OCR 模型初始化完成，不要因數分鐘沒有輸出而重複啟動。
5. 確認安裝程式最後的 5 技能、本機引擎、專案參考資料與 Excel COM 檢查結果。
6. 確認 `%USERPROFILE%\Documents\大統工作助手\PROJECT_MEMORY.md` 已建立並可讀取。
7. 安裝完成後提醒使用者重開 Codex。

安裝要求：

- 複製 repo 內的 `skills` 到 `%USERPROFILE%\.codex\skills`
- 複製 repo 內的 `project` 到 `%USERPROFILE%\Documents\大統工作助手`
- 安裝本機 PaddleOCR 引擎到 `%USERPROFILE%\Documents\大統工作助手\.venv-paddleocr`
- 同時安裝 OpenCV、openpyxl、RapidFuzz、Pillow 與 NumPy
- 設定使用者環境變數 `DATONG_WORKSPACE`
- 若有同名舊技能或專案檔案，保留安裝前備份
- 安裝 `memory-seed/datong-project-memory.md` 為專案根目錄的 `PROJECT_MEMORY.md`
- 安裝完成後提醒我重開 Codex

安裝完成後，請先幫我跑以下檢查：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\scripts\verify-install.ps1" `
  -ProjectRoot "$env:USERPROFILE\Documents\大統工作助手" `
  -CodexHome "$env:USERPROFILE\.codex"
```

然後告訴我是否可以開始使用 `convert-vendor-invoice-image`。
