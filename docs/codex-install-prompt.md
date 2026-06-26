# 給同事貼到 Codex 的安裝指令

請用 GitHub 讀取 `ontaiko/datong-vendor-invoice-workflow-kit`，把 repo 下載到本機，然後執行 repo 根目錄的 `install.ps1`。

安裝要求：

- 複製 repo 內的 `skills` 到 `%USERPROFILE%\.codex\skills`
- 複製 repo 內的 `project` 到 `%USERPROFILE%\Documents\大統工作助手`
- 安裝本機 PaddleOCR 引擎到 `%USERPROFILE%\Documents\大統工作助手\.venv-paddleocr`
- 設定使用者環境變數 `DATONG_WORKSPACE`
- 安裝完成後提醒我重開 Codex

安裝完成後，請先幫我跑以下檢查：

```powershell
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT='0'
& "$env:USERPROFILE\Documents\大統工作助手\.venv-paddleocr\Scripts\python.exe" -X utf8 -c "import paddleocr, paddle, openpyxl; print('ok')"
```

然後告訴我是否可以開始使用 `convert-vendor-invoice-image`。
