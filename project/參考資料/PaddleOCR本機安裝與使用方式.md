# PaddleOCR 本機安裝與使用方式

## 安裝位置

- 專案環境：`C:\Users\user\Documents\大統工作助手\.venv-paddleocr`
- Python：`.venv-paddleocr\Scripts\python.exe`
- PaddlePaddle：`3.3.1`
- PaddleOCR：`3.7.0`
- openpyxl：`3.1.5`
- 模型快取：`C:\Users\user\.paddlex\official_models`

## Windows CPU 執行注意事項

目前 Windows CPU 推論會遇到 PaddlePaddle oneDNN 相容問題。執行 OCR 前先設定：

```powershell
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT='0'
```

## 驗證指令

```powershell
cd "C:\Users\user\Documents\大統工作助手"
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT='0'
.\.venv-paddleocr\Scripts\python.exe -X utf8 -c "import paddle; print(paddle.__version__); paddle.utils.run_check(); import paddleocr; print(paddleocr.__version__)"
```

## Python 使用範例

```python
from paddleocr import PaddleOCR

ocr = PaddleOCR(lang="ch")
result = ocr.predict(r"C:\path\to\image.png")

for page in result:
    print(page["rec_texts"])
```

## 進貨圖片轉表腳本

```powershell
cd "C:\Users\user\Documents\大統工作助手"
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT='0'
.\.venv-paddleocr\Scripts\python.exe -X utf8 "C:\Users\user\.codex\skills\extract-vendor-invoice-image\scripts\local_paddleocr_invoice_to_xlsx.py" "C:\path\to\invoice.jpg"
```

輸出會放在：

`C:\Users\user\Documents\大統工作助手\建檔進貨用\進貨圖片轉試算表`

## 已完成測試

測試圖片：

`C:\Users\user\Documents\大統工作助手\.codex-tmp\ocr-test\paddleocr_test.png`

成功辨識文字：

- `測試品名 NS2 遊戲片`
- `數量 3 進價 1280 金額 3840`
