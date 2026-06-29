# 本機 PaddleOCR 規則

## 預設路線

進貨圖片擷取預設先使用本機 OpenCV 做圖片前處理，再使用本機 PaddleOCR，不先使用 AI 視覺 OCR。
照片先偵測是否包含上下排列的多張單據。單頁只產生一張 OCR 暫存圖；多頁依頁縫切開後每頁各產生一張，逐頁辨識並合併商品列。每頁必要時旋轉後用 OpenCV 做灰階、去雜訊、局部對比與輕微銳化，不產生同頁多種增強版。若 `rotation auto` 分數接近則保留原方向；第一次結果幾乎沒有欄位關鍵字或貨號命中時，才自動改試相反旋轉方向一次。OpenCV 不可用時才退回 PIL 灰階/對比/銳化。

使用環境：

- 專案 Python：`C:\Users\user\Documents\大統工作助手\.venv-paddleocr\Scripts\python.exe`
- 腳本：`scripts/local_paddleocr_invoice_to_xlsx.py`
- 輸出資料夾：`C:\Users\user\Documents\大統工作助手\建檔進貨用\進貨圖片轉試算表`
- 設定檔：`C:\Users\user\Documents\大統工作助手\參考資料\OCR設定.json`

Windows CPU 執行前必須設定：

```powershell
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT='0'
```

## 執行指令

```powershell
cd "C:\Users\user\Documents\大統工作助手"
$env:PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT='0'
.\.venv-paddleocr\Scripts\python.exe -X utf8 "C:\Users\user\.codex\skills\extract-vendor-invoice-image\scripts\local_paddleocr_invoice_to_xlsx.py" "C:\path\to\invoice.jpg"
```

可用參數：

- `--settings`：指定 OCR 設定 JSON；未指定時讀取 `參考資料/OCR設定.json`。
- `--rotation auto`：預設，自動判斷是否旋轉後再灰階增強；若第一次 OCR 結果太差，會自動試相反旋轉方向一次。
- `--rotation none`：強制不旋轉，只做灰階增強。
- `--rotation cw` / `--rotation ccw` / `--rotation 180`：手動指定旋轉方向。
- `--multi-variant`：回到舊式多版本 OCR；僅在單張灰階增強圖無法讀出表頭或商品列時使用。多版本模式會在結果達到可用門檻後早停。
- `--contrast`：覆寫灰階對比倍率。
- `--sharpness`：覆寫銳化倍率。
- `--page-split auto`：預設，自動偵測上下排列多張單據並逐頁辨識。
- `--page-split off`：停用自動分頁，整張照片只視為一張單據。

執行 OCR 指令時，外層命令時限至少設為 180 秒。模型建立或辨識仍在執行時應等待同一程序，不要因 5 至 10 秒沒有輸出就重啟；一般模型載入與高解析圖片辨識時間分開記錄在 `OCR測試紀錄`。

腳本會在 `.codex-tmp/local-paddleocr/.locks/` 依圖片絕對路徑建立執行鎖。同一張圖片已有存活程序時，第二次啟動會停止並顯示原 PID；正常完成、錯誤或中斷離開主流程時都會釋放自己的鎖，已失效 PID 的舊鎖會在下次啟動時自動移除。

本地引擎責任：

- OpenCV：只負責 OCR 前圖片清理與增強，不負責判斷商品是否正確。
- PaddleOCR：只負責把圖片文字轉成可覆核資料。
- AI 視覺：只在本地 OCR 失敗、表頭/商品列無法讀出，或使用者明確要求時作為備援。

`OCR設定.json` 可調整：

```json
{
  "rotation": "auto",
  "page_split": "auto",
  "multi_variant": false,
  "contrast": 1.8,
  "sharpness": 1.15
}
```

## 腳本輸出

產生一個 `.xlsx`，包含：

- `進貨明細`：後續 `$match-product-catalog` 使用的正式交接表。
- `OCR測試紀錄`：最佳旋轉版本、平均信心、原始 OCR 文字、商品列疑點與金額核對；版本選擇優先看可解析商品列與金額核對通過列數。
- 多頁照片另記錄偵測頁數與各頁暫存圖，商品列逐頁解析後合併。
- `OCR測試紀錄` 會記錄模型載入秒數與 OCR 總秒數，用來判斷是否需要再調整速度策略。

`進貨明細` 仍必須遵守原本欄位規則：

- `產品代號` 一律留空。
- 單據上的 `貨號`、`品號`、`客戶品號` 不寫入本店 `產品代號`。
- 不建立 `單位` 欄。
- 單據上的單位值，例如 `抽`、`個`、`PCS`，必須在轉表時排除，不得併入 `品名`。
- 有明確 `單位` 欄時，品名合併範圍只到單位欄之前；不得跨欄抓到單位、數量、單價、金額或備註欄文字。
- 最後一列放 `總價格`。

## 停止與人工確認

下列情況不得直接交給後續建檔或採購匯入，必須先回報疑點：

- 腳本輸出的 `needs_review` 為 `true`。
- `OCR測試紀錄` 中任一商品列的 `金額核對` 不是 `通過`。
- 任一商品列的 `OCR疑點` 不是 `無`。
- 商品列數明顯少於原圖表格列數。
- 商品合計無法對上單據合計或總計。
- 手拍圖有手寫圈選、摺痕、淡字、歪斜、反光、壓線或欄位黏在一起。

## AI 視覺備援

只有在下列情況才改用 AI 視覺判讀輔助：

- PaddleOCR 環境不存在或無法啟動。
- PaddleOCR 無法讀出表頭或商品列。
- 原圖非常模糊、被遮蔽或表格線破碎，導致本機 OCR 無法產出可覆核 xlsx。
- 使用者明確要求使用 AI 視覺 OCR。

即使使用 AI 視覺備援，最後仍要輸出同樣格式的 xlsx，並保留人工確認停點。
