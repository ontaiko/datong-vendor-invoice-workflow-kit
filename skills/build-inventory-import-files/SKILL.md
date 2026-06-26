---
name: build-inventory-import-files
description: 建檔採購匯入、建檔用檔案、採購單匯入檔、進貨匯入 xls、成本含稅檢查、廠商代號填入。Use when building inventory setup and purchase import .xls files from user-reviewed adjusted invoice xlsx files while preserving legacy Excel templates.
---

# 建檔與採購匯入檔

## 目標

依使用者已檢查完成的調整後進貨單 `.xlsx` 產生正式 `.xls`。此技能只負責套用範本、成本含稅判斷、廠商代號與輸出驗證，不負責圖片 OCR、產品資料比對、產品代號檢查或大類推論。

## 必要參考資料

- `參考資料/建檔用.xls`
- `參考資料/採購單匯入範例.xls`
- 優先使用 `參考資料/廠商代號.xlsx`；若不存在，才使用最新的 `參考資料/廠商代號*.xlsx`

## 輸入 xlsx

只參照經 `$review-invoice-product-check` 覆核調整後、且使用者確認可建檔的進貨單 `.xlsx`，不要再要求或讀取 JSON。輸入表必須符合既定格式並包含：

- `產品代號`
- `品名`
- `零售價`
- `數量`
- `進價`
- `金額`，若單據有列印金額
- `大類`，由 `$review-invoice-product-check` 事先處理；本技能不再檢查或推論
- `比對狀態` 必填；只有 `已建檔` 不寫入建檔用，其餘列寫入建檔用。沒有此欄時停止，避免將已建檔商品誤判為新品。

## 工作流程

1. 在建立任何目標 `.xls` 前，必須先暫停並請使用者確認：已檢查並調整完進貨單資料，可以進行建檔與採購匯入檔產生。
2. 使用者明確確認後，才可執行 `scripts/fill-import-templates.ps1`，並傳入 `-ConfirmedReviewed`。
3. 讀取廠商代號表，優先參考 `參考資料/廠商代號.xlsx`，依表內格式填入，不補零；若未手動指定廠商代號，腳本必須自動用廠商簡稱查表。若廠商為萬榮或萬榮國際，固定使用廠商代號 `38`。
4. 檢查每筆 `數量 × 進價` 是否符合 `金額`。若只有乘以 `1.05` 後符合，將成本改為 `進價 × 1.05`。若廠商為南波或南波丸，不論列印金額是否已可對上，成本一律改為 `進價 × 1.05`。
5. 產品代號與大類只依輸入 `.xlsx` 現有內容寫入，不在本技能中補值、改值、檢查六位數或推論大類。
   - 但所有商品必須有 `產品代號`；非 `已建檔` 商品必須有 `大類`。
6. 建檔用 `.xls` 只放新商品；全部都是既有商品時不建立建檔用檔案。
   - 建檔用範本若包含名為 `分類` 的工作表，輸出正式檔案前必須刪除，不可留在建檔用成品內。
   - 同一次流程若有多個廠商，建檔用只產生一份合併檔，彙整所有廠商的新商品，不依廠商拆檔。
7. 採購單匯入 `.xls` 放本次全部商品；多廠商時仍依廠商分開產生，因為採購單需要各自的廠商代號。
8. 輸出到工作區根目錄 `建檔進貨用/`。檔名固定為 `[廠商簡稱][建檔用or採購單用][日期]-[流水號].xls`，例如 `萬榮建檔用1150606-01.xls`、`萬榮採購單用1150606-01.xls`。多廠商合併建檔用檔名的廠商簡稱使用該次流程的共同簡稱或 `多廠商`。若同組檔名已存在，流水號遞增為 `02`、`03`，不得覆蓋既有檔案。
9. 重新開啟輸出檔，確認格式為 `.xls`、標題列符合範本、使用範圍只有標題與實際商品列。
10. 本技能不主動刪除輸入或中間檔；正式成品驗證後的中間檔清理由 `$convert-vendor-invoice-image` 統一負責。

## 腳本

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\fill-import-templates.ps1 `
  -WorkspaceRoot "C:\path\to\workspace" `
  -ProductsXlsx "C:\path\to\產品比對檢查.xlsx" `
  -VendorCode "40" `
  -VendorShortName "萬榮" `
  -InvoiceTotal "6492" `
  -OutputDir "C:\path\to\建檔進貨用" `
  -ConfirmedReviewed
```

`-VendorCode` 可省略；省略時會優先從 `參考資料/廠商代號.xlsx` 以廠商簡稱查表。

詳細欄位規則見 [references/import-rules.md](references/import-rules.md)。
