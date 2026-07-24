---
name: convert-vendor-invoice-image
description: Detect and coordinate the user's vendor invoice workflow for uploaded vendor invoices, sales slips, delivery notes, and product purchase detail images. Route extraction, catalog matching, product review, user confirmation, inventory setup, purchase import, and cleanup without redefining child-skill rules.
---

# 進貨流程管理

## 責任

只負責辨識進貨文件、安排階段、傳遞檔案、保存狀態、處理停止點與回報結果。不得重寫子技能的 OCR、比對、命名、大類、稅額、廠商代號或匯入格式規則；發生衝突時以當前子技能為準。

圖片含廠商/客戶、日期、品名、數量、單價、金額、合計或表格型商品明細時啟動。遊戲封面、網站商品圖與一般照片不要啟動。

## 省 Token 執行方式

1. 啟動時只讀本檔，不要預先讀取全部子技能。
2. 依狀態檔或現有成果判斷目前階段；只讀並執行該階段的子技能及其直接要求的參考檔。
3. 子技能腳本正常時直接執行，不讀程式碼。只有執行失敗、輸出契約不明或使用者要求修改時才讀相關程式碼。
4. 完整 OCR、工作表內容、候選清單與執行紀錄保留在檔案，不貼入對話。對話只列疑點、待確認列、排除項目、統計與檔案連結。
5. 每次工具輸出只取判斷下一步需要的欄位；不要整本列印 xlsx、CSV 或 OCR 日誌。

## 階段路由

| 階段 | 執行內容 | 完成依據 |
|---|---|---|
| `reference-precheck` | 執行 `scripts/check-product-csv-date.py` | 今日產品 CSV 通過 |
| `extract` | `$extract-vendor-invoice-image` | 原始進貨 xlsx |
| `match` | `$match-product-catalog --no-suggestion-txt` | `_產品比對檢查.xlsx` |
| `review` | `$review-invoice-product-check` | 使用者確認後產生 `[調整].xlsx` |
| `build` | `$build-inventory-import-files` | 正式建檔用/採購單用檔案 |
| `cleanup` | 本技能狀態工具 | 成品驗證後刪除已登記中間檔 |

只有比對結果存在新品、相似商品、缺代號、缺大類或其他待確認項目時才進入 `review`；否則依建檔技能規則直接進入 `build`。不要為了流程完整而載入不需要的子技能。

## 狀態檔

以 `scripts/workflow-state.py` 維護小型 JSON 狀態檔，建議放在 `<工作區>/.codex-tmp/invoice-workflow/`。狀態只保存來源檔、目前階段、當前子技能、成果路徑、待確認項目、排除項目與筆數，不保存完整 OCR 或工作表資料。

```powershell
python scripts/workflow-state.py init --state <state.json> --workspace-root <工作區> --source <圖片>
python scripts/workflow-state.py update --state <state.json> --stage match --status active --current-skill match-product-catalog
python scripts/workflow-state.py show --state <state.json> --compact
```

每個子技能完成後，登記其正式交接檔為 `--artifact 名稱=路徑`；待清理檔案另以 `--intermediate 路徑` 登記，成品以 `--final 路徑` 登記。跨回合先讀 compact 狀態，不重新展開前面階段的內容。

## 停止與繼續

遇到以下情況停止並只列需要使用者補充的內容：

- 產品 CSV 不是今天版本、不存在，或仍是 GitHub 安裝包隨附的舊快照。
- OCR 關鍵欄位低信心。
- 新品、相似商品、缺代號、缺大類或拆分方式不明。
- 子技能要求正式輸出前確認。
- 金額、範本、Excel COM、檔案權限或輸出驗證失敗。

使用者回覆後，從狀態檔記錄的階段與交接檔繼續，不重跑已完成階段。使用者提供代號、正式名稱與大類時，依覆核技能要求建立 UTF-8 TSV，直接產生 `[調整].xlsx`。

## 清理與回報

成品存在且通過子技能驗證後，執行：

```powershell
python scripts/workflow-state.py cleanup --state <state.json> --remove-state
```

清理只處理狀態檔明確登記、位於工作區內且不屬於來源圖片、正式成品或 `參考資料` 的中間檔。驗證失敗時保留全部中間檔。

正式成品存在且通過驗證後，最後必須把來源進貨圖片加上 `_已處理` 後綴再回報。只改來源圖片，不改 OCR 暫存圖、中間 xlsx 或正式 xls。檔名格式為 `原檔名_已處理.副檔名`，例如 `1784533042698.jpg` 改為 `1784533042698_已處理.jpg`；如果來源檔名已經包含 `_已處理`，不要重複加；如果目標檔名已存在，不得覆蓋，改用 `_已處理_2`、`_已處理_3` 避免撞名。只有正式建檔用/採購單用檔案完成並驗證通過後才改名；若流程停在 OCR、比對、覆核、使用者確認或輸出失敗，不要改名。

最後只回報一份簡短驗證摘要，固定包含：正式檔案連結、來源圖片、商品/新品/已建檔/採購筆數、廠商代號、採購日期、單據總額、逐筆數量×單價核對方式、含稅調整、排除項目、特殊處理、來源圖片改名結果，以及中間檔是否已清理。來源圖片改名結果必須列出改名後路徑，例如 `來源圖片：已改名為 xxx_已處理.jpg`。大類一律顯示 `代號 中文名稱`；待填清單使用可複製的 Tab 分隔格式且不得以空白欄開頭。
