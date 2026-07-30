# 大統進貨助手

本機 Windows APP，不使用 ChatGPT、Codex 或雲端 AI API。APP 依序使用 OpenCV、PaddleOCR、RapidFuzz、openpyxl 與 Excel COM，完成進貨圖片 OCR、商品比對、人工覆核、名稱與大類調整、建檔用及採購單匯入 `.xls`。

APP 啟動檢查、OCR、產品比對、覆核與 PowerShell／Excel 正式輸出都以背景隱藏模式執行，不會另外彈出 CMD 或 PowerShell 視窗。

## 一鍵安裝

到 [GitHub Releases](https://github.com/ontaiko/datong-vendor-invoice-workflow-kit/releases) 下載最新版：

```text
Datong-Invoice-Assistant-Setup-v1.0.0.exe
```

雙擊後會自動完成：

1. 將 APP 安裝到 `%LOCALAPPDATA%\Programs\大統進貨助手`。
2. 建立開始功能表與桌面捷徑。
3. 偵測或安裝 Python 3.12，並在 APP 資料夾建立專用 OCR 環境。
4. 安裝 PaddleOCR、OpenCV、RapidFuzz、openpyxl 等本機套件。
5. 安裝並逐檔驗證離線 OCR 模型。
6. 自動產生目前電腦可用的 `app_settings.json`，最後執行 APP 自我檢查。

第一次安裝需要網路下載 Python 套件，依網路與電腦速度可能需要數分鐘；完成後日常 OCR、產品比對與命名規則不需要 ChatGPT、Codex 或雲端 AI。正式輸出 `.xls` 仍需要桌面版 Microsoft Excel。

安裝包尚未使用商用程式碼簽章憑證；Windows SmartScreen 可能顯示未知發行者。請只從本 repo 的 Releases 下載，並以 Release 一併提供的 `.sha256` 核對檔案。

### 更新與資料保留

直接執行新版安裝程式即可升級。升級時會：

- 覆蓋 APP、流程腳本與 OCR 引擎規格。
- 保留 `app_settings.json`、產品資料、命名規則、廠商代號、OCR 設定與匯入範本。
- 先把既有設定及 `reference_data` 備份到 `user-data-backups\pre-update-日期時間`。
- 重新使用原有專用 OCR 環境；只補裝有變動的套件。

從 Windows「已安裝的應用程式」解除安裝時，APP 與專用 Python 環境會移除，但設定、參考資料與升級備份會保留，避免日後重裝遺失規則。

開始功能表的「驗證安裝」可隨時重新檢查檔案、Python 套件與 APP 自我測試。

## 開啟 APP

雙擊：

```text
大統進貨助手.exe
```

資料夾只保留 `大統進貨助手.exe`，避免同一個 APP 出現兩份相同的執行檔。

## 操作順序

1. 選擇一張或多張 `.jpg`、`.jpeg`、`.png` 或 `.webp` 進貨圖片；多張圖片必須是同一廠商。
2. 選擇 OCR 中間檔輸出資料夾。
3. 按「開始 OCR」；APP 只做圖片辨識，多張同廠商圖片會合併成一份 OCR 確認檔。
4. 若圖片偵測為不同廠商，APP 會停止合併並列出每張圖片的廠商。
5. OCR 完成後會跳出「請確認產品輸出資料為最新後再繼續」提醒。
6. 在 OCR 原文確認表核對品名、數量、進價、金額與合併總額。
7. 在「產品資料輸出」欄選擇今天匯出的 CSV，再按「開始產品重複比對」。
8. 候選分成「一般候選」（相似度 60% 以上）與「低相似候選」（55%～59.99%，最多 3 筆）；沒有合適候選時，可選取單筆後按「更低候選單筆檢查」，另外查看約 35%～54.99% 的人工候選。
9. 可雙擊第一筆「已建檔代號」後貼上 Excel 整欄代號，APP 會由目前列往下逐列填入、保留六位前導零，並從產品 CSV 帶入正式品名、自動勾選已建檔及更新狀態。
10. OCR／已建檔確認與名稱調整階段的所有可輸入資料欄位，都支援從選取列開始整欄貼上；中間空白儲存格會保留，不會讓後方資料錯位。
11. 大類分成「大類代號」與「大類名稱」兩欄；貼入有效大類代號會自動帶入大類名稱，貼入完全一致的大類名稱也會反帶代號。
12. 視窗頂端「大類 → 管理大類…」可查看 APP 目前內建的大類、新增或刪除大類，也可恢復內建預設。修改結果保存在 `app_settings.json`，不依賴外部大類 Excel／CSV。
13. OCR 原文確認、已建檔勾選與名稱調整階段都可選取一列或多列後按「刪除選取項目」；APP 會同步刪除中間 Excel 的商品列並重算總額。
14. 要複製整欄時，先點任一格或欄位標題，再按「複製整欄」；也可在欄位上按滑鼠右鍵，或使用 `Ctrl+Shift+C`。複製內容不含標題，可直接貼到 Excel。
15. 按「回到上一步」可依序返回：正式輸出前確認 → 名稱調整 → 已建檔勾選 → OCR 原文確認。名稱調整草稿會暫時保留；正式檔產生後不可返回。
16. 需要時可拆分品項；拆分後數量與金額合計必須保持一致。
17. 按「確認資料」；APP 會在背景建立正式輸出所需的內部交接資料。
18. 再次確認資料後，產生建檔用與採購單匯入 `.xls`。
19. 正式檔通過驗證後，內部交接資料會自動刪除；若輸出失敗則保留，方便從失敗步驟重跑。
20. 正式檔通過重開驗證後，所有來源圖片才會各自加上 `_已處理`。

使用「複製整欄」複製商品名稱時，每筆名稱只使用一個標準換行，貼到其他 Excel 不會在名稱中間插入空白列。

## 人工停止點

- 產品資料不是今天版本，或仍是安裝包舊快照。
- OCR 漏列、低信心、數量乘單價不符或總額不符。
- 新品、相似候選、缺產品代號或缺大類。
- 廠商無法從廠商代號表確認。
- 正式 `.xls` 範本、Excel COM 或重開驗證失敗。

一般候選與低相似候選都不會自動當成同商品；只有精確命中或有效產品代號才會自動勾選。廠商也不會自行猜定。

多圖批次只允許同一廠商。APP 不會把不同廠商的商品、總額或正式採購檔合併。

單張長圖若包含上下排列的多頁單據，APP 會以 OpenCV 尋找頁面空白帶，自動切成 2 或 3 頁後逐頁 OCR，並稽核頁碼與缺頁。若分頁品質較差，只額外嘗試一次不分頁版本，再依完整商品列、錨點、總額與疑點保留較佳結果。四頁以上建議拆成多張圖片後一次選取。

## 本機引擎與參考資料

APP 自帶流程腳本：

```text
scripts\
```

APP 使用並同步下列參考資料：

```text
reference_data\
```

- `產品資料輸出.CSV`
- `產品比對身份關鍵詞.csv`
- `品牌括號命名規則.csv`
- `廠商代號.xlsx`
- `OCR設定.json`
- `建檔用.xls`
- `採購單匯入範例.xls`

每次開始產品比對都會重新驗證產品 CSV。若第一次因過期而停止，更新工作區 `參考資料` 的 CSV 後可直接再次按比對，不必重開 APP；同步流程也不會用較舊檔案覆蓋今天的新副本。

大類代號、名稱與自動推論關鍵字已內建於 APP。第一次新增或刪除大類時，完整清單會寫入 `app_settings.json`；之後以這份設定為準。

## 設定檔

`app_settings.json`：

```json
{
  "workspace_root": "C:\\Users\\user\\Documents\\大統工作助手",
  "python_exe": "C:\\Users\\user\\AppData\\Local\\Programs\\Python\\Python312\\python.exe",
  "product_csv_path": "C:\\path\\to\\產品資料輸出.CSV"
}
```

一鍵安裝程式會依其他電腦的使用者路徑自動產生此檔，不要求人工修改。

## 自我檢查

```powershell
& ".\大統進貨助手.exe" --self-test
```

檢查項目包含工作區、Python、PaddleOCR、OpenCV、產品 CSV 是否存在、內建流程腳本、命名規則、廠商表與 `.xls` 範本。產品 CSV 是否為今天版本，會在按「開始產品重複比對」時另外強制檢查；不會阻擋單純 OCR。

正式輸出腳本固定使用 UTF-8 回傳中文路徑，避免從無主控台 EXE 執行時把建檔／採購檔路徑解碼成亂碼。

## 實圖回歸測試

```powershell
python -X utf8 ".\tests\run_real_image_regression.py" `
  --image "進貨圖片.jpg" `
  --expected-raw "歷史確認結果.xlsx" `
  --expected-adjusted "歷史確認結果_產品比對檢查[調整].xlsx" `
  --output-dir "測試輸出資料夾"
```

測試使用圖片副本，不改動原始廠商圖片；會比對 OCR 明細、產品比對欄位與正式 `.xls` 輸出。

同廠商多圖合併回歸測試：

```powershell
python -X utf8 ".\tests\run_multi_image_batch_regression.py" `
  --first-raw "第一張歷史 OCR.xlsx" `
  --second-raw "第二張歷史 OCR.xlsx" `
  --output-dir "測試輸出資料夾"
```

元大多行品名與欄位分列回歸測試：

```powershell
python -X utf8 ".\tests\run_yuanda_ocr_regression.py" `
  --raw "第一張元大 OCR.xlsx" `
  --raw "第二張元大 OCR.xlsx" `
  --raw "第三張元大 OCR.xlsx" `
  --output-dir "測試輸出資料夾"
```

## 重新打包

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\build_exe.ps1"
```

建置前會產生 `package-manifest.json`，記錄所有必要腳本、規則、資料、範本與安裝腳本的 SHA256。

重新製作安裝包：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\installer\build-installer.ps1" -Version "1.0.0"
```

輸出位於 `dist-installer`，並同時產生 SHA256 與 Release manifest。

## 已驗證實圖

- `S__69607455_已處理.jpg`：麗嬰 4 筆、總額 28,195；與歷史人工確認結果逐欄相同；建檔 4 筆、採購 4 筆。
- `S__69607465_已處理.jpg`：麗嬰 10 筆、總額 8,054；與歷史人工確認結果逐欄相同；全部已建檔，只產生採購 10 筆。
- 以上兩張麗嬰圖片批次合併：14 筆、合併總額 36,249；不同廠商測試會被拒絕合併。
- `S__63975144.jpg`、`S__63975145.jpg`、`S__63975146.jpg`：元大 6／3／5 筆，數量 97／48／60，金額 9,516／7,908／12,696；合併為 14 筆、數量 205、總額 30,120。
- 元大解析以 9 位品號為列錨點；多行品名依相鄰品號上緣分列，數量欄不再誤抓箱／袋數。

GitHub 套件不包含 `app_settings.json`、`build`、`dist`、舊版 EXE、個人路徑或測試暫存。首次使用請從 `app_settings.example.json` 建立本機設定。
