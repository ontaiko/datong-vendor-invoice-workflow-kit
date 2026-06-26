param(
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceRoot,

    [Parameter(Mandatory = $true)]
    [string]$ProductsXlsx,

    [string]$VendorCode = "",

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$PurchaseDate = "",

    [string]$Note1 = "阿榮代打",

    [string]$NewProductTemplate = "",

    [string]$PurchaseTemplate = "",

    [string]$InvoiceTotal = "",

    [string]$VendorShortName = "",

    [switch]$ConfirmedReviewed
)

$SummaryRowNames = @("總價格", "總價", "總計", "合計", "小計", "稅金", "稅額", "折扣", "總數量", "合計數量", "頁碼", "頁次")

function Test-SummaryRowText {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $true }
    $normalized = $Text.Trim().TrimEnd("：", ":")
    foreach ($name in $SummaryRowNames) {
        if ($normalized -eq $name -or $normalized.StartsWith($name)) { return $true }
    }
    return $false
}

$ErrorActionPreference = "Stop"

function Get-RocDate {
    param([string]$Value)

    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    $taipei = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId(
        [DateTimeOffset]::UtcNow,
        "Taipei Standard Time"
    )
    return ("{0:000}.{1:00}.{2:00}" -f ($taipei.Year - 1911), $taipei.Month, $taipei.Day)
}

function Assert-Headers {
    param(
        $Worksheet,
        [string[]]$Expected
    )

    for ($column = 1; $column -le $Expected.Count; $column++) {
        $actual = [string]$Worksheet.Cells.Item(1, $column).Text
        if ($actual -ne $Expected[$column - 1]) {
            throw "Unexpected template header at column $column. Expected '$($Expected[$column - 1])', got '$actual'."
        }
    }
}

function Release-ComObjects {
    param([object[]]$Objects)

    foreach ($object in $Objects) {
        if ($null -ne $object) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($object) | Out-Null
        }
    }
}

function Test-PrintedAmountMatch {
    param(
        [double]$Calculated,
        [double]$Printed
    )

    $calculatedRounded = [Math]::Round($Calculated, 0, [MidpointRounding]::AwayFromZero)
    $printedRounded = [Math]::Round($Printed, 0, [MidpointRounding]::AwayFromZero)
    return [Math]::Abs($calculatedRounded - $printedRounded) -lt 0.01
}

function Resolve-TaxInclusiveCosts {
    param(
        [object[]]$Items,
        [string]$PrintedInvoiceTotal
    )

    $adjustedCodes = @()
    $itemsWithoutLineAmount = @()

    foreach ($item in $Items) {
        $quantity = [double]$item.quantity
        $cost = [double]$item.cost

        if ($null -eq $item.lineAmount -or [string]::IsNullOrWhiteSpace([string]$item.lineAmount)) {
            $itemsWithoutLineAmount += $item
            continue
        }

        $lineAmount = [double]$item.lineAmount
        if (Test-PrintedAmountMatch -Calculated ($quantity * $cost) -Printed $lineAmount) {
            continue
        }

        if (Test-PrintedAmountMatch -Calculated ($quantity * $cost * 1.05) -Printed $lineAmount) {
            $item.cost = [Math]::Round($cost * 1.05, 6)
            $adjustedCodes += [string]$item.productCode
            continue
        }

        throw "Cost mismatch for product '$($item.productCode)': quantity * cost and quantity * cost * 1.05 both differ from printed line amount '$lineAmount'."
    }

    if ([string]::IsNullOrWhiteSpace($PrintedInvoiceTotal)) {
        return $adjustedCodes
    }

    $invoiceTotalValue = [double]$PrintedInvoiceTotal
    $currentTotal = 0.0
    foreach ($item in $Items) {
        $currentTotal += [Math]::Round(
            ([double]$item.quantity * [double]$item.cost),
            0,
            [MidpointRounding]::AwayFromZero
        )
    }

    if (Test-PrintedAmountMatch -Calculated $currentTotal -Printed $invoiceTotalValue) {
        return $adjustedCodes
    }

    if ($itemsWithoutLineAmount.Count -gt 0) {
        $mixedTaxTotal = 0.0
        foreach ($item in $Items) {
            $cost = [double]$item.cost
            if ($itemsWithoutLineAmount -contains $item) {
                $cost *= 1.05
            }
            $mixedTaxTotal += [Math]::Round(
                ([double]$item.quantity * $cost),
                0,
                [MidpointRounding]::AwayFromZero
            )
        }

        if (Test-PrintedAmountMatch -Calculated $mixedTaxTotal -Printed $invoiceTotalValue) {
            foreach ($item in $itemsWithoutLineAmount) {
                $item.cost = [Math]::Round(([double]$item.cost * 1.05), 6)
                $adjustedCodes += [string]$item.productCode
            }
            return $adjustedCodes
        }
    }

    throw "Invoice total mismatch: calculated total and tax-adjusted total both differ from printed invoice total '$invoiceTotalValue'."
}

function Test-AlwaysTaxInclusiveVendor {
    param([string]$VendorName)

    if ([string]::IsNullOrWhiteSpace($VendorName)) {
        return $false
    }
    return $VendorName -match "南波"
}

function Apply-AlwaysTaxInclusiveCosts {
    param(
        [object[]]$Items,
        [string[]]$AlreadyAdjustedCodes
    )

    $adjustedCodes = @()
    $alreadyAdjusted = @{}
    foreach ($code in $AlreadyAdjustedCodes) {
        if (-not [string]::IsNullOrWhiteSpace($code)) {
            $alreadyAdjusted[[string]$code] = $true
        }
    }

    foreach ($item in $Items) {
        $code = [string]$item.productCode
        if ($alreadyAdjusted.ContainsKey($code)) {
            continue
        }

        $item.cost = [Math]::Round(([double]$item.cost * 1.05), 6)
        $adjustedCodes += $code
    }

    return $adjustedCodes
}

function Find-HeaderColumn {
    param(
        $Worksheet,
        [int]$HeaderRow,
        [string]$HeaderName,
        [bool]$Required = $true
    )

    $usedColumns = $Worksheet.UsedRange.Columns.Count
    for ($column = 1; $column -le $usedColumns; $column++) {
        $actual = [string]$Worksheet.Cells.Item($HeaderRow, $column).Text
        if ($actual -eq $HeaderName) {
            return $column
        }
    }

    if ($Required) {
        throw "Input xlsx is missing required column '$HeaderName'."
    }
    return 0
}

function Get-VendorShortNameFromXlsx {
    param(
        [string]$Path,
        [string]$Fallback
    )

    if (-not [string]::IsNullOrWhiteSpace($Fallback)) {
        return $Fallback
    }

    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($Path)
        $worksheet = $workbook.Worksheets.Item(1)

        for ($row = 1; $row -le [Math]::Min($worksheet.UsedRange.Rows.Count, 5); $row++) {
            for ($column = 1; $column -le [Math]::Min($worksheet.UsedRange.Columns.Count, 5); $column++) {
                $text = [string]$worksheet.Cells.Item($row, $column).Text
                if ($text -match "^廠商：(.+)$") {
                    $vendorName = $Matches[1].Trim()
                    $vendorName = $vendorName -replace "國際企業股份有限公司", ""
                    $vendorName = $vendorName -replace "國際股份有限公司", ""
                    $vendorName = $vendorName -replace "股份有限公司", ""
                    $vendorName = $vendorName -replace "有限公司", ""
                    $vendorName = $vendorName.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($vendorName)) {
                        return $vendorName
                    }
                }
            }
        }
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false)
        }
        if ($null -ne $excel) {
            $excel.Quit()
        }
        Release-ComObjects @($worksheet, $workbook, $excel)
    }

    return "進貨"
}

function Get-VendorCodeWorkbookPath {
    param([string]$WorkspaceRoot)

    $preferred = Join-Path $WorkspaceRoot "參考資料\廠商代號.xlsx"
    if (Test-Path -LiteralPath $preferred) {
        return (Resolve-Path -LiteralPath $preferred).Path
    }

    $referenceDir = Join-Path $WorkspaceRoot "參考資料"
    $candidates = @(Get-ChildItem -LiteralPath $referenceDir -Filter "廠商代號*.xlsx" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    if ($candidates.Count -gt 0) {
        return $candidates[0].FullName
    }

    return ""
}

function Resolve-VendorCodeFromWorkbook {
    param(
        [string]$WorkbookPath,
        [string]$VendorName
    )

    if ([string]::IsNullOrWhiteSpace($VendorName)) {
        throw "Vendor short name is empty; cannot resolve vendor code."
    }
    if ([string]::IsNullOrWhiteSpace($WorkbookPath) -or -not (Test-Path -LiteralPath $WorkbookPath)) {
        throw "找不到廠商代號表，請確認 參考資料\廠商代號.xlsx 是否存在。"
    }

    $searchName = $VendorName.Trim()
    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($WorkbookPath)

        for ($sheetIndex = 1; $sheetIndex -le $workbook.Worksheets.Count; $sheetIndex++) {
            $worksheet = $workbook.Worksheets.Item($sheetIndex)
            $usedRows = $worksheet.UsedRange.Rows.Count
            $usedColumns = $worksheet.UsedRange.Columns.Count
            for ($row = 1; $row -le $usedRows; $row++) {
                $rowValues = @()
                $matched = $false
                for ($column = 1; $column -le $usedColumns; $column++) {
                    $text = [string]$worksheet.Cells.Item($row, $column).Text
                    $rowValues += $text
                    if (-not [string]::IsNullOrWhiteSpace($text) -and ($text -eq $searchName -or $text.Contains($searchName))) {
                        $matched = $true
                    }
                }

                if ($matched) {
                    foreach ($value in $rowValues) {
                        if (-not [string]::IsNullOrWhiteSpace($value)) {
                            return $value.Trim()
                        }
                    }
                }
            }
        }
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false)
        }
        if ($null -ne $excel) {
            $excel.Quit()
        }
        Release-ComObjects @($workbook, $excel)
    }

    throw "廠商代號表找不到廠商 '$VendorName'，請更新 參考資料\廠商代號.xlsx 或手動指定 -VendorCode。"
}

function Convert-ToSafeFileNamePart {
    param([string]$Value)

    $safe = $Value
    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace([string]$char, "")
    }
    return $safe.Trim()
}

function Read-ProductsFromXlsx {
    param([string]$Path)

    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($Path)
        $worksheet = $workbook.Worksheets.Item(1)
        $usedRows = $worksheet.UsedRange.Rows.Count

        $headerRow = 0
        for ($row = 1; $row -le [Math]::Min($usedRows, 20); $row++) {
            for ($column = 1; $column -le $worksheet.UsedRange.Columns.Count; $column++) {
                if ([string]$worksheet.Cells.Item($row, $column).Text -eq "產品代號") {
                    $headerRow = $row
                    break
                }
            }
            if ($headerRow -gt 0) {
                break
            }
        }

        if ($headerRow -eq 0) {
            throw "Input xlsx is missing header column '產品代號'."
        }

        $columns = @{
            productCode = Find-HeaderColumn -Worksheet $worksheet -HeaderRow $headerRow -HeaderName "產品代號"
            name = Find-HeaderColumn -Worksheet $worksheet -HeaderRow $headerRow -HeaderName "品名"
            recommendedPrice = Find-HeaderColumn -Worksheet $worksheet -HeaderRow $headerRow -HeaderName "零售價" -Required $false
            quantity = Find-HeaderColumn -Worksheet $worksheet -HeaderRow $headerRow -HeaderName "數量"
            cost = Find-HeaderColumn -Worksheet $worksheet -HeaderRow $headerRow -HeaderName "進價"
            lineAmount = Find-HeaderColumn -Worksheet $worksheet -HeaderRow $headerRow -HeaderName "金額" -Required $false
            status = Find-HeaderColumn -Worksheet $worksheet -HeaderRow $headerRow -HeaderName "比對狀態" -Required $false
            category = Find-HeaderColumn -Worksheet $worksheet -HeaderRow $headerRow -HeaderName "大類" -Required $false
        }

        $items = @()
        for ($row = $headerRow + 1; $row -le $usedRows; $row++) {
            $name = [string]$worksheet.Cells.Item($row, $columns.name).Text
            if (Test-SummaryRowText -Text $name) {
                continue
            }

            $quantityText = [string]$worksheet.Cells.Item($row, $columns.quantity).Text
            $costText = [string]$worksheet.Cells.Item($row, $columns.cost).Text
            if ([string]::IsNullOrWhiteSpace($quantityText) -and [string]::IsNullOrWhiteSpace($costText)) {
                continue
            }

            $codeCell = $worksheet.Cells.Item($row, $columns.productCode)
            $codeText = [string]$codeCell.Text

            $status = if ($columns.status -gt 0) { [string]$worksheet.Cells.Item($row, $columns.status).Text } else { "" }
            $existingProduct = if ($columns.status -gt 0) { $status -match "已建檔" } else { $false }
            $category = if ($columns.category -gt 0) { [string]$worksheet.Cells.Item($row, $columns.category).Text } else { "" }

            $items += [PSCustomObject]@{
                productCode = $codeText
                name = $name
                recommendedPrice = if ($columns.recommendedPrice -gt 0) { [string]$worksheet.Cells.Item($row, $columns.recommendedPrice).Text } else { "" }
                quantity = $quantityText
                cost = $costText
            lineAmount = if ($columns.lineAmount -gt 0) { [string]$worksheet.Cells.Item($row, $columns.lineAmount).Text } else { "" }
            existingProduct = $existingProduct
            status = $status
            category = $category
        }
        }

        return $items
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false)
        }
        if ($null -ne $excel) {
            $excel.Quit()
        }
        Release-ComObjects @($worksheet, $workbook, $excel)
    }
}

function Remove-TrailingBlankRows {
    param(
        $Worksheet,
        [int]$LastDataRow,
        [string]$LastColumn
    )

    $deleteStart = $LastDataRow + 1
    if ($deleteStart -le 65536) {
        $Worksheet.Range("A${deleteStart}:${LastColumn}65536").EntireRow.Delete() | Out-Null
    }
}

function Get-UniqueOutputPath {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $Path
    }

    $directory = Split-Path -Parent $Path
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [System.IO.Path]::GetExtension($Path)

    for ($index = 2; $index -lt 1000; $index++) {
        $candidate = Join-Path $directory ("{0}_{1}{2}" -f $baseName, $index, $extension)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    throw "Unable to find a unique output path for '$Path'."
}

function Get-OutputPairPaths {
    param(
        [string]$Directory,
        [string]$VendorShortName,
        [string]$FileDate,
        [bool]$IncludeNewProduct
    )

    $safeVendorShortName = Convert-ToSafeFileNamePart -Value $VendorShortName
    if ([string]::IsNullOrWhiteSpace($safeVendorShortName)) {
        $safeVendorShortName = "進貨"
    }

    for ($index = 1; $index -lt 1000; $index++) {
        $serial = "{0:00}" -f $index
        $newProductPath = Join-Path $Directory ("{0}建檔用{1}-{2}.xls" -f $safeVendorShortName, $FileDate, $serial)
        $purchasePath = Join-Path $Directory ("{0}採購單用{1}-{2}.xls" -f $safeVendorShortName, $FileDate, $serial)
        $newProductExists = $IncludeNewProduct -and (Test-Path -LiteralPath $newProductPath)
        $purchaseExists = Test-Path -LiteralPath $purchasePath
        if (-not $newProductExists -and -not $purchaseExists) {
            return [PSCustomObject]@{
                newProduct = $newProductPath
                purchase = $purchasePath
            }
        }
    }

    throw "Unable to find a unique output pair for '$safeVendorShortName' on '$FileDate'."
}

function Remove-WorksheetIfExists {
    param(
        $Workbook,
        [string]$WorksheetName
    )

    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $Workbook.Worksheets.Item($index)
        if ([string]$worksheet.Name -eq $WorksheetName) {
            if ($Workbook.Worksheets.Count -le 1) {
                return
            }
            $worksheet.Delete()
            return
        }
    }
}

$workspace = (Resolve-Path -LiteralPath $WorkspaceRoot).Path
$productsPath = (Resolve-Path -LiteralPath $ProductsXlsx).Path

if (-not $ConfirmedReviewed) {
    throw "建立建檔用與採購單匯入檔前，請先向使用者確認：進貨單資料已檢查並調整完成，可以進行建檔。確認後再以 -ConfirmedReviewed 執行。"
}

$output = New-Item -ItemType Directory -Force -Path $OutputDir
$outputPath = $output.FullName

if ([string]::IsNullOrWhiteSpace($NewProductTemplate)) {
    $NewProductTemplate = Join-Path $workspace "參考資料\建檔用.xls"
}
if ([string]::IsNullOrWhiteSpace($PurchaseTemplate)) {
    $PurchaseTemplate = Join-Path $workspace "參考資料\採購單匯入範例.xls"
}

$newProductTemplatePath = (Resolve-Path -LiteralPath $NewProductTemplate).Path
$purchaseTemplatePath = (Resolve-Path -LiteralPath $PurchaseTemplate).Path
$items = @(Read-ProductsFromXlsx -Path $productsPath)

if ($items.Count -eq 0) {
    throw "Products xlsx has no product rows."
}

$hasAnyStatus = @($items | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.status) }).Count -gt 0
if (-not $hasAnyStatus) {
    throw "Input xlsx is missing '比對狀態'. 請先使用產品比對/覆核後的檔案，避免將已建檔商品誤判為新品。"
}

foreach ($item in $items) {
    foreach ($property in @("productCode", "name", "cost", "quantity")) {
        if ($null -eq $item.$property -or [string]::IsNullOrWhiteSpace([string]$item.$property)) {
            throw "Missing required property '$property' in products xlsx."
        }
    }

    if (-not ($item.existingProduct -eq $true) -and [string]::IsNullOrWhiteSpace([string]$item.category)) {
        throw "Missing required column value '大類' for new product '$($item.productCode)' in products xlsx."
    }
}

$resolvedVendorShortName = Get-VendorShortNameFromXlsx -Path $productsPath -Fallback $VendorShortName
if ([string]::IsNullOrWhiteSpace($VendorCode)) {
    if ($resolvedVendorShortName -match "萬榮") {
        $VendorCode = "38"
    } else {
        $vendorWorkbookPath = Get-VendorCodeWorkbookPath -WorkspaceRoot $workspace
        $VendorCode = Resolve-VendorCodeFromWorkbook -WorkbookPath $vendorWorkbookPath -VendorName $resolvedVendorShortName
    }
}
$taxAdjustedCodes = @(Resolve-TaxInclusiveCosts -Items @($items) -PrintedInvoiceTotal $InvoiceTotal)
if (Test-AlwaysTaxInclusiveVendor -VendorName $resolvedVendorShortName) {
    $forcedTaxAdjustedCodes = @(Apply-AlwaysTaxInclusiveCosts -Items @($items) -AlreadyAdjustedCodes $taxAdjustedCodes)
    $taxAdjustedCodes = @($taxAdjustedCodes + $forcedTaxAdjustedCodes | Select-Object -Unique)
}
$newItems = @($items | Where-Object { -not ($_.existingProduct -eq $true) })

$dateText = Get-RocDate -Value $PurchaseDate
$fileDate = $dateText.Replace(".", "")
$outputPair = Get-OutputPairPaths -Directory $outputPath -VendorShortName $resolvedVendorShortName -FileDate $fileDate -IncludeNewProduct ($newItems.Count -gt 0)
$newProductOutput = $outputPair.newProduct
$purchaseOutput = $outputPair.purchase

if ($newItems.Count -gt 0) {
    Copy-Item -LiteralPath $newProductTemplatePath -Destination $newProductOutput
}
Copy-Item -LiteralPath $purchaseTemplatePath -Destination $purchaseOutput

$excel = $null
$newProductWorkbook = $null
$newProductWorksheet = $null
$purchaseWorkbook = $null
$purchaseWorksheet = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    if ($newItems.Count -gt 0) {
        $newProductWorkbook = $excel.Workbooks.Open($newProductOutput)
        Remove-WorksheetIfExists -Workbook $newProductWorkbook -WorksheetName "分類"
        $newProductWorksheet = $newProductWorkbook.Worksheets.Item("工作表1")
        Assert-Headers -Worksheet $newProductWorksheet -Expected @(
            "產品代號", "產品名稱", "銷售單價1", "銷售單價2", "建議售價",
            "單位", "成本", "CO128", "大類"
        )
        $newProductWorksheet.Range("A2:I65536").ClearContents() | Out-Null
        $newProductWorksheet.Columns.Item(1).NumberFormat = "@"
        $newProductWorksheet.Columns.Item(8).NumberFormat = "@"

        for ($index = 0; $index -lt $newItems.Count; $index++) {
            $row = $index + 2
            $item = $newItems[$index]
            $recommendedPrice = if ($null -eq $item.recommendedPrice -or [string]::IsNullOrWhiteSpace([string]$item.recommendedPrice)) { 0 } else { [double]$item.recommendedPrice }

            $newProductWorksheet.Cells.Item($row, 1).Value2 = [string]$item.productCode
            $newProductWorksheet.Cells.Item($row, 2).Value2 = [string]$item.name
            $newProductWorksheet.Cells.Item($row, 3).Value2 = [double]0
            $newProductWorksheet.Cells.Item($row, 4).Value2 = [double]0
            $newProductWorksheet.Cells.Item($row, 5).Value2 = $recommendedPrice
            $newProductWorksheet.Cells.Item($row, 6).Value2 = [string]"PCS"
            $newProductWorksheet.Cells.Item($row, 7).Value2 = [double]$item.cost
            $newProductWorksheet.Cells.Item($row, 8).Value2 = [string]$item.productCode
            if ([string]::IsNullOrWhiteSpace([string]$item.category)) {
                $newProductWorksheet.Cells.Item($row, 9).Value2 = [string]""
            } else {
                $newProductWorksheet.Cells.Item($row, 9).Value2 = [double]$item.category
            }
        }
        Remove-TrailingBlankRows -Worksheet $newProductWorksheet -LastDataRow ($newItems.Count + 1) -LastColumn "I"
        $newProductWorkbook.Save()
        $newProductWorkbook.Close($true)
        Release-ComObjects @($newProductWorksheet, $newProductWorkbook)
        $newProductWorksheet = $null
        $newProductWorkbook = $null
    }

    $purchaseWorkbook = $excel.Workbooks.Open($purchaseOutput)
    $purchaseWorksheet = $purchaseWorkbook.Worksheets.Item("工作表1")
    Assert-Headers -Worksheet $purchaseWorksheet -Expected @(
        "採購日期", "廠商代號", "外幣幣別", "產品代號", "數量",
        "單位", "單價", "外幣單價", "產品備註", "備註1",
        "備註2", "備註3", "預定進貨日", "廠商訂單", "自訂櫃號"
    )
    $purchaseWorksheet.Range("A2:O65536").ClearContents() | Out-Null
    $purchaseWorksheet.Columns.Item(4).NumberFormat = "@"

    for ($index = 0; $index -lt $items.Count; $index++) {
        $row = $index + 2
        $item = $items[$index]

        $purchaseWorksheet.Cells.Item($row, 1).Value2 = [string]$dateText
        $purchaseWorksheet.Cells.Item($row, 2).Value2 = [string]$VendorCode
        $purchaseWorksheet.Cells.Item($row, 4).Value2 = [string]$item.productCode
        $purchaseWorksheet.Cells.Item($row, 5).Value2 = [double]$item.quantity
        $purchaseWorksheet.Cells.Item($row, 6).Value2 = [string]"pcs"
        $purchaseWorksheet.Cells.Item($row, 7).Value2 = [double]$item.cost
        $purchaseWorksheet.Cells.Item($row, 10).Value2 = [string]$Note1
    }
    Remove-TrailingBlankRows -Worksheet $purchaseWorksheet -LastDataRow ($items.Count + 1) -LastColumn "O"
    $purchaseWorkbook.Save()
    $purchaseWorkbook.Close($true)
    Release-ComObjects @($purchaseWorksheet, $purchaseWorkbook)
    $purchaseWorksheet = $null
    $purchaseWorkbook = $null
}
finally {
    if ($null -ne $newProductWorkbook) {
        $newProductWorkbook.Close($false)
    }
    if ($null -ne $purchaseWorkbook) {
        $purchaseWorkbook.Close($false)
    }
    if ($null -ne $excel) {
        $excel.Quit()
    }
    Release-ComObjects @(
        $newProductWorksheet,
        $purchaseWorksheet,
        $newProductWorkbook,
        $purchaseWorkbook,
        $excel
    )
}

[PSCustomObject]@{
    newProductFile = if ($newItems.Count -gt 0) { $newProductOutput } else { $null }
    purchaseImportFile = $purchaseOutput
    purchaseDate = $dateText
    vendorCode = $VendorCode
    rowCount = $items.Count
    newProductRowCount = $newItems.Count
    existingProductRowCount = ($items.Count - $newItems.Count)
    taxAdjustedProductCodes = $taxAdjustedCodes
} | ConvertTo-Json
