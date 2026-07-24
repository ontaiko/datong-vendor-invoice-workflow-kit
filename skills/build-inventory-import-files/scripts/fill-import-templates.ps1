param(
    [string]$WorkspaceRoot = "",

    [string[]]$ProductsXlsx = @(),

    [string]$VendorCode = "",

    [string]$OutputDir = "",

    [string]$PurchaseDate = "",

    [string]$Note1 = "阿榮代打",

    [string]$NewProductTemplate = "",

    [string]$PurchaseTemplate = "",

    [string]$InvoiceTotal = "",

    [string]$VendorShortName = "",

    [switch]$ConfirmedReviewed,

    [switch]$RunRegressionTests
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

function Assert-AlwaysTaxInclusiveInvoiceTotal {
    param(
        [object[]]$Items,
        [string]$PrintedInvoiceTotal,
        [string]$VendorName
    )

    if ([string]::IsNullOrWhiteSpace($PrintedInvoiceTotal)) {
        return
    }

    $taxInclusiveTotal = 0.0
    foreach ($item in $Items) {
        $taxInclusiveTotal += ([double]$item.quantity * [double]$item.cost)
    }
    if (-not (Test-PrintedAmountMatch -Calculated $taxInclusiveTotal -Printed ([double]$PrintedInvoiceTotal))) {
        throw "Invoice total mismatch after applying the fixed 1.05 tax rule for vendor '$VendorName'."
    }
}

function Invoke-TaxRuleRegressionTests {
    $items = @(
        [pscustomobject]@{ productCode = "TEST01"; quantity = 1; cost = 32250.0 }
    )

    $adjusted = @(Apply-AlwaysTaxInclusiveCosts -Items $items -AlreadyAdjustedCodes @())
    if ($adjusted.Count -ne 1 -or [Math]::Abs(([double]$items[0].cost) - 33862.5) -gt 0.000001) {
        throw "Regression failed: fixed 1.05 tax adjustment was not applied correctly."
    }

    Assert-AlwaysTaxInclusiveInvoiceTotal -Items $items -PrintedInvoiceTotal "33863" -VendorName "南波"

    $mismatchRejected = $false
    try {
        Assert-AlwaysTaxInclusiveInvoiceTotal -Items $items -PrintedInvoiceTotal "33864" -VendorName "南波"
    } catch {
        $mismatchRejected = $true
    }
    if (-not $mismatchRejected) {
        throw "Regression failed: a one-dollar invoice mismatch was not rejected."
    }

    $fractionalCostItems = @(
        [pscustomobject]@{
            productCode = "TEST02"
            quantity = 20
            cost = 87.5
            lineAmount = 1750
        }
    )
    $fractionalAdjusted = @(Resolve-TaxInclusiveCosts -Items $fractionalCostItems -PrintedInvoiceTotal "1750")
    if ($fractionalAdjusted.Count -ne 0 -or [Math]::Abs(([double]$fractionalCostItems[0].cost) - 87.5) -gt 0.000001) {
        throw "Regression failed: fractional unit cost was rounded or tax-adjusted unexpectedly."
    }

    [pscustomobject]@{
        passed = 4
        taxAdjustedTotal = $items[0].cost
        roundedPrintedTotal = 33863
        mismatchRejected = $mismatchRejected
        fractionalCostPreserved = $fractionalCostItems[0].cost
    } | ConvertTo-Json
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

            # Use the underlying numeric value instead of formatted Text. A cell
            # displayed with zero decimals can show 87.5 as 88 and corrupt cost checks.
            $quantityValue = $worksheet.Cells.Item($row, $columns.quantity).Value2
            $costValue = $worksheet.Cells.Item($row, $columns.cost).Value2
            if ($null -eq $quantityValue -and $null -eq $costValue) {
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
                recommendedPrice = if ($columns.recommendedPrice -gt 0) { $worksheet.Cells.Item($row, $columns.recommendedPrice).Value2 } else { "" }
                quantity = $quantityValue
                cost = $costValue
            lineAmount = if ($columns.lineAmount -gt 0) { $worksheet.Cells.Item($row, $columns.lineAmount).Value2 } else { "" }
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

function Get-OutputMutexName {
    param(
        [string]$Directory,
        [string]$VendorShortName,
        [string]$FileDate
    )

    $identity = ("{0}|{1}|{2}" -f $Directory, $VendorShortName, $FileDate).ToLowerInvariant()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($identity)
    $hashBytes = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    $hash = [System.BitConverter]::ToString($hashBytes).Replace("-", "")
    return "Global\DatongInventoryImport_$hash"
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

if ($RunRegressionTests) {
    Invoke-TaxRuleRegressionTests
    return
}

foreach ($required in @{
    WorkspaceRoot = $WorkspaceRoot
    ProductsXlsx = $ProductsXlsx
    OutputDir = $OutputDir
}.GetEnumerator()) {
    if ([string]::IsNullOrWhiteSpace([string]$required.Value)) {
        throw "Missing required parameter '$($required.Key)'."
    }
}

$workspace = (Resolve-Path -LiteralPath $WorkspaceRoot).Path
$productPathInputs = @()
foreach ($rawProductPath in $ProductsXlsx) {
    foreach ($part in ([string]$rawProductPath -split ",")) {
        $cleanPath = $part.Trim().Trim('"').Trim("'")
        if (-not [string]::IsNullOrWhiteSpace($cleanPath)) {
            $productPathInputs += $cleanPath
        }
    }
}
$productPaths = @($productPathInputs | ForEach-Object {
    (Resolve-Path -LiteralPath $_).Path
})
if ($productPaths.Count -eq 0) {
    throw "Missing required parameter 'ProductsXlsx'."
}

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
$items = @()
foreach ($productPath in $productPaths) {
    $items += @(Read-ProductsFromXlsx -Path $productPath)
}

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

$resolvedVendorShortName = ""
if (-not [string]::IsNullOrWhiteSpace($VendorShortName)) {
    $resolvedVendorShortName = $VendorShortName
} else {
    $vendorNames = @($productPaths | ForEach-Object { Get-VendorShortNameFromXlsx -Path $_ -Fallback "" } | Select-Object -Unique)
    if ($vendorNames.Count -eq 1) {
        $resolvedVendorShortName = $vendorNames[0]
    } else {
        throw "多份 ProductsXlsx 含不同廠商：$($vendorNames -join ', ')。採購單匯入需依廠商分開產生，請分批執行或明確指定同一廠商資料。"
    }
}
if ([string]::IsNullOrWhiteSpace($VendorCode)) {
    if ($resolvedVendorShortName -match "萬榮") {
        $VendorCode = "38"
    } else {
        $vendorWorkbookPath = Get-VendorCodeWorkbookPath -WorkspaceRoot $workspace
        $VendorCode = Resolve-VendorCodeFromWorkbook -WorkbookPath $vendorWorkbookPath -VendorName $resolvedVendorShortName
    }
}
$alwaysTaxInclusiveVendor = Test-AlwaysTaxInclusiveVendor -VendorName $resolvedVendorShortName
$invoiceTotalForInitialCheck = if ($alwaysTaxInclusiveVendor) { "" } else { $InvoiceTotal }
$taxAdjustedCodes = @(Resolve-TaxInclusiveCosts -Items @($items) -PrintedInvoiceTotal $invoiceTotalForInitialCheck)
if ($alwaysTaxInclusiveVendor) {
    $forcedTaxAdjustedCodes = @(Apply-AlwaysTaxInclusiveCosts -Items @($items) -AlreadyAdjustedCodes $taxAdjustedCodes)
    $taxAdjustedCodes = @($taxAdjustedCodes + $forcedTaxAdjustedCodes | Select-Object -Unique)

    Assert-AlwaysTaxInclusiveInvoiceTotal -Items @($items) -PrintedInvoiceTotal $InvoiceTotal -VendorName $resolvedVendorShortName
}
$newItems = @($items | Where-Object { -not ($_.existingProduct -eq $true) })

$dateText = Get-RocDate -Value $PurchaseDate
$fileDate = $dateText.Replace(".", "")
$outputMutexName = Get-OutputMutexName -Directory $outputPath -VendorShortName $resolvedVendorShortName -FileDate $fileDate
$outputMutex = New-Object System.Threading.Mutex($false, $outputMutexName)
$outputMutexAcquired = $false
try {
    $outputMutexAcquired = $outputMutex.WaitOne([TimeSpan]::FromMinutes(2))
    if (-not $outputMutexAcquired) {
        throw "Timed out waiting for output filename lock '$outputMutexName'."
    }

    $outputPair = Get-OutputPairPaths -Directory $outputPath -VendorShortName $resolvedVendorShortName -FileDate $fileDate -IncludeNewProduct ($newItems.Count -gt 0)
    $newProductOutput = $outputPair.newProduct
    $purchaseOutput = $outputPair.purchase

    if ($newItems.Count -gt 0) {
        Copy-Item -LiteralPath $newProductTemplatePath -Destination $newProductOutput
    }
    Copy-Item -LiteralPath $purchaseTemplatePath -Destination $purchaseOutput
}
finally {
    if ($outputMutexAcquired) {
        $outputMutex.ReleaseMutex()
    }
    $outputMutex.Dispose()
}

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
    $fullPurchaseHeaders = @(
        "採購日期", "廠商代號", "外幣幣別", "產品代號", "數量",
        "單位", "單價", "外幣單價", "產品備註", "備註1",
        "備註2", "備註3", "預定進貨日", "廠商訂單", "自訂櫃號"
    )
    $compactPurchaseHeaders = @("採購日期", "廠商代號", "產品代號", "數量", "備註1")
    $isCompactPurchaseTemplate = $true
    for ($column = 1; $column -le $compactPurchaseHeaders.Count; $column++) {
        $actual = [string]$purchaseWorksheet.Cells.Item(1, $column).Text
        if ($actual -ne $compactPurchaseHeaders[$column - 1]) {
            $isCompactPurchaseTemplate = $false
            break
        }
    }

    if ($isCompactPurchaseTemplate) {
        $purchaseWorksheet.Range("A2:E65536").ClearContents() | Out-Null
        $purchaseWorksheet.Columns.Item(3).NumberFormat = "@"
    } else {
        Assert-Headers -Worksheet $purchaseWorksheet -Expected $fullPurchaseHeaders
        $purchaseWorksheet.Range("A2:O65536").ClearContents() | Out-Null
        $purchaseWorksheet.Columns.Item(4).NumberFormat = "@"
    }

    for ($index = 0; $index -lt $items.Count; $index++) {
        $row = $index + 2
        $item = $items[$index]

        $purchaseWorksheet.Cells.Item($row, 1).Value2 = [string]$dateText
        $purchaseWorksheet.Cells.Item($row, 2).Value2 = [string]$VendorCode
        if ($isCompactPurchaseTemplate) {
            $purchaseWorksheet.Cells.Item($row, 3).Value2 = [string]$item.productCode
            $purchaseWorksheet.Cells.Item($row, 4).Value2 = [double]$item.quantity
            $purchaseWorksheet.Cells.Item($row, 5).Value2 = [string]$Note1
        } else {
            $purchaseWorksheet.Cells.Item($row, 4).Value2 = [string]$item.productCode
            $purchaseWorksheet.Cells.Item($row, 5).Value2 = [double]$item.quantity
            $purchaseWorksheet.Cells.Item($row, 6).Value2 = [string]"pcs"
            $purchaseWorksheet.Cells.Item($row, 7).Value2 = [double]$item.cost
            $purchaseWorksheet.Cells.Item($row, 10).Value2 = [string]$Note1
        }
    }
    $purchaseLastColumn = if ($isCompactPurchaseTemplate) { "E" } else { "O" }
    Remove-TrailingBlankRows -Worksheet $purchaseWorksheet -LastDataRow ($items.Count + 1) -LastColumn $purchaseLastColumn
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
    sourceFileCount = $productPaths.Count
    newProductRowCount = $newItems.Count
    existingProductRowCount = ($items.Count - $newItems.Count)
    taxAdjustedProductCodes = $taxAdjustedCodes
} | ConvertTo-Json
