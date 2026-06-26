param(
    [Parameter(Mandatory = $true)]
    [string]$WorkspaceRoot,

    [Parameter(Mandatory = $true)]
    [string]$ProductsJson,

    [Parameter(Mandatory = $true)]
    [string]$VendorCode,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [string]$PurchaseDate = "",

    [string]$Note1 = "阿榮代打",

    [string]$NewProductTemplate = "",

    [string]$PurchaseTemplate = "",

    [string]$InvoiceTotal = ""
)

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

$workspace = (Resolve-Path -LiteralPath $WorkspaceRoot).Path
$productsPath = (Resolve-Path -LiteralPath $ProductsJson).Path
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
$items = Get-Content -LiteralPath $productsPath -Raw -Encoding UTF8 | ConvertFrom-Json

if ($items.Count -eq 0) {
    throw "Products JSON has no rows."
}

foreach ($item in $items) {
    foreach ($property in @("productCode", "name", "cost", "quantity")) {
        if ($null -eq $item.$property -or [string]::IsNullOrWhiteSpace([string]$item.$property)) {
            throw "Missing required property '$property' in products JSON."
        }
    }

    if (-not ($item.existingProduct -eq $true)) {
        if ($null -eq $item.category -or [string]::IsNullOrWhiteSpace([string]$item.category)) {
            throw "Missing required property 'category' for new product '$($item.productCode)' in products JSON."
        }
    }
}

$taxAdjustedCodes = @(Resolve-TaxInclusiveCosts -Items @($items) -PrintedInvoiceTotal $InvoiceTotal)
$newItems = @($items | Where-Object { -not ($_.existingProduct -eq $true) })

$dateText = Get-RocDate -Value $PurchaseDate
$fileDate = $dateText.Replace(".", "")
$newProductOutput = Join-Path $outputPath "建檔用_進貨_$fileDate.xls"
$purchaseOutput = Join-Path $outputPath "採購單匯入_進貨_$fileDate.xls"

if ($newItems.Count -gt 0) {
    Copy-Item -LiteralPath $newProductTemplatePath -Destination $newProductOutput -Force
}
Copy-Item -LiteralPath $purchaseTemplatePath -Destination $purchaseOutput -Force

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
        $newProductWorksheet = $newProductWorkbook.Worksheets.Item("工作表1")
        Assert-Headers -Worksheet $newProductWorksheet -Expected @(
            "產品代號", "產品名稱", "銷售單價1", "銷售單價2", "建議售價",
            "單位", "成本", "CO128", "大類"
        )
        $newProductWorksheet.Range("A2:I65536").ClearContents() | Out-Null

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
            $newProductWorksheet.Cells.Item($row, 9).Value2 = [double]$item.category
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
