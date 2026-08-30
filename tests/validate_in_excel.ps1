param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [string]$CopyTo,

    [string]$SheetName = "DynamicArrays",

    [int]$ExpectedXlookup = 450,

    [string]$ExpectedMarker = "second openpyxl edit after Excel save"
)

$ErrorActionPreference = "Stop"
$excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$oldAlerts = $excel.DisplayAlerts
$book = $null

try {
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Open($WorkbookPath, 0, $false)
    $sheet = $book.Worksheets.Item($SheetName)

    $formulaD2 = [string]$sheet.Range("D2").Formula2
    $formulaF2 = [string]$sheet.Range("F2").Formula2
    $formulaH2 = [string]$sheet.Range("H2").Formula2
    $formulaJ2 = [string]$sheet.Range("J2").Formula2

    $unique = @(
        [string]$sheet.Range("D2").Value2,
        [string]$sheet.Range("D3").Value2,
        [string]$sheet.Range("D4").Value2
    )
    $vstack = @(
        [string]$sheet.Range("F2").Value2,
        [string]$sheet.Range("F3").Value2,
        [string]$sheet.Range("F4").Value2,
        [string]$sheet.Range("F5").Value2,
        [string]$sheet.Range("F6").Value2,
        [string]$sheet.Range("F7").Value2
    )
    $xlookupScalar = [string]$sheet.Range("H2").Value2
    $xlookupSpill = @(
        [string]$sheet.Range("J2").Value2,
        [string]$sheet.Range("J3").Value2,
        [string]$sheet.Range("J4").Value2
    )
    $marker = [string]$sheet.Range("C3").Value2
    $modifiedInput = [string]$sheet.Range("B5").Value2

    $errors = @()
    if (($formulaD2 + $formulaF2 + $formulaH2 + $formulaJ2).Contains("@")) {
        $errors += "Formula2 contains @"
    }
    if (($unique -join "|") -ne "apple|banana|pear") {
        $errors += "UNIQUE spill mismatch"
    }
    if (($vstack -join "|") -ne "10|20|30|10|20|30") {
        $errors += "VSTACK spill mismatch"
    }
    if ($xlookupScalar -ne [string]$ExpectedXlookup) {
        $errors += "XLOOKUP scalar mismatch"
    }
    if (($xlookupSpill -join "|") -ne "10|20|$ExpectedXlookup") {
        $errors += "XLOOKUP spill mismatch"
    }
    if ($marker -ne $ExpectedMarker) {
        $errors += "openpyxl marker mismatch"
    }
    if ($modifiedInput -ne [string]$ExpectedXlookup) {
        $errors += "modified input mismatch"
    }

    $book.Save()
    $book.Close($false)
    $book = $null

    if ($CopyTo) {
        Copy-Item -LiteralPath $WorkbookPath -Destination $CopyTo -Force
    }

    [pscustomobject]@{
        Status = if ($errors.Count -eq 0) { "PASS" } else { "FAIL" }
        ExcelVersion = [string]$excel.Version
        Formulas = @($formulaD2, $formulaF2, $formulaH2, $formulaJ2)
        Unique = $unique
        Vstack = $vstack
        XlookupScalar = $xlookupScalar
        XlookupSpill = $xlookupSpill
        Marker = $marker
        ModifiedInput = $modifiedInput
        SavedAndClosed = $true
        CopiedTo = $CopyTo
        Errors = $errors
    } | ConvertTo-Json -Depth 4 -Compress
}
finally {
    if ($null -ne $book) {
        $book.Close($false)
    }
    $excel.DisplayAlerts = $oldAlerts
}
