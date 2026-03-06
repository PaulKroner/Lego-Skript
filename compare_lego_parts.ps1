param(
    [string]$SetAPath = "",
    [string]$SetBPath = "",
    [string]$Output = "lego_parts_match.xlsx"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ItemIdKeys = @("itemid", "item_id", "designid", "itemno", "item", "partid", "part_id", "partno")
$QtyKeys   = @("qty", "quantity", "count", "amount", "q", "quantityneed", "minqty", "qtyowned")
$ColorKeys = @("color", "colorid", "color_id", "colour", "colourid", "colour_id")
$NameKeys  = @("name", "itemname", "partname", "description", "desc")

function Normalize-Tag {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    $clean = $Value.ToLowerInvariant()
    if ($clean.Contains(":")) {
        $clean = $clean.Substring($clean.LastIndexOf(":") + 1)
    }
    return ($clean -replace "[^a-z0-9]", "")
}

function Parse-Int {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return 0 }
    $m = [regex]::Match($Value, "-?\d+")
    if (-not $m.Success) { return 0 }
    return [int]$m.Value
}

function Try-GetValue {
    param(
        [System.Xml.XmlElement]$Node,
        [string[]]$Keys
    )

    $keySet = @{}
    foreach ($k in $Keys) { $keySet[$k] = $true }

    foreach ($attr in $Node.Attributes) {
        if ($keySet.ContainsKey((Normalize-Tag $attr.Name))) {
            if (-not [string]::IsNullOrWhiteSpace($attr.Value)) { return $attr.Value.Trim() }
        }
    }

    foreach ($child in $Node.ChildNodes) {
        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        if ($keySet.ContainsKey((Normalize-Tag $child.Name)) -and -not [string]::IsNullOrWhiteSpace($child.InnerText)) {
            return $child.InnerText.Trim()
        }
    }

    return $null
}

function Parse-XmlParts {
    param([string]$Path)

    [xml]$xml = Get-Content -Path $Path -Raw
    if ($null -eq $xml.DocumentElement) {
        throw "Ungültiges XML: $Path"
    }

    $counts = @{}
    $names = @{}
    $colors = @{}
    $totalPieces = 0

    $nodes = $xml.SelectNodes(".//*")
    $itemKeys  = $script:ItemIdKeys | ForEach-Object { Normalize-Tag $_ }
    $qtyKeys   = $script:QtyKeys | ForEach-Object { Normalize-Tag $_ }
    $colorKeys = $script:ColorKeys | ForEach-Object { Normalize-Tag $_ }
    $nameKeys  = $script:NameKeys | ForEach-Object { Normalize-Tag $_ }

    foreach ($node in $nodes) {
        $itemId = Try-GetValue -Node $node -Keys $itemKeys
        $qtyRaw = Try-GetValue -Node $node -Keys $qtyKeys
        if ([string]::IsNullOrWhiteSpace($itemId) -or [string]::IsNullOrWhiteSpace($qtyRaw)) {
            continue
        }

        $qty = Parse-Int $qtyRaw
        if ($qty -le 0) { continue }

        $itemIdNorm = $itemId.Trim()
        if ([string]::IsNullOrWhiteSpace($itemIdNorm)) { continue }

        $colorRaw = Try-GetValue -Node $node -Keys $colorKeys
        if ([string]::IsNullOrWhiteSpace($colorRaw)) { $colorNorm = "" } else { $colorNorm = $colorRaw.Trim() }
        $key = "{0}|{1}" -f $itemIdNorm, $colorNorm

        if (-not $counts.ContainsKey($key)) { $counts[$key] = 0 }
        $counts[$key] += $qty
        $totalPieces += $qty

        if (-not $colors.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace($colorRaw)) {
            $colors[$key] = $colorRaw
        }
        if (-not $names.ContainsKey($key)) {
            $nameRaw = Try-GetValue -Node $node -Keys $nameKeys
            if (-not [string]::IsNullOrWhiteSpace($nameRaw)) {
                $names[$key] = $nameRaw
            }
        }
    }

    [pscustomobject]@{
        Counts = $counts
        Names = $names
        Colors = $colors
        TotalPieces = $totalPieces
    }
}

function Get-DetectedFiles {
    $xmlFiles = Get-ChildItem -Path $PSScriptRoot -Filter "*.xml" -File | Sort-Object Length -Descending
    if ($xmlFiles.Count -lt 2) {
        throw "Mindestens zwei XML-Dateien werden benötigt."
    }
    return $xmlFiles[0].FullName, $xmlFiles[1].FullName
}

function Export-LegoComparisonWorkbook {
    param(
        [string]$WorkbookPath,
        [System.Collections.Generic.List[object]]$Rows,
        [int]$TotalInA,
        [int]$TotalInB,
        [int]$TotalUsable,
        [int]$UnusedB,
        [int]$TotalMissing,
        [double]$CoverageA,
        [double]$CoverageFromB
    )

    $excel = New-Object -ComObject Excel.Application
    try {
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $sheet1 = $workbook.Worksheets.Item(1)
        $sheet1.Name = "Vergleich"

        $headers = @(
            "TeilID", "Farbe", "Menge im A", "Menge in B", "nutzbar", "Restmangel", "in B vorhanden"
        )
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $sheet1.Cells.Item(1, $c + 1).Value2 = $headers[$c]
        }

        $r = 2
        foreach ($row in $Rows) {
            $sheet1.Cells.Item($r, 1).Value2 = $row.TeilID
            $sheet1.Cells.Item($r, 2).Value2 = $row.Farbe
            $sheet1.Cells.Item($r, 3).Value2 = [int]$row.'Menge im A'
            $sheet1.Cells.Item($r, 4).Value2 = [int]$row.'Menge in B'
            $sheet1.Cells.Item($r, 5).Value2 = [int]$row.nutzbar
            $sheet1.Cells.Item($r, 6).Value2 = [int]$row.Restmangel
            $sheet1.Cells.Item($r, 7).Value2 = $row.'in B vorhanden'
            $r++
        }

        $sheet1.UsedRange.EntireColumn.AutoFit() | Out-Null

        # Entferne die Spalten H bis K im Vergleich-Sheet (wie gewünscht)
        $sheet1.Columns("H:K").Delete()

        $sheet2 = $workbook.Worksheets.Add()
        $sheet2.Name = "Summary"
        $sheet2.Cells.Item(1,1).Value2 = "Metrik"
        $sheet2.Cells.Item(1,2).Value2 = "Wert"

        $s = 2
        $sheet2.Cells.Item($s,1).Value2 = "Gesamtteile in A"
        $sheet2.Cells.Item($s,2).Value2 = $TotalInA
        $s++
        $sheet2.Cells.Item($s,1).Value2 = "Gesamtteile in B"
        $sheet2.Cells.Item($s,2).Value2 = $TotalInB
        $s++
        $sheet2.Cells.Item($s,1).Value2 = "Nutzbar aus B für A"
        $sheet2.Cells.Item($s,2).Value2 = ("{0}/{1}" -f $TotalUsable, $TotalInB)
        $s++
        $sheet2.Cells.Item($s,1).Value2 = "Nutzbare Teile aus B gesamt"
        $sheet2.Cells.Item($s,2).Value2 = $TotalUsable
        $s++
        $sheet2.Cells.Item($s,1).Value2 = "Nicht nutzbare Teile aus B"
        $sheet2.Cells.Item($s,2).Value2 = $UnusedB
        $s++
        $sheet2.Cells.Item($s,1).Value2 = "Nutzung von B in A"
        $sheet2.Cells.Item($s,2).Value2 = ("{0:N2}%" -f $CoverageFromB)
        $s++
        $sheet2.Cells.Item($s,1).Value2 = "Nicht im A abgedeckt"
        $sheet2.Cells.Item($s,2).Value2 = $TotalMissing
        $s++
        $sheet2.Cells.Item($s,1).Value2 = "Abdeckung A/B"
        $sheet2.Cells.Item($s,2).Value2 = ("{0:N2}%" -f $CoverageA)
        $sheet2.UsedRange.EntireColumn.AutoFit() | Out-Null

        $workbook.SaveAs($WorkbookPath) | Out-Null
    }
    finally {
        if ($workbook) { $null = $workbook.Close($true) }
        if ($excel) { $excel.Quit() }
        if ($sheet2) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet2) }
        if ($sheet1) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet1) }
        if ($workbook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) }
        if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

if ([string]::IsNullOrWhiteSpace($SetAPath) -or [string]::IsNullOrWhiteSpace($SetBPath)) {
    $detected = Get-DetectedFiles
    if ([string]::IsNullOrWhiteSpace($SetAPath)) { $SetAPath = $detected[0] }
    if ([string]::IsNullOrWhiteSpace($SetBPath)) { $SetBPath = $detected[1] }
}

if (-not (Test-Path -LiteralPath $SetAPath)) { throw "Datei nicht gefunden: $SetAPath" }
if (-not (Test-Path -LiteralPath $SetBPath)) { throw "Datei nicht gefunden: $SetBPath" }

$setA = Parse-XmlParts -Path $SetAPath
$setB = Parse-XmlParts -Path $SetBPath
$totalInA = $setA.TotalPieces
$totalInB = $setB.TotalPieces

$rows = New-Object System.Collections.Generic.List[object]
$totalUsable = 0
$totalMissing = 0

foreach ($entry in $setA.Counts.GetEnumerator() | Sort-Object { $_.Name }) {
    $key = $entry.Name
    $needed = [int]$entry.Value
    $have = if ($setB.Counts.ContainsKey($key)) { [int]$setB.Counts[$key] } else { 0 }

    $split = $key -split "\|", 2
    $itemId = $split[0]
    $color  = if ($split.Count -gt 1) { $split[1] } else { "unbekannt" }
    if ([string]::IsNullOrWhiteSpace($color)) { $color = "unbekannt" }

    $usable = [Math]::Min($needed, $have)
    $missing = [Math]::Max(0, ($needed - $have))
    $present = if ($have -gt 0) { "Ja" } else { "Nein" }

    $totalUsable += $usable
    $totalMissing += $missing

    $rows.Add([pscustomobject]@{
        TeilID = $itemId
        Farbe = $color
        "Menge im A" = $needed
        "Menge in B" = $have
        "nutzbar" = $usable
        "Restmangel" = $missing
        "in B vorhanden" = $present
    })
}

$coverageA = if ($totalInA -gt 0) { [double]$totalUsable / $totalInA * 100 } else { 0.0 }
$coverageFromB = if ($totalInB -gt 0) { [double]$totalUsable / $totalInB * 100 } else { 0.0 }
$unusedB = $totalInB - $totalUsable

$outputFile = if ([System.IO.Path]::IsPathRooted($Output)) { $Output } else { Join-Path -Path $PSScriptRoot -ChildPath $Output }

if ([System.IO.Path]::GetExtension($outputFile).ToLowerInvariant() -eq ".xlsx") {
    Export-LegoComparisonWorkbook -WorkbookPath $outputFile -Rows $rows -TotalInA $totalInA -TotalInB $totalInB -TotalUsable $totalUsable -UnusedB $unusedB -TotalMissing $totalMissing -CoverageA $coverageA -CoverageFromB $coverageFromB
    Write-Host "Excel-Ausgabe: $outputFile"
} else {
    throw "Nur .xlsx-Ausgabe wird unterstützt, da du ein zweites Sheet im gleichen File gefordert hast."
}

Write-Host ("Untersuchte Teilkataloge: {0} Einträge" -f $setA.Counts.Count)
Write-Host ("Benötigte Teile in A: {0}" -f $totalInA)
Write-Host ("Benötigte Teile in B: {0}" -f $totalInB)
Write-Host ("Nutzbar aus B für A: {0}/{1}" -f $totalUsable, $totalInB)
Write-Host ("Nutzbare Teile aus B: {0}" -f $totalUsable)
Write-Host ("Nicht nutzbare Teile aus B: {0}" -f $unusedB)
Write-Host ("Nutzung von B in A: {0:N2}%" -f $coverageFromB)
Write-Host ("Abdeckung von A: {0:N2}%" -f $coverageA)
Write-Host ("Nicht im A abgedeckt: {0}" -f $totalMissing)
