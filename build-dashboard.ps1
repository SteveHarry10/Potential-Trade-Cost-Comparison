param(
  [string]$WorkbookPath = "work\Floorplans Costs.xlsx",
  [string]$OutputPath = "outputs\floorplans-cost-dashboard.html",
  [string]$SheetName = "",
  [int]$PlanNameRow = 4,
  [int]$PlanStartColumn = 4,
  [int]$CostCodeColumn = 2,
  [int]$CostItemColumn = 3,
  [int]$GroupRow = 1,
  [int]$PlanCodeRow = 2,
  [int]$PlanImportNameRow = 3,
  [int]$MainSqftRow = 6,
  [int]$SecondSqftRow = 7,
  [int]$ThirdSqftRow = 8,
  [int]$BasementSqftRow = 9,
  [int]$CostStartRow = 12,
  [string]$DataOutputPath = "",
  [string]$ExternalDataPath = ""
)

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Convert-ColumnNameToIndex {
  param([string]$Name)
  $index = 0
  foreach ($char in $Name.ToCharArray()) {
    $index = ($index * 26) + ([int][char]$char - [int][char]'A' + 1)
  }
  return $index
}

function Get-CellParts {
  param([string]$Ref)
  if ($Ref -match '^([A-Z]+)([0-9]+)$') {
    return @{
      Column = Convert-ColumnNameToIndex $matches[1]
      Row = [int]$matches[2]
    }
  }
  return $null
}

function Get-RangeParts {
  param([string]$Ref)
  $bounds = $Ref -split ':'
  if ($bounds.Count -eq 1) {
    $start = Get-CellParts $bounds[0]
    return @{
      Start = $start
      End = $start
    }
  }
  return @{
    Start = Get-CellParts $bounds[0]
    End = Get-CellParts $bounds[1]
  }
}

function Get-EntryText {
  param($Zip, [string]$Name)
  $entry = $Zip.GetEntry($Name)
  if (-not $entry) { return $null }
  $stream = $entry.Open()
  try {
    $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() }
    finally { $reader.Dispose() }
  }
  finally {
    $stream.Dispose()
  }
}

function Get-Xml {
  param($Zip, [string]$Name)
  $text = Get-EntryText $Zip $Name
  if (-not $text) { return $null }
  return [xml]$text
}

function Get-CellValue {
  param($Cell, [string[]]$SharedStrings)

  $type = $Cell.t
  if ($type -eq 's') {
    if ($null -eq $Cell.v -or [string]$Cell.v -eq '') { return $null }
    return $SharedStrings[[int]$Cell.v]
  }
  if ($type -eq 'inlineStr') {
    $nodes = $Cell.is.GetElementsByTagName('t')
    return (($nodes | ForEach-Object { $_.'#text' }) -join '')
  }
  if ($type -eq 'str' -or $type -eq 'e') {
    return [string]$Cell.v
  }
  if ($type -eq 'b') {
    return ([string]$Cell.v) -eq '1'
  }
  if ($null -eq $Cell.v -or [string]$Cell.v -eq '') { return $null }

  $numeric = 0.0
  if ([double]::TryParse([string]$Cell.v, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$numeric)) {
    return $numeric
  }
  return [string]$Cell.v
}

function Get-Value {
  param($Matrix, [int]$Row, [int]$Column)
  if ($Matrix.ContainsKey($Row) -and $Matrix[$Row].ContainsKey($Column)) {
    return $Matrix[$Row][$Column]
  }
  return $null
}

function Get-Text {
  param($Value)
  if ($null -eq $Value) { return "" }
  return ([string]$Value).Trim()
}

function Test-Numeric {
  param($Value)
  return ($Value -is [int] -or $Value -is [long] -or $Value -is [float] -or $Value -is [double] -or $Value -is [decimal])
}

function Get-NumberOrNull {
  param($Value)
  if (Test-Numeric $Value) { return [double]$Value }
  return $null
}

function Get-PlanId {
  param([string]$Name, [int]$Column)
  $slug = $Name.ToLowerInvariant() -replace '[^a-z0-9]+', '-'
  $slug = $slug.Trim('-')
  if (-not $slug) { $slug = "plan" }
  return "$slug-$Column"
}

function Get-Category {
  param([string]$Item)
  if ($Item -match '^\s*([^-]+?)\s+-\s+(.+)$') {
    return $matches[1].Trim()
  }
  return "Other"
}

$resolvedWorkbook = Resolve-Path $WorkbookPath
$zip = [System.IO.Compression.ZipFile]::OpenRead($resolvedWorkbook)
try {
  $sharedStrings = @()
  $sharedXml = Get-Xml $zip 'xl/sharedStrings.xml'
  if ($sharedXml) {
    foreach ($si in $sharedXml.sst.si) {
      $texts = $si.GetElementsByTagName('t') | ForEach-Object { $_.'#text' }
      $sharedStrings += ($texts -join '')
    }
  }

  $relsXml = Get-Xml $zip 'xl/_rels/workbook.xml.rels'
  $rels = @{}
  foreach ($rel in $relsXml.Relationships.Relationship) {
    $rels[$rel.Id] = $rel.Target
  }

  $workbookXml = Get-Xml $zip 'xl/workbook.xml'
  $availableSheets = @($workbookXml.workbook.sheets.sheet)
  if ($SheetName) {
    $sheet = $availableSheets | Where-Object { [string]$_.name -eq $SheetName } | Select-Object -First 1
    if (-not $sheet) {
      $names = ($availableSheets | ForEach-Object { [string]$_.name }) -join ', '
      throw "Sheet '$SheetName' was not found in '$resolvedWorkbook'. Available sheets: $names"
    }
  }
  else {
    $sheet = $availableSheets | Select-Object -First 1
  }
  $rid = $sheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $sheetPath = $rels[$rid]
  if ($sheetPath -and -not $sheetPath.StartsWith('xl/')) {
    $sheetPath = 'xl/' + $sheetPath
  }

  $sheetXml = Get-Xml $zip $sheetPath
  $matrix = @{}
  $maxColumn = 0
  $maxRow = 0

  foreach ($row in $sheetXml.worksheet.sheetData.row) {
    foreach ($cell in $row.c) {
      $parts = Get-CellParts $cell.r
      if (-not $parts) { continue }
      $value = Get-CellValue $cell $sharedStrings
      if ($null -eq $value) { continue }
      if (-not $matrix.ContainsKey($parts.Row)) {
        $matrix[$parts.Row] = @{}
      }
      $matrix[$parts.Row][$parts.Column] = $value
      if ($parts.Column -gt $maxColumn) { $maxColumn = $parts.Column }
      if ($parts.Row -gt $maxRow) { $maxRow = $parts.Row }
    }
  }

  $rowOneMergedLabels = @{}
  if ($sheetXml.worksheet.mergeCells) {
    foreach ($mergeCell in $sheetXml.worksheet.mergeCells.mergeCell) {
      $range = Get-RangeParts $mergeCell.ref
      if (-not $range.Start -or -not $range.End) { continue }
      if ($range.Start.Row -le $GroupRow -and $range.End.Row -ge $GroupRow) {
        $label = Get-Text (Get-Value $matrix $range.Start.Row $range.Start.Column)
        if ($label) {
          for ($col = $range.Start.Column; $col -le $range.End.Column; $col++) {
            $rowOneMergedLabels[$col] = $label
          }
        }
      }
    }
  }

  $planColumns = @()
  for ($col = $PlanStartColumn; $col -le $maxColumn; $col++) {
    $name = Get-Text (Get-Value $matrix $PlanNameRow $col)
    if ($name) {
      $planColumns += $col
    }
  }

  $costRows = @()
  for ($rowIndex = $CostStartRow; $rowIndex -le $maxRow; $rowIndex++) {
    $code = Get-Text (Get-Value $matrix $rowIndex $CostCodeColumn)
    $item = Get-Text (Get-Value $matrix $rowIndex $CostItemColumn)
    if ($code -and $item) {
      $costRows += [ordered]@{
        id = "$rowIndex|$code"
        row = $rowIndex
        code = $code
        item = $item
        category = Get-Category $item
      }
    }
  }

  $plans = @()
  foreach ($col in $planColumns) {
    $rawGroup = Get-Text (Get-Value $matrix $GroupRow $col)
    if (-not $rawGroup -and $rowOneMergedLabels.ContainsKey($col)) {
      $rawGroup = $rowOneMergedLabels[$col]
    }

    $name = Get-Text (Get-Value $matrix $PlanNameRow $col)
    $mainSqft = Get-NumberOrNull (Get-Value $matrix $MainSqftRow $col)
    $secondSqft = Get-NumberOrNull (Get-Value $matrix $SecondSqftRow $col)
    $thirdSqft = Get-NumberOrNull (Get-Value $matrix $ThirdSqftRow $col)
    $basementSqft = Get-NumberOrNull (Get-Value $matrix $BasementSqftRow $col)

    $values = @()
    $totalCost = 0.0
    $missingCount = 0
    $categoryTotals = [ordered]@{}

    foreach ($costRow in $costRows) {
      $raw = Get-Value $matrix $costRow.row $col
      $value = Get-NumberOrNull $raw
      if ($null -eq $value) {
        $values += $null
        $missingCount += 1
      }
      else {
        $values += $value
        $totalCost += $value
        if (-not $categoryTotals.Contains($costRow.category)) {
          $categoryTotals[$costRow.category] = 0.0
        }
        $categoryTotals[$costRow.category] = [double]$categoryTotals[$costRow.category] + $value
      }
    }

   $totalSqft = 0.0
foreach ($sqft in @($mainSqft, $secondSqft, $thirdSqft)) {
  if ($null -ne $sqft) { $totalSqft += $sqft }
}

    $plans += [ordered]@{
      id = Get-PlanId $name $col
      column = $col
      name = $name
      code = Get-Text (Get-Value $matrix $PlanCodeRow $col)
      importName = Get-Text (Get-Value $matrix $PlanImportNameRow $col)
      group = $rawGroup
      sqft = [ordered]@{
        main = $mainSqft
        second = $secondSqft
        third = $thirdSqft
        basement = $basementSqft
        total = $totalSqft
      }
      totalCost = $totalCost
      costPerSqft = $(if ($totalSqft -gt 0) { $totalCost / $totalSqft } else { $null })
      missingCount = $missingCount
      values = $values
      categoryTotals = $categoryTotals
    }
  }

  $data = [ordered]@{
    source = [ordered]@{
      workbook = Split-Path $resolvedWorkbook -Leaf
      sheet = [string]$sheet.name
      importedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
      planCount = $plans.Count
      costLineCount = $costRows.Count
      planNameRow = $PlanNameRow
      costStartRow = $CostStartRow
      mainSqftRow = $MainSqftRow
      secondSqftRow = $SecondSqftRow
      thirdSqftRow = $ThirdSqftRow
      basementSqftRow = $BasementSqftRow
    }
    items = $costRows
    plans = $plans
  }

  $json = ($data | ConvertTo-Json -Depth 100 -Compress).Replace('</script>', '<\/script>')

  $html = @'
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Arive Homes Floorplan Cost Comparison</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f7f8f4;
      --surface: #ffffff;
      --surface-strong: #f2f5ee;
      --line: #dbe2d5;
      --line-strong: #b8c4b0;
      --text: #303436;
      --muted: #68716a;
      --accent: #7ac143;
      --accent-dark: #548a2d;
      --accent-soft: #eef7e7;
      --charcoal: #5d6264;
      --blue: #667276;
      --amber: #8a6d3b;
      --red: #b42318;
      --green: #067647;
      --shadow: 0 18px 45px rgba(48, 52, 54, 0.09);
      font-family: "Segoe UI", Roboto, Arial, sans-serif;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--text);
    }

    button,
    input,
    select {
      font: inherit;
    }

    button {
      border: 1px solid var(--line-strong);
      background: var(--surface);
      color: var(--text);
      border-radius: 7px;
      min-height: 36px;
      padding: 7px 10px;
      cursor: pointer;
    }

    button:hover,
    button:focus-visible {
      border-color: var(--accent);
      outline: 2px solid transparent;
    }

    .wide-button {
      width: 100%;
      min-height: 42px;
      border-color: var(--accent-dark);
      background: var(--accent);
      color: #ffffff;
      font-weight: 800;
    }

    .wide-button:hover,
    .wide-button:focus-visible {
      border-color: #3f6f20;
      background: var(--accent-dark);
    }

    input[type="search"],
    select {
      width: 100%;
      min-height: 38px;
      border: 1px solid var(--line-strong);
      border-radius: 7px;
      padding: 8px 10px;
      background: var(--surface);
      color: var(--text);
    }

    label {
      display: block;
      color: var(--muted);
      font-size: 12px;
      font-weight: 700;
      letter-spacing: 0;
      margin: 0 0 6px;
      text-transform: uppercase;
    }

    .sr-only {
      position: absolute;
      width: 1px;
      height: 1px;
      padding: 0;
      margin: -1px;
      overflow: hidden;
      clip: rect(0, 0, 0, 0);
      white-space: nowrap;
      border: 0;
    }

    .app {
      width: min(100%, 1580px);
      margin: 0 auto;
      padding: 20px;
    }

    .topbar {
      display: grid;
      grid-template-columns: minmax(260px, 1fr) auto;
      gap: 16px;
      align-items: center;
      padding: 16px 18px;
      border: 1px solid var(--line);
      border-top: 5px solid var(--accent);
      border-radius: 8px;
      background: var(--surface);
      box-shadow: var(--shadow);
    }

    .brand-lockup {
      display: flex;
      align-items: center;
      gap: 14px;
      min-width: 0;
    }

    .brand-mark {
      position: relative;
      width: 44px;
      height: 54px;
      flex: 0 0 44px;
    }

    .brand-mark::before,
    .brand-mark::after {
      content: "";
      position: absolute;
      bottom: 0;
      border-radius: 1px;
      transform-origin: bottom center;
    }

    .brand-mark::before {
      left: 4px;
      width: 14px;
      height: 54px;
      background: var(--accent);
      transform: skew(-24deg);
    }

    .brand-mark::after {
      left: 25px;
      width: 13px;
      height: 42px;
      background: var(--charcoal);
      transform: skew(26deg);
    }

    .wordmark {
      margin: 0 0 2px;
      font-size: 28px;
      line-height: 1;
      font-weight: 800;
      letter-spacing: 0;
      text-transform: lowercase;
    }

    .wordmark-arive {
      color: var(--charcoal);
    }

    .wordmark-homes {
      color: var(--accent);
    }

    .eyebrow {
      margin: 0 0 5px;
      color: var(--accent);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0;
      text-transform: uppercase;
    }

    h1,
    h2,
    h3,
    p {
      margin: 0;
    }

    h1 {
      color: var(--text);
      font-size: clamp(24px, 3vw, 36px);
      line-height: 1.08;
      letter-spacing: 0;
    }

    h2 {
      font-size: 18px;
      letter-spacing: 0;
    }

    h3 {
      font-size: 14px;
      letter-spacing: 0;
    }

    .source-meta {
      display: flex;
      flex-wrap: wrap;
      justify-content: flex-end;
      gap: 8px;
      color: var(--muted);
      font-size: 13px;
    }

    .source-meta span {
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 6px 10px;
      background: var(--accent-soft);
      color: var(--text);
    }

    .topbar-actions {
      display: flex;
      flex-wrap: wrap;
      justify-content: flex-end;
      align-items: center;
      gap: 8px;
    }

    .topbar-button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 36px;
      border: 1px solid var(--accent-dark);
      border-radius: 7px;
      padding: 7px 12px;
      background: var(--accent);
      color: #ffffff;
      font-size: 13px;
      font-weight: 800;
      line-height: 1;
      text-decoration: none;
      white-space: nowrap;
    }

    .topbar-button:hover,
    .topbar-button:focus-visible {
      border-color: #3f6f20;
      background: var(--accent-dark);
      outline: 2px solid transparent;
    }

    .topbar-button[aria-disabled="true"] {
      pointer-events: none;
      border-color: var(--line);
      background: var(--surface-strong);
      color: var(--muted);
    }

    .layout {
      display: grid;
      grid-template-columns: 330px minmax(0, 1fr);
      gap: 18px;
      padding-top: 18px;
      align-items: start;
    }

    .controls {
      position: sticky;
      top: 14px;
      background: var(--surface);
      border: 1px solid var(--line);
      border-top: 5px solid var(--accent);
      border-radius: 8px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }

    .control-section {
      padding: 14px;
      border-bottom: 1px solid var(--line);
    }

    .control-section:last-child {
      border-bottom: 0;
    }

    .control-row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
      margin-top: 10px;
    }

    .select-stack {
      display: grid;
      gap: 12px;
    }

    .toggle-line {
      display: flex;
      gap: 9px;
      align-items: center;
      color: var(--text);
      font-size: 14px;
      margin-top: 10px;
    }

    .toggle-line input {
      width: 18px;
      height: 18px;
      accent-color: var(--accent);
      flex: 0 0 auto;
    }

    .plan-list {
      display: grid;
      gap: 6px;
      max-height: 48vh;
      overflow: auto;
      padding-right: 2px;
      margin-top: 10px;
    }

    .plan-option {
      display: grid;
      grid-template-columns: 22px minmax(0, 1fr) auto;
      gap: 8px;
      align-items: center;
      border: 1px solid var(--line);
      border-radius: 7px;
      padding: 8px;
      background: var(--surface);
    }

    .plan-option input {
      width: 17px;
      height: 17px;
      accent-color: var(--accent);
    }

    .plan-option.is-base {
      border-color: var(--accent);
      background: var(--accent-soft);
    }

    .plan-name {
      min-width: 0;
      font-size: 14px;
      font-weight: 700;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .plan-code {
      color: var(--muted);
      font-size: 12px;
      justify-self: end;
      white-space: nowrap;
    }

    .workspace {
      display: grid;
      gap: 18px;
      min-width: 0;
    }

    .metric-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(160px, 1fr));
      gap: 10px;
    }

    .metric {
      background: var(--surface);
      border: 1px solid var(--line);
      border-top: 4px solid var(--accent);
      border-radius: 8px;
      padding: 13px;
      min-height: 88px;
    }

    .metric-title {
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0;
      text-transform: uppercase;
      margin-bottom: 8px;
    }

    .metric-value {
      font-size: 24px;
      font-weight: 800;
      line-height: 1.08;
      overflow-wrap: anywhere;
    }

    .metric-note {
      color: var(--muted);
      font-size: 12px;
      margin-top: 6px;
    }

    .panel {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      overflow: hidden;
    }

    .panel-header {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 14px;
      border-bottom: 1px solid var(--line);
      background: var(--surface-strong);
    }

    .panel-tools {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
    }

    .panel-tools input,
    .panel-tools select {
      width: min(280px, 72vw);
    }

    .bar-list {
      display: grid;
      gap: 10px;
      padding: 14px;
    }

    .bar-row {
      display: grid;
      grid-template-columns: minmax(120px, 180px) minmax(160px, 1fr) minmax(132px, auto);
      gap: 10px;
      align-items: center;
    }

    .bar-label {
      min-width: 0;
      font-size: 13px;
      font-weight: 700;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .bar-track {
      height: 18px;
      border-radius: 999px;
      background: #e8eee1;
      overflow: hidden;
      border: 1px solid var(--line);
    }

    .bar-fill {
      height: 100%;
      min-width: 3px;
      border-radius: inherit;
      background: var(--blue);
    }

    .bar-fill.is-base {
      background: var(--accent);
    }

    .bar-value {
      text-align: right;
      font-variant-numeric: tabular-nums;
      color: var(--text);
      font-size: 13px;
      white-space: nowrap;
    }

    .table-wrap {
      overflow: auto;
      max-width: 100%;
      scrollbar-gutter: stable;
    }

    table {
      width: 100%;
      border-collapse: collapse;
      min-width: 780px;
      font-size: 13px;
    }

    .comparison-table {
      width: max-content;
      min-width: 100%;
    }

    .comparison-table th,
    .comparison-table td {
      min-width: 155px;
    }

    .comparison-table th:first-child,
    .comparison-table td:first-child {
      min-width: 120px;
    }

    th,
    td {
      border-bottom: 1px solid var(--line);
      padding: 9px 10px;
      text-align: left;
      vertical-align: top;
    }

    th {
      position: sticky;
      top: 0;
      z-index: 1;
      background: var(--surface-strong);
      color: var(--muted);
      font-size: 11px;
      font-weight: 800;
      letter-spacing: 0;
      text-transform: uppercase;
      white-space: nowrap;
    }

    td.numeric,
    th.numeric {
      text-align: right;
      font-variant-numeric: tabular-nums;
      white-space: nowrap;
    }

    .item-cell {
      min-width: 320px;
      max-width: 420px;
    }

    .summary-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(220px, 1fr));
      gap: 12px;
      padding: 14px;
    }

    .summary-highlight {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 13px;
      min-height: 120px;
      background: var(--surface);
    }

    .summary-highlight-title {
      color: var(--muted);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0;
      text-transform: uppercase;
      margin-bottom: 8px;
    }

    .summary-highlight-value {
      font-size: 18px;
      font-weight: 800;
      line-height: 1.2;
      overflow-wrap: anywhere;
    }

    .summary-highlight-note {
      color: var(--muted);
      font-size: 12px;
      margin-top: 8px;
      line-height: 1.4;
    }

    .chart-shell {
      padding: 14px;
      overflow-x: auto;
    }

    .chart-svg {
      display: block;
      width: 100%;
      min-width: 720px;
      height: auto;
      min-height: 360px;
    }

    .chart-axis {
      stroke: var(--line-strong);
      stroke-width: 1.5;
      shape-rendering: crispEdges;
    }

    .chart-gridline {
      stroke: var(--line);
      stroke-width: 1;
      shape-rendering: crispEdges;
    }

    .chart-tick {
      fill: var(--muted);
      font-size: 11px;
      font-variant-numeric: tabular-nums;
    }

    .chart-axis-title {
      fill: var(--muted);
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0;
      text-transform: uppercase;
    }

    .chart-point {
      stroke: #ffffff;
      stroke-width: 2.5;
    }

    .chart-point-label {
      fill: var(--text);
      font-size: 11px;
      font-weight: 700;
      paint-order: stroke;
      stroke: #ffffff;
      stroke-width: 3px;
      stroke-linejoin: round;
    }

    .chart-legend {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      padding-top: 10px;
    }

    .legend-item {
      display: inline-flex;
      align-items: center;
      gap: 7px;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 5px 9px;
      color: var(--muted);
      font-size: 12px;
      white-space: nowrap;
    }

    .legend-swatch {
      width: 10px;
      height: 10px;
      border-radius: 50%;
      flex: 0 0 auto;
    }

    .muted {
      color: var(--muted);
    }

    .positive {
      color: var(--red);
      font-weight: 800;
    }

    .negative {
      color: var(--green);
      font-weight: 800;
    }

    .neutral {
      color: var(--muted);
      font-weight: 700;
    }

    .value-stack {
      display: grid;
      justify-items: end;
      gap: 3px;
      white-space: nowrap;
    }

    .value-stack small {
      font-size: 11px;
    }

    .value-stack small.positive {
      color: var(--red);
      font-weight: 800;
    }

    .value-stack small.negative {
      color: var(--green);
      font-weight: 800;
    }

    .value-stack small.neutral {
      color: var(--muted);
      font-weight: 700;
    }

    .snapshot-backdrop {
      position: fixed;
      inset: 0;
      z-index: 20;
      display: grid;
      place-items: center;
      padding: 18px;
      background: rgba(15, 23, 42, 0.46);
    }

    .snapshot-backdrop[hidden] {
      display: none;
    }

    .snapshot-editor {
      width: min(1180px, 96vw);
      max-height: 92vh;
      display: grid;
      grid-template-rows: auto 1fr;
      overflow: hidden;
      border: 1px solid var(--line-strong);
      border-radius: 8px;
      background: #ffffff;
      box-shadow: 0 24px 70px rgba(15, 23, 42, 0.25);
    }

    .snapshot-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 18px;
      border-bottom: 1px solid var(--line);
    }

    .snapshot-header h2 {
      margin: 2px 0 0;
      font-size: 21px;
    }

    .snapshot-actions {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      justify-content: flex-end;
    }

    .snapshot-actions button:first-child {
      border-color: var(--accent-dark);
      background: var(--accent);
      color: #ffffff;
      font-weight: 800;
    }

    .snapshot-body {
      overflow: auto;
      padding: 18px;
    }

    .snapshot-note {
      width: 100%;
      min-height: 90px;
      resize: vertical;
      border: 1px solid var(--line-strong);
      border-radius: 7px;
      padding: 10px 12px;
      color: var(--text);
      font: inherit;
    }

    .snapshot-section {
      display: grid;
      gap: 10px;
      margin-top: 18px;
    }

    .snapshot-section h3 {
      margin: 0;
      font-size: 16px;
    }

    .comment-cell {
      min-width: 230px;
      white-space: normal;
    }

    .comment-box {
      width: 100%;
      min-height: 72px;
      resize: vertical;
      border: 1px solid var(--line-strong);
      border-radius: 7px;
      padding: 8px 10px;
      color: var(--text);
      font: inherit;
    }

    body.modal-open {
      overflow: hidden;
    }

    .empty-state {
      padding: 26px;
      color: var(--muted);
      text-align: center;
    }

    .pill {
      display: inline-flex;
      align-items: center;
      min-height: 24px;
      border-radius: 999px;
      background: var(--surface-strong);
      border: 1px solid var(--line);
      color: var(--muted);
      font-size: 12px;
      padding: 3px 8px;
      white-space: nowrap;
    }

    @media (max-width: 1060px) {
      .layout {
        grid-template-columns: 1fr;
      }

      .controls {
        position: static;
      }

      .metric-grid {
        grid-template-columns: repeat(2, minmax(150px, 1fr));
      }
    }

    @media (max-width: 680px) {
      .app {
        padding: 12px;
      }

      .topbar {
        grid-template-columns: 1fr;
      }

      .source-meta {
        justify-content: flex-start;
      }

      .topbar-actions {
        justify-content: flex-start;
      }

      .metric-grid,
      .summary-grid,
      .control-row {
        grid-template-columns: 1fr;
      }

      .bar-row {
        grid-template-columns: 1fr;
      }

      .bar-value {
        text-align: left;
      }
    }
  </style>
</head>
<body>
  <div class="app">
    <header class="topbar">
      <div class="brand-lockup">
        <div class="brand-mark" aria-hidden="true"></div>
        <div>
          <p class="wordmark" aria-label="Arive Homes"><span class="wordmark-arive">arive</span><span class="wordmark-homes">homes</span></p>
          <h1>Floorplan Cost Comparison</h1>
        </div>
      </div>
      <div class="topbar-actions">
        <a class="topbar-button" id="refresh-workflow-link" href="#" target="_blank" rel="noopener">Refresh From SharePoint</a>
        <div class="source-meta" id="source-meta"></div>
      </div>
    </header>

    <script>
      (function setupRefreshWorkflowLink() {
        const link = document.getElementById("refresh-workflow-link");
        if (!link) return;

        const host = window.location.hostname.toLowerCase();
        if (host.endsWith(".github.io")) {
          const owner = host.replace(".github.io", "");
          const repo = window.location.pathname.split("/").filter(Boolean)[0];
          if (owner && repo) {
            link.href = `https://github.com/${owner}/${repo}/actions/workflows/refresh-dashboard.yml`;
            link.title = "Open the GitHub workflow to refresh the dashboard from SharePoint.";
            return;
          }
        }

        link.setAttribute("aria-disabled", "true");
        link.title = "Open this dashboard from GitHub Pages to use this refresh shortcut.";
      })();
    </script>

    <main class="layout">
      <aside class="controls" aria-label="Dashboard controls">
        <section class="control-section">
          <label for="plan-search">Plans</label>
          <input id="plan-search" type="search" placeholder="Search plans">
          <div class="control-row">
            <button type="button" id="select-visible">Select visible</button>
            <button type="button" id="clear-plans">Clear</button>
          </div>
          <div class="plan-list" id="plan-list"></div>
        </section>

        <section class="control-section select-stack">
          <div>
            <label for="baseline-select">Baseline</label>
            <select id="baseline-select"></select>
          </div>
          <div>
            <label for="sort-select">Sort selected</label>
            <select id="sort-select">
              <option value="total-desc">Highest total</option>
              <option value="total-asc">Lowest total</option>
              <option value="diff-desc">Largest increase</option>
              <option value="diff-asc">Largest decrease</option>
              <option value="name">Plan name</option>
            </select>
          </div>
          <label class="toggle-line" for="changed-only">
            <input id="changed-only" type="checkbox" checked>
            <span>Changed line items only</span>
          </label>
          <button type="button" id="create-snapshot" class="wide-button">Create Snapshot</button>
        </section>
      </aside>

      <section class="workspace" aria-live="polite">
        <section class="metric-grid" id="metric-grid"></section>

        <section class="panel">
          <div class="panel-header">
            <h2>Selected Plans</h2>
            <span class="pill" id="selected-pill"></span>
          </div>
          <div class="bar-list" id="bar-list"></div>
        </section>

        <section class="panel">
          <div class="panel-header">
            <h2>Cost Summary</h2>
          </div>
          <div class="table-wrap" id="summary-table"></div>
        </section>

        <section class="panel">
          <div class="panel-header">
            <h2>Cost vs Sq Ft</h2>
          </div>
          <div id="cost-sqft-chart"></div>
        </section>

        <section class="panel">
          <div class="panel-header">
            <h2>Differences Summary</h2>
          </div>
          <div id="differences-summary"></div>
        </section>

        <section class="panel">
          <div class="panel-header">
            <h2>Category Differences</h2>
            <div class="panel-tools">
              <label class="sr-only" for="category-sort">Sort categories</label>
              <select id="category-sort" aria-label="Sort categories">
                <option value="sheet">Sheet order</option>
                <option value="diff-desc">Largest to smallest difference</option>
                <option value="diff-asc">Smallest to largest difference</option>
              </select>
            </div>
          </div>
          <div class="table-wrap" id="category-table"></div>
        </section>

        <section class="panel">
          <div class="panel-header">
            <h2>Line Item Differences</h2>
            <div class="panel-tools">
              <input id="item-search" type="search" placeholder="Search line items">
            </div>
          </div>
          <div class="table-wrap" id="item-table"></div>
        </section>
      </section>
    </main>
  </div>

  <div class="snapshot-backdrop" id="snapshot-panel" hidden>
    <section class="snapshot-editor" role="dialog" aria-modal="true" aria-labelledby="snapshot-title">
      <div class="snapshot-header">
        <div>
          <p class="eyebrow">Snapshot</p>
          <h2 id="snapshot-title">Comparison Snapshot</h2>
        </div>
        <div class="snapshot-actions">
          <button type="button" id="export-snapshot">Save Snapshot</button>
          <button type="button" id="close-snapshot">Close</button>
        </div>
      </div>
      <div class="snapshot-body" id="snapshot-body"></div>
    </section>
  </div>

  <script type="application/json" id="dashboard-data">__DASHBOARD_DATA__</script>
  <script>
    const DATA = JSON.parse(document.getElementById("dashboard-data").textContent);

    const money0 = new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
      maximumFractionDigits: 0
    });

    const money2 = new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: "USD",
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    });

    const number0 = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
    const number1 = new Intl.NumberFormat("en-US", { maximumFractionDigits: 1 });

    const state = {
      baselineId: DATA.plans[0]?.id,
      selected: new Set(DATA.plans.slice(0, 5).map((plan) => plan.id)),
      planQuery: "",
      itemQuery: "",
      sort: "total-desc",
      categorySort: "sheet",
      changedOnly: true,
      snapshot: null
    };

    const byId = new Map(DATA.plans.map((plan) => [plan.id, plan]));
    const allCategories = [...new Set(DATA.items.map((item) => item.category))];

    const els = {
      sourceMeta: document.getElementById("source-meta"),
      planSearch: document.getElementById("plan-search"),
      planList: document.getElementById("plan-list"),
      selectVisible: document.getElementById("select-visible"),
      clearPlans: document.getElementById("clear-plans"),
      createSnapshot: document.getElementById("create-snapshot"),
      baseline: document.getElementById("baseline-select"),
      sort: document.getElementById("sort-select"),
      changedOnly: document.getElementById("changed-only"),
      metrics: document.getElementById("metric-grid"),
      selectedPill: document.getElementById("selected-pill"),
      bars: document.getElementById("bar-list"),
      summary: document.getElementById("summary-table"),
      chart: document.getElementById("cost-sqft-chart"),
      differences: document.getElementById("differences-summary"),
      categorySort: document.getElementById("category-sort"),
      category: document.getElementById("category-table"),
      itemSearch: document.getElementById("item-search"),
      item: document.getElementById("item-table"),
      snapshotPanel: document.getElementById("snapshot-panel"),
      snapshotBody: document.getElementById("snapshot-body"),
      exportSnapshot: document.getElementById("export-snapshot"),
      closeSnapshot: document.getElementById("close-snapshot")
    };

    function escapeHtml(value) {
      return String(value ?? "").replace(/[&<>"']/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        "\"": "&quot;",
        "'": "&#039;"
      })[char]);
    }

    function formatMoney(value, cents = false) {
      if (!Number.isFinite(value)) return "n/a";
      return (cents ? money2 : money0).format(value);
    }

    function formatDelta(value, cents = false) {
      if (!Number.isFinite(value)) return "n/a";
      const formatted = formatMoney(Math.abs(value), cents);
      if (value > 0) return "+" + formatted;
      if (value < 0) return "-" + formatted;
      return formatted;
    }

    function formatNumberDelta(value) {
      if (!Number.isFinite(value)) return "n/a";
      const formatted = number0.format(Math.abs(value));
      if (value > 0) return "+" + formatted;
      if (value < 0) return "-" + formatted;
      return formatted;
    }

    function formatPercent(value) {
      if (!Number.isFinite(value)) return "n/a";
      const sign = value > 0 ? "+" : "";
      return sign + number1.format(value * 100) + "%";
    }

    function percentDiff(value, base) {
      if (!Number.isFinite(value) || !Number.isFinite(base) || base === 0) return null;
      return (value - base) / Math.abs(base);
    }

    function diffClass(value) {
      if (!Number.isFinite(value) || Math.abs(value) < 0.005) return "neutral";
      return value > 0 ? "positive" : "negative";
    }

    function getPlanValue(plan, index) {
      const value = plan.values[index];
      return Number.isFinite(value) ? value : null;
    }

    function getBaseline() {
      return byId.get(state.baselineId) || DATA.plans[0];
    }

    function getVisiblePlans() {
      const query = state.planQuery.trim().toLowerCase();
      if (!query) return DATA.plans;
      return DATA.plans.filter((plan) => {
        return [plan.name, plan.code, plan.importName, plan.group]
          .join(" ")
          .toLowerCase()
          .includes(query);
      });
    }

    function getSelectedPlans() {
      const baseline = getBaseline();
      const selected = DATA.plans.filter((plan) => state.selected.has(plan.id));
      if (baseline && !selected.some((plan) => plan.id === baseline.id)) {
        selected.unshift(baseline);
      }

      selected.sort((a, b) => {
        const base = baseline.totalCost;
        if (state.sort === "total-asc") return a.totalCost - b.totalCost;
        if (state.sort === "total-desc") return b.totalCost - a.totalCost;
        if (state.sort === "diff-asc") return (a.totalCost - base) - (b.totalCost - base);
        if (state.sort === "diff-desc") return (b.totalCost - base) - (a.totalCost - base);
        return a.name.localeCompare(b.name);
      });

      if (baseline) {
        const baseIndex = selected.findIndex((plan) => plan.id === baseline.id);
        if (baseIndex > 0) {
          selected.splice(baseIndex, 1);
          selected.unshift(baseline);
        }
      }

      return selected;
    }

    function setSourceMeta() {
      els.sourceMeta.innerHTML = [
        DATA.source.sheet,
        `${DATA.source.planCount} plans`,
        `${DATA.source.costLineCount} cost lines`,
        `Imported ${DATA.source.importedAt}`
      ].map((value) => `<span>${escapeHtml(value)}</span>`).join("");
    }

    function renderBaselineOptions() {
      els.baseline.innerHTML = DATA.plans
        .map((plan) => `<option value="${escapeHtml(plan.id)}">${escapeHtml(plan.name)} (${escapeHtml(plan.code)})</option>`)
        .join("");
      els.baseline.value = state.baselineId;
    }

    function renderPlanList() {
      const visible = getVisiblePlans();
      const baseline = getBaseline();
      if (!visible.length) {
        els.planList.innerHTML = `<div class="empty-state">No matching plans.</div>`;
        return;
      }

      els.planList.innerHTML = visible.map((plan) => {
        const isBase = baseline && plan.id === baseline.id;
        const checked = isBase || state.selected.has(plan.id);
        const disabled = isBase ? "disabled" : "";
        return `
          <label class="plan-option ${isBase ? "is-base" : ""}">
            <input type="checkbox" data-plan-id="${escapeHtml(plan.id)}" ${checked ? "checked" : ""} ${disabled}>
            <span class="plan-name" title="${escapeHtml(plan.name)}">${escapeHtml(plan.name)}</span>
            <span class="plan-code">${escapeHtml(plan.code)}</span>
          </label>
        `;
      }).join("");
    }

    function renderMetrics() {
      const selected = getSelectedPlans();
      const baseline = getBaseline();
      const totals = selected.map((plan) => plan.totalCost);
      const max = Math.max(...totals);
      const min = Math.min(...totals);
      const average = totals.reduce((sum, value) => sum + value, 0) / Math.max(1, totals.length);
      const highest = selected.find((plan) => plan.totalCost === max);
      const lowest = selected.find((plan) => plan.totalCost === min);
      const spread = max - min;

      const metrics = [
        {
          title: "Baseline total",
          value: formatMoney(baseline.totalCost),
          note: `${baseline.name} (${baseline.code})`
        },
        {
          title: "Selected plans",
          value: number0.format(selected.length),
          note: `${DATA.source.planCount} available`
        },
        {
          title: "Cost spread",
          value: formatMoney(spread),
          note: `${lowest?.name ?? ""} to ${highest?.name ?? ""}`
        },
        {
          title: "Average selected",
          value: formatMoney(average),
          note: `${formatDelta(average - baseline.totalCost)} vs baseline`
        }
      ];

      els.metrics.innerHTML = metrics.map((metric) => `
        <article class="metric">
          <div class="metric-title">${escapeHtml(metric.title)}</div>
          <div class="metric-value">${escapeHtml(metric.value)}</div>
          <div class="metric-note">${escapeHtml(metric.note)}</div>
        </article>
      `).join("");
    }

    function renderBars() {
      const selected = getSelectedPlans();
      const baseline = getBaseline();
      const max = Math.max(...selected.map((plan) => plan.totalCost), 1);
      els.selectedPill.textContent = `${selected.length} selected`;

      els.bars.innerHTML = selected.map((plan) => {
        const diff = plan.totalCost - baseline.totalCost;
        const pct = percentDiff(plan.totalCost, baseline.totalCost);
        const width = Math.max(2, (plan.totalCost / max) * 100);
        return `
          <div class="bar-row">
            <div class="bar-label" title="${escapeHtml(plan.name)}">${escapeHtml(plan.name)}</div>
            <div class="bar-track" aria-hidden="true">
              <div class="bar-fill ${plan.id === baseline.id ? "is-base" : ""}" style="width: ${width.toFixed(2)}%"></div>
            </div>
            <div class="bar-value">
              ${escapeHtml(formatMoney(plan.totalCost))}
              <span class="${diffClass(diff)}">${escapeHtml(formatDelta(diff))}</span>
              <span class="muted">${escapeHtml(formatPercent(pct))}</span>
            </div>
          </div>
        `;
      }).join("");
    }

    function renderSummaryTable() {
      const selected = getSelectedPlans();
      const baseline = getBaseline();
      els.summary.innerHTML = `
        <table>
          <thead>
            <tr>
              <th>Plan</th>
              <th>Type</th>
              <th class="numeric">Sq Ft</th>
              <th class="numeric">Sq Ft diff</th>
              <th class="numeric">Total cost</th>
              <th class="numeric">$ diff</th>
              <th class="numeric">% diff</th>
              <th class="numeric">Cost / sq ft</th>
              <th class="numeric">Missing</th>
            </tr>
          </thead>
          <tbody>
            ${selected.map((plan) => {
              const diff = plan.totalCost - baseline.totalCost;
              const pct = percentDiff(plan.totalCost, baseline.totalCost);
              const sqftDiff = (plan.sqft.total || 0) - (baseline.sqft.total || 0);
              return `
                <tr>
                  <td><strong>${escapeHtml(plan.name)}</strong><br><span class="muted">${escapeHtml(plan.code)}</span></td>
                  <td>${escapeHtml(plan.group || "Unassigned")}</td>
                  <td class="numeric">${escapeHtml(number0.format(plan.sqft.total || 0))}</td>
                  <td class="numeric ${diffClass(sqftDiff)}">${escapeHtml(formatNumberDelta(sqftDiff))}</td>
                  <td class="numeric">${escapeHtml(formatMoney(plan.totalCost))}</td>
                  <td class="numeric ${diffClass(diff)}">${escapeHtml(formatDelta(diff))}</td>
                  <td class="numeric ${diffClass(diff)}">${escapeHtml(formatPercent(pct))}</td>
                  <td class="numeric">${escapeHtml(formatMoney(plan.costPerSqft, true))}</td>
                  <td class="numeric">${escapeHtml(number0.format(plan.missingCount))}</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      `;
    }

    function renderCostSqftChart() {
      const selected = getSelectedPlans();
      const baseline = getBaseline();
      const plans = selected.filter((plan) => Number.isFinite(plan.totalCost) && Number.isFinite(plan.sqft.total));

      if (!plans.length) {
        els.chart.innerHTML = `<div class="empty-state">Select plans to chart cost against square footage.</div>`;
        return;
      }

      const width = 840;
      const height = 390;
      const margin = { top: 26, right: 30, bottom: 70, left: 82 };
      const plotWidth = width - margin.left - margin.right;
      const plotHeight = height - margin.top - margin.bottom;
      const palette = ["#5d6264", "#7ac143", "#8a6d3b", "#3f6f20", "#70787b", "#a3b18a", "#45494b", "#5f8f35"];

      function paddedRange(values, floorAtZero = true) {
        let min = Math.min(...values);
        let max = Math.max(...values);
        if (min === max) {
          const pad = Math.max(Math.abs(max) * 0.08, 1);
          min -= pad;
          max += pad;
        } else {
          const pad = (max - min) * 0.08;
          min -= pad;
          max += pad;
        }
        if (floorAtZero) min = Math.max(0, min);
        return { min, max };
      }

      const xRange = paddedRange(plans.map((plan) => plan.totalCost));
      const yRange = paddedRange(plans.map((plan) => plan.sqft.total));
      const xScale = (value) => margin.left + ((value - xRange.min) / (xRange.max - xRange.min)) * plotWidth;
      const yScale = (value) => margin.top + plotHeight - ((value - yRange.min) / (yRange.max - yRange.min)) * plotHeight;
      const tickCount = 5;
      const ticks = Array.from({ length: tickCount }, (_, index) => index / (tickCount - 1));

      const xTicks = ticks.map((ratio) => {
        const value = xRange.min + (xRange.max - xRange.min) * ratio;
        const x = margin.left + plotWidth * ratio;
        return `
          <line class="chart-gridline" x1="${x.toFixed(2)}" y1="${margin.top}" x2="${x.toFixed(2)}" y2="${margin.top + plotHeight}"></line>
          <text class="chart-tick" x="${x.toFixed(2)}" y="${margin.top + plotHeight + 24}" text-anchor="middle">${escapeHtml(formatMoney(value))}</text>
        `;
      }).join("");

      const yTicks = ticks.map((ratio) => {
        const value = yRange.min + (yRange.max - yRange.min) * ratio;
        const y = margin.top + plotHeight - plotHeight * ratio;
        return `
          <line class="chart-gridline" x1="${margin.left}" y1="${y.toFixed(2)}" x2="${margin.left + plotWidth}" y2="${y.toFixed(2)}"></line>
          <text class="chart-tick" x="${margin.left - 12}" y="${(y + 4).toFixed(2)}" text-anchor="end">${escapeHtml(number0.format(value))}</text>
        `;
      }).join("");

      const points = plans.map((plan, index) => {
        const x = xScale(plan.totalCost);
        const y = yScale(plan.sqft.total);
        const color = plan.id === baseline.id ? "#7ac143" : palette[index % palette.length];
        const labelAnchor = x > width - 180 ? "end" : "start";
        const labelX = x > width - 180 ? x - 10 : x + 10;
        const label = plans.length <= 10
          ? `<text class="chart-point-label" x="${labelX.toFixed(2)}" y="${(y - 10).toFixed(2)}" text-anchor="${labelAnchor}">${escapeHtml(plan.name)}</text>`
          : "";
        return `
          <circle class="chart-point" cx="${x.toFixed(2)}" cy="${y.toFixed(2)}" r="${plan.id === baseline.id ? 7 : 6}" fill="${color}">
            <title>${escapeHtml(`${plan.name}: ${formatMoney(plan.totalCost)}, ${number0.format(plan.sqft.total)} sq ft`)}</title>
          </circle>
          ${label}
        `;
      }).join("");

      const legend = plans.map((plan, index) => {
        const color = plan.id === baseline.id ? "#7ac143" : palette[index % palette.length];
        return `
          <span class="legend-item">
            <span class="legend-swatch" style="background: ${color}"></span>
            ${escapeHtml(plan.name)} &middot; ${escapeHtml(formatMoney(plan.totalCost))} &middot; ${escapeHtml(number0.format(plan.sqft.total))} sq ft
          </span>
        `;
      }).join("");

      els.chart.innerHTML = `
        <div class="chart-shell">
          <svg class="chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Selected plans plotted by total cost and square footage">
            <rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff"></rect>
            ${xTicks}
            ${yTicks}
            <line class="chart-axis" x1="${margin.left}" y1="${margin.top + plotHeight}" x2="${margin.left + plotWidth}" y2="${margin.top + plotHeight}"></line>
            <line class="chart-axis" x1="${margin.left}" y1="${margin.top}" x2="${margin.left}" y2="${margin.top + plotHeight}"></line>
            <text class="chart-axis-title" x="${margin.left + plotWidth / 2}" y="${height - 20}" text-anchor="middle">Total cost</text>
            <text class="chart-axis-title" transform="translate(18 ${margin.top + plotHeight / 2}) rotate(-90)" text-anchor="middle">Sq ft</text>
            ${points}
          </svg>
          <div class="chart-legend">${legend}</div>
        </div>
      `;
    }

    function renderDifferencesSummary() {
      const selected = getSelectedPlans();
      const baseline = getBaseline();
      const comparisonPlans = selected.filter((plan) => plan.id !== baseline.id);

      if (!comparisonPlans.length) {
        els.differences.innerHTML = `<div class="empty-state">Select another plan to summarize differences.</div>`;
        return;
      }

      const categorySummary = allCategories
        .map((category) => {
          const baseValue = baseline.categoryTotals[category] || 0;
          const diffs = comparisonPlans.map((plan) => {
            const value = plan.categoryTotals[category] || 0;
            return {
              plan,
              value,
              diff: value - baseValue
            };
          });
          const averageAbsoluteDiff = diffs.reduce((sum, entry) => sum + Math.abs(entry.diff), 0) / diffs.length;
          const largestPlanDiff = diffs.reduce((best, entry) => {
            if (!best || Math.abs(entry.diff) > Math.abs(best.diff)) return entry;
            return best;
          }, null);
          return {
            category,
            baseValue,
            averageAbsoluteDiff,
            largestPlanDiff
          };
        })
        .sort((a, b) => b.averageAbsoluteDiff - a.averageAbsoluteDiff)[0];

      let largestLineDiff = null;
      DATA.items.forEach((item, index) => {
        const baseValue = getPlanValue(baseline, index);
        const baseForDiff = baseValue || 0;
        comparisonPlans.forEach((plan) => {
          const value = getPlanValue(plan, index);
          const valueForDiff = value || 0;
          const diff = valueForDiff - baseForDiff;
          if (!largestLineDiff || Math.abs(diff) > Math.abs(largestLineDiff.diff)) {
            largestLineDiff = {
              item,
              plan,
              baselineValue: baseValue,
              value,
              diff,
              percent: percentDiff(valueForDiff, baseForDiff)
            };
          }
        });
      });

      const categoryDiff = categorySummary?.largestPlanDiff?.diff || 0;
      const categoryPercent = percentDiff(categorySummary?.largestPlanDiff?.value || 0, categorySummary?.baseValue || 0);

      els.differences.innerHTML = `
        <div class="summary-grid">
          <article class="summary-highlight">
            <div class="summary-highlight-title">Highest average category difference</div>
            <div class="summary-highlight-value">${escapeHtml(categorySummary?.category || "n/a")}</div>
            <div class="summary-highlight-note">
              Average difference: <strong>${escapeHtml(formatMoney(categorySummary?.averageAbsoluteDiff || 0))}</strong><br>
              Largest plan change:
              <span class="${diffClass(categoryDiff)}">${escapeHtml(formatDelta(categoryDiff))}</span>
              <span class="muted">${escapeHtml(formatPercent(categoryPercent))}</span>
              in ${escapeHtml(categorySummary?.largestPlanDiff?.plan?.name || "n/a")}
            </div>
          </article>
          <article class="summary-highlight">
            <div class="summary-highlight-title">Largest overall line item difference</div>
            <div class="summary-highlight-value">${escapeHtml(largestLineDiff?.item?.code || "n/a")} ${escapeHtml(largestLineDiff?.item?.item || "")}</div>
            <div class="summary-highlight-note">
              ${escapeHtml(largestLineDiff?.plan?.name || "n/a")}:
              <span class="${diffClass(largestLineDiff?.diff || 0)}">${escapeHtml(formatDelta(largestLineDiff?.diff || 0, true))}</span>
              <span class="muted">${escapeHtml(formatPercent(largestLineDiff?.percent))}</span><br>
              Baseline ${escapeHtml(formatMoney(largestLineDiff?.baselineValue || 0, true))} to ${escapeHtml(formatMoney(largestLineDiff?.value || 0, true))}
            </div>
          </article>
        </div>
      `;
    }

    function renderCategoryTable() {
      const selected = getSelectedPlans();
      const baseline = getBaseline();
      const comparisonPlans = selected.filter((plan) => plan.id !== baseline.id);

      if (!comparisonPlans.length) {
        els.category.innerHTML = `<div class="empty-state">Select another plan to compare category differences.</div>`;
        return;
      }

      const categories = allCategories.map((category, index) => {
        const baseValue = baseline.categoryTotals[category] || 0;
        const maxAbsoluteDiff = Math.max(...comparisonPlans.map((plan) => {
          const value = plan.categoryTotals[category] || 0;
          return Math.abs(value - baseValue);
        }));
        return { category, index, maxAbsoluteDiff };
      });

      if (state.categorySort === "diff-desc") {
        categories.sort((a, b) => b.maxAbsoluteDiff - a.maxAbsoluteDiff || a.index - b.index);
      } else if (state.categorySort === "diff-asc") {
        categories.sort((a, b) => a.maxAbsoluteDiff - b.maxAbsoluteDiff || a.index - b.index);
      }

      els.category.innerHTML = `
        <table class="comparison-table">
          <thead>
            <tr>
              <th>Category</th>
              <th class="numeric">Baseline</th>
              ${comparisonPlans.map((plan) => `<th class="numeric">${escapeHtml(plan.name)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${categories.map(({ category }) => {
              const baseValue = baseline.categoryTotals[category] || 0;
              return `
                <tr>
                  <td><strong>${escapeHtml(category)}</strong></td>
                  <td class="numeric">${escapeHtml(formatMoney(baseValue))}</td>
                  ${comparisonPlans.map((plan) => {
                    const value = plan.categoryTotals[category] || 0;
                    const diff = value - baseValue;
                    return `
                      <td class="numeric">
                        <div class="value-stack">
                          <span>${escapeHtml(formatMoney(value))}</span>
                          <small class="${diffClass(diff)}">${escapeHtml(formatDelta(diff))} ${escapeHtml(formatPercent(percentDiff(value, baseValue)))}</small>
                        </div>
                      </td>
                    `;
                  }).join("")}
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      `;
    }

    function itemMatchesQuery(item) {
      const query = state.itemQuery.trim().toLowerCase();
      if (!query) return true;
      return [item.code, item.item, item.category].join(" ").toLowerCase().includes(query);
    }

    function renderItemTable() {
      const selected = getSelectedPlans();
      const baseline = getBaseline();
      const comparisonPlans = selected.filter((plan) => plan.id !== baseline.id);

      if (!comparisonPlans.length) {
        els.item.innerHTML = `<div class="empty-state">Select another plan to compare line item differences.</div>`;
        return;
      }

      const rows = DATA.items
        .map((item, index) => ({ item, index }))
        .filter(({ item }) => itemMatchesQuery(item))
        .filter(({ index }) => {
          if (!state.changedOnly) return true;
          const baseValue = getPlanValue(baseline, index) || 0;
          return comparisonPlans.some((plan) => {
            const value = getPlanValue(plan, index) || 0;
            return Math.abs(value - baseValue) >= 0.005;
          });
        });

      if (!rows.length) {
        els.item.innerHTML = `<div class="empty-state">No line items match the current filters.</div>`;
        return;
      }

      els.item.innerHTML = `
        <table class="comparison-table">
          <thead>
            <tr>
              <th>Code</th>
              <th class="item-cell">Line item</th>
              <th class="numeric">Baseline</th>
              ${comparisonPlans.map((plan) => `<th class="numeric">${escapeHtml(plan.name)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${rows.map(({ item, index }) => {
              const baseValue = getPlanValue(baseline, index);
              const baseForDiff = baseValue || 0;
              return `
                <tr>
                  <td><strong>${escapeHtml(item.code)}</strong><br><span class="muted">${escapeHtml(item.category)}</span></td>
                  <td class="item-cell">${escapeHtml(item.item)}</td>
                  <td class="numeric">${escapeHtml(baseValue === null ? "n/a" : formatMoney(baseValue, true))}</td>
                  ${comparisonPlans.map((plan) => {
                    const value = getPlanValue(plan, index);
                    const diff = (value || 0) - baseForDiff;
                    return `
                      <td class="numeric">
                        <div class="value-stack">
                          <span>${escapeHtml(value === null ? "n/a" : formatMoney(value, true))}</span>
                          <small class="${diffClass(diff)}">${escapeHtml(formatDelta(diff, true))} ${escapeHtml(formatPercent(percentDiff(value || 0, baseForDiff)))}</small>
                        </div>
                      </td>
                    `;
                  }).join("")}
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>
      `;
    }

    function snapshotValueCell(value, diff, percent, cents = false) {
      const valueText = value === null ? "n/a" : formatMoney(value, cents);
      return `
        <div class="value-stack">
          <span>${escapeHtml(valueText)}</span>
          <small class="${diffClass(diff)}">${escapeHtml(formatDelta(diff, cents))} ${escapeHtml(formatPercent(percent))}</small>
        </div>
      `;
    }

    function summarizePlanForSnapshot(plan, baseline) {
      const totalDiff = plan.totalCost - baseline.totalCost;
      const sqftDiff = (plan.sqft.total || 0) - (baseline.sqft.total || 0);
      return {
        id: plan.id,
        name: plan.name,
        code: plan.code,
        group: plan.group || "Unassigned",
        sqft: plan.sqft.total || 0,
        sqftDiff,
        totalCost: plan.totalCost,
        totalDiff,
        totalPercent: percentDiff(plan.totalCost, baseline.totalCost),
        costPerSqft: plan.costPerSqft,
        missingCount: plan.missingCount
      };
    }

    function buildSnapshotCategoryRows(baseline, comparisonPlans) {
      const rows = allCategories.map((category, index) => {
        const baseValue = baseline.categoryTotals[category] || 0;
        const maxAbsoluteDiff = Math.max(...comparisonPlans.map((plan) => {
          const value = plan.categoryTotals[category] || 0;
          return Math.abs(value - baseValue);
        }));
        return {
          key: `category-${index}`,
          category,
          index,
          baseValue,
          maxAbsoluteDiff,
          cells: comparisonPlans.map((plan) => {
            const value = plan.categoryTotals[category] || 0;
            const diff = value - baseValue;
            return {
              planName: plan.name,
              value,
              diff,
              percent: percentDiff(value, baseValue)
            };
          })
        };
      });

      if (state.categorySort === "diff-desc") {
        rows.sort((a, b) => b.maxAbsoluteDiff - a.maxAbsoluteDiff || a.index - b.index);
      } else if (state.categorySort === "diff-asc") {
        rows.sort((a, b) => a.maxAbsoluteDiff - b.maxAbsoluteDiff || a.index - b.index);
      }

      return rows;
    }

    function buildSnapshotLineRows(baseline, comparisonPlans) {
      return DATA.items
        .map((item, index) => ({ item, index }))
        .filter(({ item }) => itemMatchesQuery(item))
        .filter(({ index }) => {
          if (!state.changedOnly) return true;
          const baseValue = getPlanValue(baseline, index) || 0;
          return comparisonPlans.some((plan) => {
            const value = getPlanValue(plan, index) || 0;
            return Math.abs(value - baseValue) >= 0.005;
          });
        })
        .map(({ item, index }) => {
          const baseValue = getPlanValue(baseline, index);
          const baseForDiff = baseValue || 0;
          return {
            key: `item-${index}`,
            code: item.code,
            item: item.item,
            category: item.category,
            baseValue,
            cells: comparisonPlans.map((plan) => {
              const value = getPlanValue(plan, index);
              const diff = (value || 0) - baseForDiff;
              return {
                planName: plan.name,
                value,
                diff,
                percent: percentDiff(value || 0, baseForDiff)
              };
            })
          };
        });
    }

    function buildSnapshotData() {
      const selected = getSelectedPlans();
      const baseline = getBaseline();
      const comparisonPlans = selected.filter((plan) => plan.id !== baseline.id);
      const createdAt = new Date();

      return {
        createdAt: createdAt.toLocaleString(),
        createdIso: createdAt.toISOString(),
        source: DATA.source,
        filters: {
          changedOnly: state.changedOnly,
          itemQuery: state.itemQuery,
          categorySort: state.categorySort,
          planSort: state.sort
        },
        baseline: summarizePlanForSnapshot(baseline, baseline),
        plans: selected.map((plan) => summarizePlanForSnapshot(plan, baseline)),
        comparisonPlans: comparisonPlans.map((plan) => summarizePlanForSnapshot(plan, baseline)),
        categoryRows: buildSnapshotCategoryRows(baseline, comparisonPlans),
        lineRows: buildSnapshotLineRows(baseline, comparisonPlans)
      };
    }

    function renderSnapshotSummaryRows(summaryRows) {
      return summaryRows.map((plan) => `
        <tr>
          <td><strong>${escapeHtml(plan.name)}</strong><br><span class="muted">${escapeHtml(plan.code)}</span></td>
          <td>${escapeHtml(plan.group)}</td>
          <td class="numeric">${escapeHtml(number0.format(plan.sqft))}</td>
          <td class="numeric ${diffClass(plan.sqftDiff)}">${escapeHtml(formatNumberDelta(plan.sqftDiff))}</td>
          <td class="numeric">${escapeHtml(formatMoney(plan.totalCost))}</td>
          <td class="numeric ${diffClass(plan.totalDiff)}">${escapeHtml(formatDelta(plan.totalDiff))}</td>
          <td class="numeric ${diffClass(plan.totalDiff)}">${escapeHtml(formatPercent(plan.totalPercent))}</td>
          <td class="numeric">${escapeHtml(formatMoney(plan.costPerSqft, true))}</td>
          <td class="numeric">${escapeHtml(number0.format(plan.missingCount))}</td>
        </tr>
      `).join("");
    }

    function renderSnapshotEditor() {
      const snapshot = state.snapshot;
      if (!snapshot) return;

      if (!snapshot.comparisonPlans.length) {
        els.snapshotBody.innerHTML = `<div class="empty-state">Select at least one plan in addition to the baseline before creating a snapshot.</div>`;
        return;
      }

      const comparedTo = snapshot.comparisonPlans.map((plan) => plan.name).join(", ");
      const lineFilterText = snapshot.filters.changedOnly ? "Changed line items only" : "All line items";
      const searchText = snapshot.filters.itemQuery ? ` matching "${snapshot.filters.itemQuery}"` : "";

      els.snapshotBody.innerHTML = `
        <p class="muted">
          Baseline <strong>${escapeHtml(snapshot.baseline.name)}</strong> &middot;
          Compared to ${escapeHtml(comparedTo)} &middot;
          Created ${escapeHtml(snapshot.createdAt)}
        </p>
        <textarea id="snapshot-note" class="snapshot-note" placeholder="Add overall notes for this comparison"></textarea>

        <section class="snapshot-section">
          <h3>Cost Summary</h3>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Plan</th>
                  <th>Type</th>
                  <th class="numeric">Sq Ft</th>
                  <th class="numeric">Sq Ft diff</th>
                  <th class="numeric">Total cost</th>
                  <th class="numeric">$ diff</th>
                  <th class="numeric">% diff</th>
                  <th class="numeric">Cost / sq ft</th>
                  <th class="numeric">Missing</th>
                </tr>
              </thead>
              <tbody>${renderSnapshotSummaryRows(snapshot.plans)}</tbody>
            </table>
          </div>
        </section>

        <section class="snapshot-section">
          <h3>Category Comments</h3>
          <div class="table-wrap">
            <table class="comparison-table">
              <thead>
                <tr>
                  <th>Category</th>
                  <th class="numeric">Baseline</th>
                  ${snapshot.comparisonPlans.map((plan) => `<th class="numeric">${escapeHtml(plan.name)}</th>`).join("")}
                  <th>Comment</th>
                </tr>
              </thead>
              <tbody>
                ${snapshot.categoryRows.map((row) => `
                  <tr>
                    <td><strong>${escapeHtml(row.category)}</strong></td>
                    <td class="numeric">${escapeHtml(formatMoney(row.baseValue))}</td>
                    ${row.cells.map((cell) => `<td class="numeric">${snapshotValueCell(cell.value, cell.diff, cell.percent)}</td>`).join("")}
                    <td class="comment-cell"><textarea class="comment-box" data-comment-key="${escapeHtml(row.key)}" placeholder="Add comment"></textarea></td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </section>

        <section class="snapshot-section">
          <h3>Line Item Comments</h3>
          <p class="muted">${escapeHtml(lineFilterText)}${escapeHtml(searchText)} &middot; ${escapeHtml(number0.format(snapshot.lineRows.length))} rows captured</p>
          <div class="table-wrap">
            <table class="comparison-table">
              <thead>
                <tr>
                  <th>Code</th>
                  <th class="item-cell">Line item</th>
                  <th class="numeric">Baseline</th>
                  ${snapshot.comparisonPlans.map((plan) => `<th class="numeric">${escapeHtml(plan.name)}</th>`).join("")}
                  <th>Comment</th>
                </tr>
              </thead>
              <tbody>
                ${snapshot.lineRows.map((row) => `
                  <tr>
                    <td><strong>${escapeHtml(row.code)}</strong><br><span class="muted">${escapeHtml(row.category)}</span></td>
                    <td class="item-cell">${escapeHtml(row.item)}</td>
                    <td class="numeric">${escapeHtml(row.baseValue === null ? "n/a" : formatMoney(row.baseValue, true))}</td>
                    ${row.cells.map((cell) => `<td class="numeric">${snapshotValueCell(cell.value, cell.diff, cell.percent, true)}</td>`).join("")}
                    <td class="comment-cell"><textarea class="comment-box" data-comment-key="${escapeHtml(row.key)}" placeholder="Add comment"></textarea></td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </section>
      `;
    }

    function openSnapshotEditor() {
      state.snapshot = buildSnapshotData();
      renderSnapshotEditor();
      els.snapshotPanel.hidden = false;
      document.body.classList.add("modal-open");
      const note = document.getElementById("snapshot-note");
      if (note) note.focus();
    }

    function closeSnapshotEditor() {
      els.snapshotPanel.hidden = true;
      document.body.classList.remove("modal-open");
    }

    function collectSnapshotComments() {
      const note = document.getElementById("snapshot-note")?.value.trim() || "";
      const rowComments = {};
      els.snapshotBody.querySelectorAll("[data-comment-key]").forEach((field) => {
        rowComments[field.dataset.commentKey] = field.value.trim();
      });
      return { note, rowComments };
    }

    function exportedComment(value) {
      return value ? escapeHtml(value) : `<span class="muted">No comment</span>`;
    }

    function buildSnapshotExportHtml(snapshot, note, rowComments) {
      const comparedTo = snapshot.comparisonPlans.map((plan) => plan.name).join(", ");
      const lineFilterText = snapshot.filters.changedOnly ? "Changed line items only" : "All line items";
      const searchText = snapshot.filters.itemQuery ? ` matching "${snapshot.filters.itemQuery}"` : "";
      const categoryRows = snapshot.categoryRows.map((row) => `
        <tr>
          <td><strong>${escapeHtml(row.category)}</strong></td>
          <td class="numeric">${escapeHtml(formatMoney(row.baseValue))}</td>
          ${row.cells.map((cell) => `<td class="numeric">${snapshotValueCell(cell.value, cell.diff, cell.percent)}</td>`).join("")}
          <td class="snapshot-comment">${exportedComment(rowComments[row.key])}</td>
        </tr>
      `).join("");
      const lineRows = snapshot.lineRows.map((row) => `
        <tr>
          <td><strong>${escapeHtml(row.code)}</strong><br><span class="muted">${escapeHtml(row.category)}</span></td>
          <td>${escapeHtml(row.item)}</td>
          <td class="numeric">${escapeHtml(row.baseValue === null ? "n/a" : formatMoney(row.baseValue, true))}</td>
          ${row.cells.map((cell) => `<td class="numeric">${snapshotValueCell(cell.value, cell.diff, cell.percent, true)}</td>`).join("")}
          <td class="snapshot-comment">${exportedComment(rowComments[row.key])}</td>
        </tr>
      `).join("");

      return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Arive Homes Floor Plan Cost Comparison Snapshot</title>
  <style>
    :root {
      --text: #303436;
      --muted: #68716a;
      --line: #dbe2d5;
      --surface: #ffffff;
      --surface-strong: #f2f5ee;
      --accent: #7ac143;
      --accent-dark: #548a2d;
      --charcoal: #5d6264;
      --red: #b42318;
      --green: #067647;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: #f7f8f4;
      color: var(--text);
      font-family: Arial, Helvetica, sans-serif;
      line-height: 1.45;
    }
    main {
      width: min(1220px, calc(100vw - 32px));
      margin: 0 auto;
      padding: 28px 0 42px;
    }
    header {
      display: grid;
      gap: 8px;
      margin-bottom: 18px;
      padding: 18px;
      border: 1px solid var(--line);
      border-top: 5px solid var(--accent);
      border-radius: 8px;
      background: var(--surface);
    }
    h1, h2 {
      margin: 0;
    }
    h1 {
      font-size: 28px;
    }
    h2 {
      margin-top: 22px;
      font-size: 18px;
    }
    .muted {
      color: var(--muted);
    }
    .brand-line {
      display: flex;
      align-items: center;
      gap: 12px;
    }
    .brand-mark {
      position: relative;
      width: 34px;
      height: 42px;
      flex: 0 0 34px;
    }
    .brand-mark::before,
    .brand-mark::after {
      content: "";
      position: absolute;
      bottom: 0;
      border-radius: 1px;
      transform-origin: bottom center;
    }
    .brand-mark::before {
      left: 3px;
      width: 11px;
      height: 42px;
      background: var(--accent);
      transform: skew(-24deg);
    }
    .brand-mark::after {
      left: 20px;
      width: 10px;
      height: 33px;
      background: var(--charcoal);
      transform: skew(26deg);
    }
    .wordmark {
      margin: 0;
      font-size: 24px;
      line-height: 1;
      font-weight: 800;
      text-transform: lowercase;
    }
    .wordmark-arive {
      color: var(--charcoal);
    }
    .wordmark-homes {
      color: var(--accent);
    }
    .panel {
      margin-top: 14px;
      padding: 16px;
      border: 1px solid var(--line);
      border-top: 4px solid var(--accent);
      border-radius: 8px;
      background: var(--surface);
    }
    .table-wrap {
      overflow: auto;
      max-width: 100%;
      margin-top: 10px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--surface);
    }
    table {
      width: 100%;
      min-width: 820px;
      border-collapse: collapse;
      font-size: 13px;
    }
    th, td {
      padding: 10px 12px;
      border-bottom: 1px solid var(--line);
      text-align: left;
      vertical-align: top;
    }
    th {
      background: var(--surface-strong);
      color: var(--muted);
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0;
      white-space: nowrap;
    }
    tr:last-child td {
      border-bottom: 0;
    }
    .numeric {
      text-align: right;
      white-space: nowrap;
    }
    .positive {
      color: var(--red);
      font-weight: 800;
    }
    .negative {
      color: var(--green);
      font-weight: 800;
    }
    .neutral {
      color: var(--muted);
      font-weight: 700;
    }
    .value-stack {
      display: grid;
      justify-items: end;
      gap: 3px;
      white-space: nowrap;
    }
    .value-stack small {
      font-size: 11px;
    }
    .snapshot-comment {
      min-width: 220px;
      white-space: pre-wrap;
    }
    @media print {
      body { background: #ffffff; }
      main { width: 100%; padding: 0; }
      .panel { break-inside: avoid; }
    }
  </style>
</head>
<body>
  <main>
    <header>
      <div class="brand-line">
        <div class="brand-mark" aria-hidden="true"></div>
        <p class="wordmark" aria-label="Arive Homes"><span class="wordmark-arive">arive</span><span class="wordmark-homes">homes</span></p>
      </div>
      <p class="muted">Created ${escapeHtml(snapshot.createdAt)} &middot; Source imported ${escapeHtml(snapshot.source.importedAt)}</p>
      <h1>Floor Plan Cost Comparison Snapshot</h1>
      <p class="muted">Baseline <strong>${escapeHtml(snapshot.baseline.name)}</strong> &middot; Compared to ${escapeHtml(comparedTo)}</p>
    </header>
    <section class="panel">
      <h2>Snapshot Notes</h2>
      <p class="snapshot-comment">${exportedComment(note)}</p>
    </section>
    <section class="panel">
      <h2>Cost Summary</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Plan</th>
              <th>Type</th>
              <th class="numeric">Sq Ft</th>
              <th class="numeric">Sq Ft diff</th>
              <th class="numeric">Total cost</th>
              <th class="numeric">$ diff</th>
              <th class="numeric">% diff</th>
              <th class="numeric">Cost / sq ft</th>
              <th class="numeric">Missing</th>
            </tr>
          </thead>
          <tbody>${renderSnapshotSummaryRows(snapshot.plans)}</tbody>
        </table>
      </div>
    </section>
    <section class="panel">
      <h2>Category Comments</h2>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Category</th>
              <th class="numeric">Baseline</th>
              ${snapshot.comparisonPlans.map((plan) => `<th class="numeric">${escapeHtml(plan.name)}</th>`).join("")}
              <th>Comment</th>
            </tr>
          </thead>
          <tbody>${categoryRows}</tbody>
        </table>
      </div>
    </section>
    <section class="panel">
      <h2>Line Item Comments</h2>
      <p class="muted">${escapeHtml(lineFilterText)}${escapeHtml(searchText)} &middot; ${escapeHtml(number0.format(snapshot.lineRows.length))} rows captured</p>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>Code</th>
              <th>Line item</th>
              <th class="numeric">Baseline</th>
              ${snapshot.comparisonPlans.map((plan) => `<th class="numeric">${escapeHtml(plan.name)}</th>`).join("")}
              <th>Comment</th>
            </tr>
          </thead>
          <tbody>${lineRows}</tbody>
        </table>
      </div>
    </section>
  </main>
</body>
</html>`;
    }

    function downloadHtml(filename, html) {
      const blob = new Blob([html], { type: "text/html;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    }

    function exportSnapshot() {
      if (!state.snapshot) return;
      const { note, rowComments } = collectSnapshotComments();
      const html = buildSnapshotExportHtml(state.snapshot, note, rowComments);
      const stamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, "-");
      downloadHtml(`Floor Plan Cost Comparison Snapshot ${stamp}.html`, html);
    }

    function render() {
      state.selected.add(state.baselineId);
      renderBaselineOptions();
      renderPlanList();
      renderMetrics();
      renderBars();
      renderSummaryTable();
      renderCostSqftChart();
      renderDifferencesSummary();
      renderCategoryTable();
      renderItemTable();
    }

    els.planSearch.addEventListener("input", (event) => {
      state.planQuery = event.target.value;
      renderPlanList();
    });

    els.planList.addEventListener("change", (event) => {
      const id = event.target?.dataset?.planId;
      if (!id) return;
      if (event.target.checked) {
        state.selected.add(id);
      } else {
        state.selected.delete(id);
      }
      render();
    });

    els.selectVisible.addEventListener("click", () => {
      for (const plan of getVisiblePlans()) {
        state.selected.add(plan.id);
      }
      render();
    });

    els.clearPlans.addEventListener("click", () => {
      state.selected = new Set([state.baselineId]);
      render();
    });

    els.baseline.addEventListener("change", (event) => {
      state.baselineId = event.target.value;
      state.selected.add(state.baselineId);
      render();
    });

    els.sort.addEventListener("change", (event) => {
      state.sort = event.target.value;
      render();
    });

    els.categorySort.addEventListener("change", (event) => {
      state.categorySort = event.target.value;
      renderCategoryTable();
    });

    els.changedOnly.addEventListener("change", (event) => {
      state.changedOnly = event.target.checked;
      renderItemTable();
    });

    els.itemSearch.addEventListener("input", (event) => {
      state.itemQuery = event.target.value;
      renderItemTable();
    });

    els.createSnapshot.addEventListener("click", openSnapshotEditor);
    els.closeSnapshot.addEventListener("click", closeSnapshotEditor);
    els.exportSnapshot.addEventListener("click", exportSnapshot);
    els.snapshotPanel.addEventListener("click", (event) => {
      if (event.target === els.snapshotPanel) {
        closeSnapshotEditor();
      }
    });

    setSourceMeta();
    render();
  </script>
</body>
</html>
'@

  if ($DataOutputPath) {
    $dataOutputDirectory = Split-Path $DataOutputPath -Parent
    if ($dataOutputDirectory) {
      New-Item -ItemType Directory -Path $dataOutputDirectory -Force | Out-Null
    }
    Set-Content -Path $DataOutputPath -Value $json -Encoding UTF8
    Write-Output "Wrote $DataOutputPath"
  }

  if ($ExternalDataPath) {
    $loaderStart = @'
  <script>
    const DASHBOARD_DATA_URL = "__DASHBOARD_DATA_URL__";
    let DATA;

    async function loadDashboardData() {
      const response = await fetch(`${DASHBOARD_DATA_URL}?v=${Date.now()}`, { cache: "no-store" });
      if (!response.ok) {
        throw new Error(`Could not load ${DASHBOARD_DATA_URL} (${response.status})`);
      }
      return response.json();
    }

    (async function bootDashboard() {
      DATA = await loadDashboardData();
'@

    $loaderEnd = @'
    setSourceMeta();
    render();
    })().catch((error) => {
      const sourceMeta = document.getElementById("source-meta");
      if (sourceMeta) sourceMeta.innerHTML = "<span>Data not loaded</span>";

      const workspace = document.querySelector(".workspace");
      if (workspace) {
        workspace.innerHTML = "";
        const panel = document.createElement("section");
        panel.className = "panel";
        const message = document.createElement("div");
        message.className = "empty-state";
        const title = document.createElement("strong");
        title.textContent = "Dashboard data could not be loaded.";
        const detail = document.createElement("p");
        detail.textContent = `${error.message} Run the GitHub Action to generate data/floorplans-data.json from the SharePoint Excel workbook.`;
        message.appendChild(title);
        message.appendChild(document.createElement("br"));
        message.appendChild(detail);
        panel.appendChild(message);
        workspace.appendChild(panel);
      }

      console.error(error);
    });
  </script>
'@

    $embeddedStart = @'
  <script type="application/json" id="dashboard-data">__DASHBOARD_DATA__</script>
  <script>
    const DATA = JSON.parse(document.getElementById("dashboard-data").textContent);
'@

    $embeddedEnd = @'
    setSourceMeta();
    render();
  </script>
'@

    $html = $html.Replace($embeddedStart, $loaderStart)
    $html = $html.Replace($embeddedEnd, $loaderEnd)
    $html = $html.Replace('__DASHBOARD_DATA_URL__', $ExternalDataPath)
  }
  else {
    $html = $html.Replace('__DASHBOARD_DATA__', $json)
  }

  $outputDirectory = Split-Path $OutputPath -Parent
  if ($outputDirectory) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
  }
  Set-Content -Path $OutputPath -Value $html -Encoding UTF8
  Write-Output "Wrote $OutputPath"
}
finally {
  $zip.Dispose()
}
