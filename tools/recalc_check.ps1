param([string]$src, [string]$dst)
# Recalculate a COPY of an Areas workbook in Excel (the original is not touched), count formula errors
# on every sheet and print the Summary rows. Needs Microsoft Excel.
# Usage: powershell -File tools/recalc_check.ps1 "<workbook>.xlsx" "<copy>.xlsx"
Copy-Item -LiteralPath $src -Destination $dst -Force
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($dst)
  $xl.CalculateFull()
  $errs = 0; $list = @()
  foreach ($ws in $wb.Worksheets) {
    $ur = $ws.UsedRange
    foreach ($c in $ur.Cells) {
      if ($c.HasFormula) {
        $t = [string]$c.Text
        if ($t -match '^#(REF|VALUE|DIV/0|NAME|N/A|NUM|NULL|SPILL|CALC)') { $errs++; if ($list.Count -lt 15) { $list += "$($ws.Name)!$($c.Address(0,0)) $t" } }
      }
    }
  }
  "sheets: " + (($wb.Worksheets | ForEach-Object { $_.Name }) -join ' | ')
  "formula errors: $errs"; $list
  $s = $wb.Worksheets.Item('Summary')
  $ur = $s.UsedRange
  for ($r = 1; $r -le $ur.Rows.Count; $r++) {
    $vals = @()
    for ($k = 1; $k -le [Math]::Min($ur.Columns.Count, 8); $k++) { $vals += [string]$s.Cells.Item($r, $k).Text }
    $line = ($vals -join ' | ').Trim(' |')
    if ($line -match 'IPMS|BUA|Skin|skin|Total') { "R$r  $line" }
  }
  $wb.Close($true)
} finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
