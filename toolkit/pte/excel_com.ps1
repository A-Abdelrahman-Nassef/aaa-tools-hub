param([Parameter(Mandatory)][string]$Path, [string]$SaveAs = '')
# Excel helper for the PTE steps (stands in for LibreOffice, which the skills use in the cloud sandbox).
#   -Path x.xlsx              recalculate in place, save, print {"errors": n, "samples": [...]} as JSON
#   -Path x.xls -SaveAs y.xlsx convert a legacy .xls to .xlsx
# Needs Microsoft Excel. Opens its own hidden instance, so it does not touch workbooks the user has open.
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AskToUpdateLinks = $false
try {
  # 0: don't update external links. A dummy password: a file that needs one to open fails at once
  # instead of waiting at a hidden password prompt (unprotected files ignore it).
  $m = [Type]::Missing
  $wb = $xl.Workbooks.Open($Path, 0, $false, $m, '~no-password~', '~no-password~')
  if ($SaveAs) {
    $wb.SaveAs($SaveAs, 51)                           # 51 = xlOpenXMLWorkbook (.xlsx)
    $wb.Close($false)
    @{ saved = $SaveAs } | ConvertTo-Json -Compress
    return
  }
  $xl.CalculateFull()
  $n = 0; $samples = New-Object System.Collections.ArrayList
  foreach ($ws in $wb.Worksheets) {
    try { $bad = $ws.UsedRange.SpecialCells(-4123, 16) }  # formulas that return errors
    catch { continue }                                    # none on this sheet
    $n += $bad.Count
    foreach ($c in $bad.Cells) { if ($samples.Count -lt 10) { [void]$samples.Add("$($ws.Name)!$($c.Address(0,0)) $($c.Text)") } }
  }
  $wb.Save(); $wb.Close($false)
  @{ errors = $n; samples = @($samples) } | ConvertTo-Json -Compress
} finally {
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
