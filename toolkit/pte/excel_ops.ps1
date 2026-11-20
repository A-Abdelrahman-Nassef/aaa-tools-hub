param([Parameter(Mandatory)][string]$Src, [Parameter(Mandatory)][string]$Ops, [string]$Out = '')
# Run a JSON list of workbook operations in a hidden Excel instance (Excel does the structural work, so
# references, merges, outlines and everything openpyxl would drop are kept). Prints {"reads": {...}} as JSON.
#   -Src in.xlsx -Ops ops.json -Out out.xlsx   save the result as out.xlsx (the source is never saved)
#   -Out ''                                     run, recalculate and read, then close without saving
# Ops: rename, delete_sheet, insert_rows, copy_rows, delete_rows, delete_cols, set, column, clear,
#      outline, print_area, autofilter, font, wrap, autofit, calc, read, errors. See toolkit/pte/excel.py run_ops().
$ErrorActionPreference = 'Stop'
# JavaScriptSerializer with no length cap (ConvertFrom-Json in Windows PowerShell stops at about 2 MB).
# Not "$ops": variable names ignore case, and $Ops is the [string] parameter.
Add-Type -AssemblyName System.Web.Extensions
$ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$ser.MaxJsonLength = [int]::MaxValue
$list = $ser.DeserializeObject([IO.File]::ReadAllText($Ops, [Text.Encoding]::UTF8))
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.AskToUpdateLinks = $false; $xl.ScreenUpdating = $false
$reads = New-Object 'System.Collections.Generic.Dictionary[string,object]'
$i = 0
function J($v) {                                # one plain value as JSON (strings, numbers, null)
  if ($null -eq $v) { return 'null' }
  if ($v -is [double] -or $v -is [int] -or $v -is [decimal]) { return ([double]$v).ToString('R', [Globalization.CultureInfo]::InvariantCulture) }
  if ($v -is [bool]) { return $(if ($v) { 'true' } else { 'false' }) }
  return $ser.Serialize([string]$v)
}
function Reply($key, $val) {
  if ($key -eq 'error') { return '{"error":' + (J $val) + '}' }
  $parts = foreach ($k in $val.Keys) { (J $k) + ':' + (J $val[$k]) }
  return '{"ops":' + $i + ',"reads":{' + ($parts -join ',') + '}}'
}

function Areas($ws, $cells) {                   # many cells, in chunks Excel's address strings can take
  $out = @(); $buf = @()
  foreach ($c in @($cells)) {
    $buf += [string]$c
    if (($buf -join ',').Length -gt 200) { $out += $ws.Range(($buf -join ',')); $buf = @() }
  }
  if ($buf.Count) { $out += $ws.Range(($buf -join ',')) }
  return $out
}

try {
  $wb = $xl.Workbooks.Open($Src, 0, $true)            # read-only: the source file is never written
  $xl.Calculation = -4135                               # manual while editing
  foreach ($o in $list) {
    $i++
    $ws = if ($o['sheet']) { $wb.Worksheets.Item([string]$o['sheet']) } else { $null }
    switch ([string]$o['op']) {
      'rename'       { $ws.Name = [string]$o['to'] }
      'delete_sheet' { $ws.Delete() }
      'insert_rows'  { [void]$ws.Rows("$($o['at']):$([int]$o['at'] + [int]$o['count'] - 1)").Insert(-4121) }
      'copy_rows'    { [void]$ws.Rows([string]$o['src']).Copy($ws.Rows([string]$o['dst'])) }
      'delete_rows'  { [void]$ws.Rows([string]$o['rows']).Delete() }
      'delete_cols'  { [void]$ws.Columns([string]$o['cols']).Delete() }
      'clear'        { [void]$ws.Range([string]$o['range']).ClearContents() }
      'set'          { $ws.Range([string]$o['cell']).Formula = [string]$o['f'] }     # numbers too: "2" is stored as 2
      'column' {                                        # one call per column: values/formulas top to bottom
        $v = @($o['values']); $n = $v.Count
        $arr = New-Object 'object[,]' $n, 1
        for ($k = 0; $k -lt $n; $k++) { $arr[$k, 0] = $v[$k] }
        $ws.Range("$($o['col'])$($o['from']):$($o['col'])$([int]$o['from'] + $n - 1)").Formula = $arr
      }
      'outline'      { $ws.Rows([string]$o['rows']).OutlineLevel = [int]$o['level'] }
      'print_area'   { $ws.PageSetup.PrintArea = [string]$o['range'] }
      'autofilter'   { if ($ws.AutoFilterMode) { $ws.AutoFilterMode = $false }; [void]$ws.Range([string]$o['range']).AutoFilter() }
      'font'         { foreach ($r in (Areas $ws $o['cells'])) {
                         if ($o.ContainsKey('bold'))   { $r.Font.Bold = [bool]$o['bold'] }
                         if ($o.ContainsKey('italic')) { $r.Font.Italic = [bool]$o['italic'] } } }
      'fill'         { foreach ($r in (Areas $ws $o['cells'])) { $r.Interior.Color = [int]$o['color'] } }
      'wrap'         { $ws.Range([string]$o['range']).WrapText = $true }
      'autofit'      { [void]$ws.Rows([string]$o['rows']).AutoFit() }
      'calc'         { $xl.CalculateFull() }
      'read'         { foreach ($c in @($o['cells'])) { $reads["$($o['sheet'])!$c"] = $ws.Range([string]$c).Value2 } }
      'errors'       {                                  # formula cells that evaluate to an error, per sheet
        $xl.CalculateFull()
        foreach ($s in $wb.Worksheets) {
          try { $bad = $s.UsedRange.SpecialCells(-4123, 16); $reads["errors!$($s.Name)"] = $bad.Count
                $reads["errors!$($s.Name)!first"] = $bad.Cells.Item(1).Address(0, 0) } catch {}
        }
      }
      default        { throw "unknown op $($o['op'])" }
    }
  }
  $xl.Calculation = -4105                               # automatic again, so the saved file is calculated
  $xl.CalculateFull()
  if ($Out) {
    if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out }
    $wb.SaveAs($Out, 51)
  }
  $wb.Close($false)
  Reply 'reads' $reads
} catch {
  $msg = "op $i ($($list[$i - 1]['op']) $($list[$i - 1]['sheet'])): $($_.Exception.Message)"
  Reply 'error' $msg
  try { $wb.Close($false) } catch {}
  exit 1
} finally {
  $xl.ScreenUpdating = $true
  $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
