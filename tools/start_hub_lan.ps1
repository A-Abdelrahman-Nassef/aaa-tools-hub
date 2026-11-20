param([int]$Port = 8765, [string]$Bind = '0.0.0.0')
# Start the AAA Tools Hub for colleagues on the office network (Stage A: the hub runs on Ahmed's laptop).
#   powershell -ExecutionPolicy Bypass -File tools\start_hub_lan.ps1
# Leave this window open. Ctrl+C (or closing the window) stops the hub.
# While it runs, Windows will not go to sleep on its own (the screen may still turn off; locking is fine).
# It does NOT change firewall, network or power settings: see docs/office-network.md for those.
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Set-Location $repo
$py = Join-Path $repo '.venv\Scripts\python.exe'

# one hub at a time: two hubs on one database means two workers on one job queue
$busy = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if ($busy) {
  Write-Host "Port $Port is already in use (process $($busy[0].OwningProcess)). Stop the other hub first." -ForegroundColor Red
  exit 1
}

# keep the PC awake while this script runs (released automatically when it exits)
Add-Type -Namespace Win32 -Name Power -MemberDefinition '[DllImport("kernel32.dll")] public static extern uint SetThreadExecutionState(uint esFlags);'
[void][Win32.Power]::SetThreadExecutionState([uint32]'0x80000001')   # ES_CONTINUOUS | ES_SYSTEM_REQUIRED

$ips = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
  Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' }
Write-Host ''
Write-Host 'AAA Tools Hub - colleagues open one of these in their browser:' -ForegroundColor Cyan
foreach ($ip in $ips) { Write-Host ("   http://{0}:{1}    ({2})" -f $ip.IPAddress, $Port, $ip.InterfaceAlias) }
Write-Host ("   http://{0}:{1}    (computer name; works on most office networks)" -f $env:COMPUTERNAME.ToLower(), $Port)
Write-Host 'Use the office network address (Ethernet), not a phone hotspot.' -ForegroundColor DarkGray
Write-Host 'Keep this window open. Ctrl+C stops the hub. The laptop will not sleep while it runs.' -ForegroundColor DarkGray
Write-Host ''
try {
  & $py -m uvicorn hub.app:app --host $Bind --port $Port
} finally {
  [void][Win32.Power]::SetThreadExecutionState([uint32]'0x80000000')   # back to normal sleep rules
}
