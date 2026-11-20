param([int]$Port = 8765)
# Is this laptop ready for colleagues to reach the hub? Read-only: it changes nothing.
#   powershell -ExecutionPolicy Bypass -File tools\check_lan.ps1
# Each line is OK, FIX (with what to do, see docs/office-network.md) or INFO.
$ErrorActionPreference = 'SilentlyContinue'
$repo = Split-Path -Parent $PSScriptRoot
function Show($state, $text) {
  $c = @{ OK = 'Green'; FIX = 'Red'; INFO = 'DarkGray' }[$state]
  Write-Host ('{0,-5} {1}' -f $state, $text) -ForegroundColor $c
}
Write-Host ''
Write-Host 'AAA Tools Hub - office network check' -ForegroundColor Cyan

# 1. network and profile
$profiles = Get-NetConnectionProfile
if (-not $profiles) { Show FIX 'No network connection.' }
foreach ($p in $profiles) {
  $ip = (Get-NetIPAddress -InterfaceAlias $p.InterfaceAlias -AddressFamily IPv4 | Select-Object -First 1).IPAddress
  $where = "$($p.InterfaceAlias) '$($p.Name)' $ip"
  if ($p.NetworkCategory -eq 'Public') {
    Show FIX "$where is a Public network: colleagues cannot connect. On the office network set it to Private (setup step 2); never on a hotspot or cafe."
  } else {
    Show OK "$where is $($p.NetworkCategory)."
  }
  if ($ip -like '172.20.10.*') { Show FIX "$where looks like a phone hotspot: use the office cable or Wi-Fi." }
}

# 2. firewall rule for the hub's port
$rule = Get-NetFirewallRule -DisplayName 'AAA Tools Hub*' | Where-Object { $_.Enabled -eq 'True' -and $_.Direction -eq 'Inbound' -and $_.Action -eq 'Allow' }
if (-not $rule) {
  Show FIX "No firewall rule lets colleagues in on port ${Port}: run setup step 3 in an administrator PowerShell."
} else {
  $ports = ($rule | Get-NetFirewallPortFilter).LocalPort -join ','
  $addr = ($rule | Get-NetFirewallAddressFilter).RemoteAddress -join ','
  $prof = ($rule.Profile | ForEach-Object { $_.ToString() }) -join ','
  if ($ports -notmatch "\b$Port\b") { Show FIX "Firewall rule found but for port(s) $ports, not $Port." }
  elseif ($prof -match 'Public' -or $prof -eq 'Any') { Show FIX "Firewall rule also opens the port on Public networks ($prof): limit it to Private,Domain." }
  else { Show OK "Firewall rule '$($rule[0].DisplayName)': port $ports from $addr on $prof." }
}

# 3. power: lid and sleep when plugged in
$lid = (powercfg /q SCHEME_CURRENT 4f971e89-eebd-4455-a8de-9e59040e7347 5ca83367-6e45-459f-a27b-476b1d01c936 |
        Select-String 'Current AC Power Setting Index') -replace '.*:\s*', ''
if ($lid -eq '0x00000000') { Show OK 'Closing the lid on the charger does nothing.' }
elseif ($lid) { Show FIX 'Closing the lid on the charger sleeps or shuts down the laptop: set it to "Do nothing" (setup step 5).' }
else { Show INFO 'Lid setting hidden by this power plan: check Control Panel > Power Options > "Choose what closing the lid does" > Plugged in = Do nothing, if you will close the lid.' }
$ac = (Get-CimInstance -Namespace root\cimv2 -ClassName Win32_Battery | Select-Object -First 1).BatteryStatus
if ($ac -eq 1) { Show FIX 'The laptop is on battery: plug it in (it sleeps after 30 minutes on battery).' }
else { Show OK 'On the charger (the start script also keeps it awake while the hub runs).' }

# 4. the port: free, or a hub already listening for the office network
$l = Get-NetTCPConnection -LocalPort $Port -State Listen
if (-not $l) { Show INFO "Nothing on port $Port yet: start the hub with tools\start_hub_lan.ps1." }
elseif ($l.LocalAddress -contains '0.0.0.0') { Show OK "A hub is listening on port $Port for the network." }
else { Show FIX "Port $Port is held by a hub on $($l.LocalAddress -join ',') only (the Claude app preview): stop it, then run tools\start_hub_lan.ps1." }

# 5. accounts
$py = Join-Path $repo '.venv\Scripts\python.exe'
$users = & $py -c "import sys; sys.path.insert(0, r'$repo'); from hub import db; db.init(); rows = db.q('select name, is_admin, initials from users'); print(sum(1 for r in rows if not r['is_admin']), sum(1 for r in rows if r['is_admin']), sum(1 for r in rows if not r['initials']))"
if ($users) {
  $m, $a, $noini = $users.Trim() -split ' '
  if ([int]$m -eq 0) { Show FIX 'No colleague (member) accounts yet: setup step 6.' } else { Show OK "$m colleague account(s), $a admin(s)." }
  if ([int]$noini -gt 0) { Show INFO "$noini account(s) without initials for file names (Account page, or manage.py initials)." }
}

# 6. what to give colleagues
foreach ($p in $profiles) {
  $ip = (Get-NetIPAddress -InterfaceAlias $p.InterfaceAlias -AddressFamily IPv4 | Select-Object -First 1).IPAddress
  if ($ip) { Show INFO "Colleagues would open http://${ip}:$Port (or http://$($env:COMPUTERNAME.ToLower()):$Port)." }
}
Write-Host ''
