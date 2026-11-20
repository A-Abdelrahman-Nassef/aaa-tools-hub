# Running the hub for colleagues on the office network (Stage A)

The hub runs on Ahmed's laptop; colleagues open it in their own browser. They install nothing.
Stage C moves it to a server; until then the laptop has to be on, awake and on the office network.

## What the laptop looked like on 18 Nov 2026 (20 Nov: home Wi-Fi 192.168.100.21, Public; no rule yet)
- Computer name **ANASEF**. Office network (Ethernet) address **192.168.10.26**, from DHCP (it can change).
  On 18 Nov the Ethernet was not connected and the laptop was on Wi-Fi 172.20.10.5, a phone hotspot
  (profile Public): colleagues cannot reach the hub that way.
- Windows Firewall is on for Domain, Private and Public, with no rule for the hub: incoming connections are blocked.
- On the charger the laptop never sleeps; on battery it sleeps after 30 minutes.
- The hub (in the Claude app preview) listens on 127.0.0.1 only: no one else can reach it.

## One-time setup (Ahmed does these; Claude does not change security settings)

1. **Be on the office network.** Cable (or the office Wi-Fi), not the phone hotspot.
2. **Mark the office network as Private.** Settings > Network & internet > Ethernet (or the office Wi-Fi) >
   Network profile type: **Private**. (On a domain network it shows "Domain" and that is fine.)
3. **Open port 8765 to the office network only.** In PowerShell **as administrator**:
   ```
   New-NetFirewallRule -DisplayName "AAA Tools Hub 8765 (office LAN)" -Direction Inbound -Protocol TCP -LocalPort 8765 -RemoteAddress LocalSubnet -Profile Private,Domain -Action Allow
   ```
   Only the hub's port, only from the local network, never on Public networks (cafes, hotspots).
   If Windows asks "Allow python.exe on networks?" the first time the hub starts, allow **Private** only.
   To remove it later: `Remove-NetFirewallRule -DisplayName "AAA Tools Hub 8765 (office LAN)"`.
4. **A fixed address.** Ask IT for a DHCP reservation for the laptop's network adapter (MAC address from
   `Get-NetAdapter | Select Name, MacAddress`, take the office one). Until then, give colleagues the address the
   start script prints each day, or try `http://anasef:8765` (the computer name works on most office networks).
5. **Keep it awake.** Work on the charger (sleep is already "never" on AC). Power options >
   "Choose what closing the lid does" > Plugged in: **Do nothing**, if the lid will be closed.
   Lock the screen (Win+L) when away: fine. Do **not** sign out, restart or shut down: the hub and its Excel
   steps run in your Windows session. Set Windows Update active hours to office hours so it does not restart.
6. **Accounts.** One per colleague, not admin, password given in person:
   ```
   .venv\Scripts\python.exe -m hub.manage adduser colleague@company.com "Name"
   ```

## Check before colleagues come (read-only)

```
powershell -ExecutionPolicy Bypass -File tools\check_lan.ps1
```
One line per item, OK / FIX / INFO: network profile, firewall rule, charger and lid, port 8765, colleague accounts,
and the address to give colleagues. It changes nothing.

## Every day

```
powershell -ExecutionPolicy Bypass -File tools\start_hub_lan.ps1
```
- Run it in the aaa-tools-hub folder, in a normal (not admin) window, and leave the window open.
- It prints the addresses to give colleagues, refuses to start if another hub already holds port 8765 (stop the
  Claude app preview "hub" first: one hub at a time, or two workers would share one job queue), and keeps Windows
  from sleeping while it runs. Ctrl+C stops the hub.
- Check the keep-awake (optional, administrator PowerShell): `powercfg /requests` lists powershell.exe under SYSTEM.

## What colleagues can and cannot do (Stage A)
- PTE Phase 1 modules (01 Unlock, 02 Register, 03 Stack, 04 Naming, 05 Load, 06 Link) are scripts only: they can run
  them end to end. Their output files carry the initials of whoever ran the step: each colleague sets theirs on the
  Account page (click their name, top right), or Ahmed runs `manage.py initials <email> XX`.
  Jobs run one at a time on the laptop, so a long Excel step makes the next job wait.
- 00 Areas: the Claude step (reading drawings) is done by Ahmed in Claude Code. Colleagues upload **JSON** specs
  only; `.py` specs are refused for non-admins because the pipeline runs them as code on this laptop.
- The hub uses plain http on the office network: passwords cross the LAN unencrypted. Acceptable for a short
  pilot on a trusted office network; never forward the port to the internet. Stage C needs https.

## Test with one colleague (about 45 minutes)

**Before (Ahmed, 10 minutes)**
1. Setup steps 1-6 above done (`tools\check_lan.ps1` shows no FIX); stop the Claude app preview "hub";
   start `tools\start_hub_lan.ps1`; note the office address it prints.
2. On the laptop, open `http://<that address>:8765` (not localhost): the sign-in page shows.
3. Create the colleague's account. Put their test files somewhere they can reach (their own PC or a share).

**Connection check (with the colleague, 5 minutes)**
1. On their PC: open `http://<address>:8765`. Sign-in page = network and firewall are right.
   - Times out: the firewall rule or the network profile (Public?), or the laptop is on the hotspot.
   - "Refused": the hub is not running, or was started on 127.0.0.1 (the Claude app preview) instead of the script.
   - Name fails but the address works: use the address (or ask IT about name resolution).
2. They sign in with their own account.

**Tasks (read them out; Ahmed only watches and notes)**
1. Areas: run `docs/colleague-test.md` with the JSON spec `jobs\SAMPLE\spec\villa_type_a.json`
   (same numbers as the .py: BUA 360.50, IPMS 1 448.21).
2. PTE: create a project (with a short name for files), upload 3-4 BOQs as "BOQ (as received)", run 01 Unlock BOQs,
   then 02 Asset / Package Register: mark each BOQ as a package or a whole asset, approve, download the register
   (`<date> - <client> - <project> - Asset & Package Register - Rev.00 - <initials>.xlsx`) and open it on their PC.
3. While their job runs, Ahmed starts a job too: both finish (one after the other).

**After**
- Note where they hesitated (table in `docs/colleague-test.md`), page speed, downloads opening on their PC.
- Stop the hub (Ctrl+C) if colleagues should not use it yet; the firewall rule can stay or be removed.

**Pass:** they reach the hub from their own PC, sign in, finish both tasks without help, and the downloaded files
open on their machine.

## Stage C (later)
Move `aaa-tools-hub` and `data/` to a server with https, a fixed name, backups, and accounts managed there.
The Excel steps (PTE modules) need Windows with Excel on that server, or a rewrite of those steps without Excel.
