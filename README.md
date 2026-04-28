# Connect-ADPrintServer

PowerShell script that connects a Windows workstation to printers shared from
an Active Directory print server. Can auto-select printers by site / building
number derived from the local IPv4 address, or pick them interactively, by
name, or all at once.

The main script is [`Connect-ADPrintServer.ps1`](./Connect-ADPrintServer.ps1).

## How building-aware printer matching works

Workstations sit on `10.<bldg>.<bldg>.<host>`. The script concatenates the
second and third octets of the local IP into a building prefix, then installs
every shared printer whose share name (or display name) starts with that
prefix.

| Workstation IP | Building prefix | Matching shares                       |
|----------------|-----------------|---------------------------------------|
| `10.65.65.47`  | `6565`          | `6565_Credentialing`, `6565_Reception`|
| `10.10.10.50`  | `1010`          | `1010_Lab`, `1010_Reception`          |

## Prerequisites

- Windows 10 / 11 or Windows Server with the `PrintManagement` and `NetTCPIP`
  modules (default on supported Windows).
- PowerShell 5.1 or PowerShell 7 (Windows).
- The user running the script needs permission to enumerate printers on the
  print server. Normal domain users usually do.
- Run **non-elevated** as the user who should receive the printers — printer
  connections are per-user, and elevating installs them for the admin account.
- If your execution policy blocks scripts, see
  [Running the script (execution policy)](#running-the-script-execution-policy)
  below — easiest is the included `.cmd` wrapper.

## Quick reference

| Mode | Command |
|---|---|
| Interactive picker (default) | `.\Connect-ADPrintServer.ps1 -PrintServer PRINTSRV01` |
| Install everything | `... -PrintServer PRINTSRV01 -All` |
| By name (wildcards OK)       | `... -PrintServer PRINTSRV01 -PrinterName 'HR-*','2626_Credentialing'` |
| Auto-detect from local IP    | `... -PrintServer PRINTSRV01 -AutoDetect` |
| Manual building prefix       | `... -PrintServer PRINTSRV01 -BuildingPrefix 2626` |
| Prompt for building          | `... -PrintServer PRINTSRV01 -PromptForBuilding` |
| Pretend a different IP       | `... -PrintServer PRINTSRV01 -LocalIP 10.26.26.47` |
| Set default after install    | `... -SetDefault 2626_Credentialing` |
| Reinstall existing           | `... -Force` |

Run `Get-Help .\Connect-ADPrintServer.ps1 -Full` for full parameter docs.

## Running the script (execution policy)

By default Windows refuses to run unsigned PowerShell scripts. If you see:

> *Connect-ADPrintServer.ps1 is not digitally signed. You cannot run this
> script on the current system.*

…you need to either bypass the policy for this run, or unblock the file.
Pick the option that fits the workstation:

### Option A — use the included `.cmd` wrapper *(easiest)*

[`Run-Connect-ADPrintServer.cmd`](./Run-Connect-ADPrintServer.cmd) launches
PowerShell with `-ExecutionPolicy Bypass` for a single invocation, no system
changes. Forwards every argument to the script:

```cmd
Run-Connect-ADPrintServer.cmd -Test -LocalIP 10.26.26.47
Run-Connect-ADPrintServer.cmd -PrintServer PRINTSRV01 -AutoDetect
```

You can double-click it from File Explorer too — it'll fall back to
`-Test` if you've added that as the default in your own copy.

### Option B — one-shot bypass from PowerShell

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Connect-ADPrintServer.ps1 -Test
```

### Option C — clear the "downloaded from internet" flag

When you grab the script from GitHub or email, Windows stamps it with a
*Mark of the Web* (MOTW). Even with `RemoteSigned` policy that's enough to
block it. Strip the marker once, then run normally:

```powershell
Unblock-File .\Connect-ADPrintServer.ps1
.\Connect-ADPrintServer.ps1 -Test
```

### Option D — change policy for the user *(persistent)*

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

`RemoteSigned` lets locally created/unblocked scripts run. Combine with
`Unblock-File` for downloaded copies.

> Avoid `Set-ExecutionPolicy Bypass` machine-wide; it disables a real
> security control. The wrapper or `-Scope CurrentUser RemoteSigned`
> covers normal day-to-day use.

## Testing the script per computer

The script ships with an offline `-Test` mode that needs **no print server**
and performs **no real `Add-Printer` / `Set-DefaultPrinter` calls**. Use it on
every new workstation before going live, then graduate to a real server call.

### 1. Smoke test — zero infrastructure required

```powershell
.\Connect-ADPrintServer.ps1 -Test
```

Pops the GridView with the built-in mock printers (`2626_*`, `1010_*`,
`HR-Color`) and prints `[TEST] Would connect: …` lines instead of installing
anything.

### 2. Verify the building prefix this computer would use

```powershell
.\Connect-ADPrintServer.ps1 -Test -AutoDetect
```

Prints both `Local IP: <yours>` and `Building prefix: <yours>`. If the prefix
doesn't match the building you're sitting in, the workstation has the wrong
IP / VLAN — fix the network config before deploying.

### 3. Simulate any other workstation's IP

You don't need to physically be in another building to test a building's
behavior:

```powershell
.\Connect-ADPrintServer.ps1 -Test -LocalIP 10.26.26.47   # building 2626
.\Connect-ADPrintServer.ps1 -Test -LocalIP 10.10.10.50   # building 1010
.\Connect-ADPrintServer.ps1 -Test -LocalIP 10.5.5.5      # no matches → expected error
```

Use this to rehearse a site rollout, reproduce a user's report from your
desk, or verify the mock data covers the buildings you care about.

### 4. Manual override for sites without IP-based naming

```powershell
.\Connect-ADPrintServer.ps1 -Test -PromptForBuilding
# Enter building number (e.g. 2626): 2626
```

### 5. Per-computer rollout checklist

1. `ipconfig` — confirm the workstation has a `10.x.x.x` address on the
   correct adapter (and no rogue VPN / Hyper-V / WSL adapter overlapping).
2. `.\Connect-ADPrintServer.ps1 -Test -AutoDetect` — confirm the detected
   prefix matches the building you're physically in.
3. `.\Connect-ADPrintServer.ps1 -PrintServer <real server> -AutoDetect` —
   real install (per-user, no elevation).
4. `Get-Printer | Where-Object Type -eq 'Connection'` — verify the expected
   printers are present.
5. *(Optional)* re-run with `-SetDefault <share>` to pin a default.

### Reference: what `-Test` mode short-circuits

When you pass `-Test`, the script swaps these calls for in-process
no-ops / fixtures so it runs anywhere, even with no network at all:

| Real call | Replaced with under `-Test` |
|---|---|
| `Test-Connection <server>` (reachability ping) | *skipped* |
| `Get-Printer -ComputerName <server>` | `Get-MockSharedPrinter` (built-in fixture, see next table) |
| `Get-CimInstance Win32_Printer -ComputerName <server>` (WMI fallback) | not invoked |
| `Get-Printer` (existing-connection lookup) | empty list |
| `Add-Printer -ConnectionName \\<server>\<share>` | yellow `[TEST] Would connect: …` log line; returns `Status = TestSimulated` |
| `Win32_Printer.SetDefaultPrinter` | yellow `[TEST] Would set default printer to: …` log line |
| `-PrintServer` parameter | defaults to `TESTSRV01` if not supplied |

`Get-NetIPAddress` is the only system call still made, and only when
`-AutoDetect` runs without `-LocalIP`. `-LocalIP` skips it entirely.

### Reference: mock printer dataset

Every `-Test` run uses this fixed list. Pick `-LocalIP` / `-BuildingPrefix`
values whose prefix matches one of these to exercise the matching path:

| ShareName            | DriverName                 | Location            | PortName        |
|----------------------|----------------------------|---------------------|-----------------|
| `2626_Credentialing` | HP Universal Printing PCL 6| Bldg 2626 Floor 2   | IP_10.26.26.100 |
| `2626_Reception`     | HP Universal Printing PCL 6| Bldg 2626 Lobby     | IP_10.26.26.101 |
| `2626_Pharmacy`      | HP Universal Printing PCL 6| Bldg 2626 Floor 1   | IP_10.26.26.102 |
| `1010_Lab`           | HP Universal Printing PCL 6| Bldg 1010 Floor 3   | IP_10.10.10.50  |
| `1010_Reception`     | HP Universal Printing PCL 6| Bldg 1010 Lobby     | IP_10.10.10.51  |
| `HR-Color`           | Xerox WorkCentre           | HR Office           | IP_10.5.5.5     |

To extend the fixture, edit the `Get-MockSharedPrinter` function in
[`Connect-ADPrintServer.ps1`](./Connect-ADPrintServer.ps1).

### Reference: test-related parameters

| Parameter | What it does | Use when |
|---|---|---|
| `-Test` | Activates all the short-circuits above. Composes with any selection mode. | Always while developing or rehearsing. |
| `-AutoDetect` | Reads the local 10.x.x.x NIC, builds prefix from octets 2+3. | Real deployments; verifying a workstation's NIC reports the right building. |
| `-LocalIP <ipv4>` | Pretends the local computer has this IP, derives prefix from it. Validated as IPv4. | Rehearsing other buildings without leaving your desk. |
| `-BuildingPrefix <str>` | Skips IP entirely; uses this prefix verbatim. | Site uses non-IP-based naming, or you want to test a specific prefix string. |
| `-PromptForBuilding` | Prompts via `Read-Host` for the building code. | Manual / training scenarios where the operator types the code. |
| `-SetDefault <share>` | Marks `\\<server>\<share>` as the default after install. Simulated under `-Test`. | Whenever a default is wanted. |
| `-Force` | Reinstalls existing connections instead of skipping them. | Driver updates, stale connection cleanup. |

**Precedence in `Auto` mode** (highest first):
`-BuildingPrefix` → `-PromptForBuilding` → `-LocalIP` → NIC autodetect.

### Reference: sample output

Successful match (3 printers under building `2626`):

```
> .\Connect-ADPrintServer.ps1 -Test -LocalIP 10.26.26.47
=== TEST MODE: no real print server will be contacted ===
Print server: TESTSRV01
Enumerating shared printers...
Found 6 shared printer(s).
Local IP (supplied): 10.26.26.47
Building prefix: 2626
Installing 3 printer connection(s)...
  [TEST] Would connect: \\TESTSRV01\2626_Credentialing
  [TEST] Would connect: \\TESTSRV01\2626_Pharmacy
  [TEST] Would connect: \\TESTSRV01\2626_Reception

Connection                       Status
----------                       ------
\\TESTSRV01\2626_Credentialing   TestSimulated
\\TESTSRV01\2626_Pharmacy        TestSimulated
\\TESTSRV01\2626_Reception       TestSimulated
```

No-match scenario (prefix has no printers):

```
> .\Connect-ADPrintServer.ps1 -Test -LocalIP 10.5.5.5
=== TEST MODE: no real print server will be contacted ===
Print server: TESTSRV01
Enumerating shared printers...
Found 6 shared printer(s).
Local IP (supplied): 10.5.5.5
Building prefix: 55
No printers on 'TESTSRV01' have a name starting with '55'.
```

`-SetDefault` simulation:

```
> .\Connect-ADPrintServer.ps1 -Test -LocalIP 10.26.26.47 -SetDefault 2626_Credentialing
...
[TEST] Would set default printer to: \\TESTSRV01\2626_Credentialing
```

### Reference: internal helpers

These are private functions inside `Connect-ADPrintServer.ps1`; documented
here so you know what to grep for if you're tracing or extending the test
path. None are exported.

| Function | Purpose | `-Test` behavior |
|---|---|---|
| `Resolve-PrintServerName` | Strips `\\…\` and trims the `-PrintServer` argument. | Same. |
| `Test-PrintServerReachable` | Pings the server with `Test-Connection`. | Not called. |
| `Get-MockSharedPrinter` | Returns the static fixture in the table above. | Used as the data source. |
| `Get-SharedPrinterOnServer -Test:$Test` | Real path: `Get-Printer -ComputerName <server>` with WMI fallback. | Returns `Get-MockSharedPrinter`. |
| `Get-LocalBuildingNumber [-FromIP]` | Splits an IPv4 address into octets and concatenates octet 2 + octet 3. | Called with `-FromIP $LocalIP` when `-LocalIP` is supplied; otherwise reads the local NIC. |
| `Get-ExistingConnection -Test:$Test` | Real path: `Get-Printer | Where Type -eq Connection`. | Returns `@()` so test runs aren't influenced by the local printer state. |
| `Install-PrinterConnection -Test:$Test` | Real path: `Add-Printer -ConnectionName \\<server>\<share>`. | Logs `[TEST] Would connect: …` and returns `Status = TestSimulated`. |
| `Set-DefaultPrinterByConnection -Test:$Test` | Real path: invokes `SetDefaultPrinter` on the matching `Win32_Printer`. | Logs `[TEST] Would set default printer to: …` and returns. |

Run `Get-Help .\Connect-ADPrintServer.ps1 -Full` for the complete parameter
documentation generated from the script's comment-based help.

### Removing test connections

`-Test` never installs anything, so there's nothing to clean up. To remove a
*real* connection installed by a non-test run:

```powershell
Remove-Printer -Name '\\PRINTSRV01\2626_Credentialing'
```

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `No active 10.x.x.x IPv4 address found` | Wi-Fi / VPN with non-10 IP. Pass `-LocalIP`, `-BuildingPrefix`, or `-PromptForBuilding`. |
| `Multiple 10.x addresses found` warning | Wi-Fi + Ethernet both up, or a virtual adapter is active. Disambiguate with `-LocalIP`. |
| `No printers on '<server>' have a name starting with '<prefix>'` | Server has no shares for that building, or shares use a different naming convention. Check with `Get-Printer -ComputerName <server>`. |
| `Get-Printer failed against …` warning | Print spooler RPC blocked or remoting denied. Script falls back to WMI automatically. |
| Per-printer `Failed to connect '…'` | Point-and-print driver fetch failed. User needs permission to install drivers, or admin must pre-stage drivers via group policy. |
| `Add-Printer` not recognized | Server Core / minimal install. Run `Add-WindowsCapability -Online -Name 'Print.Management.Console~~~~0.0.1.0'`. |
