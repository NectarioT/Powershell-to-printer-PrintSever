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
| `10.26.26.47`  | `2626`          | `2626_Credentialing`, `2626_Reception`|
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
