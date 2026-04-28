<#
.SYNOPSIS
    Connects the local computer to printers shared from an Active Directory print server.

.DESCRIPTION
    Queries an AD print server for its shared printers and installs them as network
    printer connections on the local computer for the current user. Supports
    interactive selection, installing every shared printer, or targeting specific
    printers by name. Optionally sets one of the connected printers as the default.

.PARAMETER PrintServer
    Hostname (or FQDN) of the AD print server. Accepts forms like "PRINTSRV01",
    "PRINTSRV01.contoso.local", or "\\PRINTSRV01". Defaults to "azr01print01"
    when omitted (or to "TESTSRV01" if -Test is supplied without a server).

.PARAMETER PrinterName
    One or more printer share names to install. Wildcards are supported. If omitted,
    behavior depends on -All and -Interactive.

.PARAMETER All
    Install every printer shared on the print server.

.PARAMETER Interactive
    Show a grid view to pick which printers to install. This is the default when
    neither -PrinterName nor -All is supplied.

.PARAMETER SetDefault
    Share name of the printer to set as the default after installation.

.PARAMETER AutoDetect
    Detect the local computer's building number from its 10.x.x.x IPv4 address by
    concatenating the second and third octets (e.g. 10.26.26.47 -> "2626"), then
    install every shared printer whose share name starts with that prefix.

.PARAMETER BuildingPrefix
    Override the auto-detected building number. Useful on VPN or for testing
    (e.g. -BuildingPrefix 2626). Implies -AutoDetect.

.PARAMETER PromptForBuilding
    Prompt the operator to type a building code instead of detecting from IP.

.PARAMETER LocalIP
    Pretend the local computer has this IPv4 address and derive the building
    prefix from it. Mainly useful for testing without changing real network
    config (e.g. -LocalIP 10.26.26.47 -> prefix "2626"). Validated as IPv4.

.PARAMETER Test
    Run in offline test mode: skips reachability checks, uses a hard-coded list
    of fake shared printers, and only simulates Add-Printer / SetDefaultPrinter.
    -PrintServer becomes optional (defaults to TESTSRV01). Combine with any
    selection mode (-AutoDetect, -PromptForBuilding, -All, etc.).

.PARAMETER Force
    Reinstall a connection even if it already exists.

.EXAMPLE
    # Uses the default print server (azr01print01); pick printers interactively.
    .\Connect-ADPrintServer.ps1

.EXAMPLE
    # Uses the default print server; install every match for the local building.
    .\Connect-ADPrintServer.ps1 -AutoDetect

.EXAMPLE
    .\Connect-ADPrintServer.ps1 -PrintServer PRINTSRV01 -Interactive

.EXAMPLE
    .\Connect-ADPrintServer.ps1 -PrintServer PRINTSRV01 -All

.EXAMPLE
    .\Connect-ADPrintServer.ps1 -PrintServer PRINTSRV01.contoso.local `
        -PrinterName 'HR-Color','Finance-*' -SetDefault 'HR-Color'

.EXAMPLE
    # On a workstation with IP 10.26.26.47, install every printer whose share
    # name starts with "2626" (e.g. 2626_Credentialing, 2626_Reception).
    .\Connect-ADPrintServer.ps1 -PrintServer PRINTSRV01 -AutoDetect

.EXAMPLE
    .\Connect-ADPrintServer.ps1 -PrintServer PRINTSRV01 -BuildingPrefix 2626

.EXAMPLE
    # Offline test: pick the building interactively, simulate the install.
    .\Connect-ADPrintServer.ps1 -Test -PromptForBuilding

.EXAMPLE
    # Offline test of auto-detect against a specific prefix, no real install.
    .\Connect-ADPrintServer.ps1 -Test -BuildingPrefix 2626

.EXAMPLE
    # Offline test simulating the workstation in building 2626.
    .\Connect-ADPrintServer.ps1 -Test -LocalIP 10.26.26.47

.EXAMPLE
    # Offline test simulating a different building's workstation.
    .\Connect-ADPrintServer.ps1 -Test -LocalIP 10.10.10.50
#>

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(Position = 0)]
    [string]$PrintServer = 'azr01print01',

    [Parameter(ParameterSetName = 'ByName')]
    [string[]]$PrinterName,

    [Parameter(ParameterSetName = 'All')]
    [switch]$All,

    [Parameter(ParameterSetName = 'Interactive')]
    [switch]$Interactive,

    [Parameter(ParameterSetName = 'Auto')]
    [switch]$AutoDetect,

    [Parameter(ParameterSetName = 'Auto')]
    [string]$BuildingPrefix,

    [Parameter(ParameterSetName = 'Auto')]
    [switch]$PromptForBuilding,

    [Parameter(ParameterSetName = 'Auto')]
    [ValidateScript({
        $parsed = $null
        if ([System.Net.IPAddress]::TryParse($_, [ref]$parsed) -and $parsed.AddressFamily -eq 'InterNetwork') { $true }
        else { throw "Not a valid IPv4 address: '$_'" }
    })]
    [string]$LocalIP,

    [string]$SetDefault,

    [switch]$Force,

    [switch]$Test
)

function Resolve-PrintServerName {
    param([string]$Name)
    $clean = $Name.Trim().TrimStart('\').TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($clean)) {
        throw "PrintServer name is empty."
    }
    return $clean
}

function Get-LocalBuildingNumber {
    # Concatenate octets 2 and 3 of an IPv4 address. Uses -FromIP if given,
    # otherwise picks the active 10.x.x.x address from the local NICs.
    param([string]$FromIP)

    if ($FromIP) {
        $ip = $FromIP
        Write-Host "Local IP (supplied): $ip" -ForegroundColor Cyan
    } else {
        $candidates = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
            Where-Object { $_.IPAddress -like '10.*' -and $_.AddressState -eq 'Preferred' }
        if (-not $candidates) {
            throw "No active 10.x.x.x IPv4 address found on this computer. Use -LocalIP, -BuildingPrefix, or -PromptForBuilding."
        }
        $chosen = @($candidates)[0]
        if (@($candidates).Count -gt 1) {
            $list = ($candidates | ForEach-Object { $_.IPAddress }) -join ', '
            Write-Warning "Multiple 10.x addresses found ($list). Using $($chosen.IPAddress). Override with -LocalIP or -BuildingPrefix if wrong."
        }
        $ip = $chosen.IPAddress
        Write-Host "Local IP: $ip" -ForegroundColor Cyan
    }

    $octets = $ip -split '\.'
    return '{0}{1}' -f $octets[1], $octets[2]
}

function Test-PrintServerReachable {
    param([string]$Server)
    try {
        return Test-Connection -ComputerName $Server -Count 1 -Quiet -ErrorAction Stop
    } catch {
        return $false
    }
}

function Get-MockSharedPrinter {
    # Fixed sample data for offline testing.
    @(
        [pscustomobject]@{ Name='2626_Credentialing'; ShareName='2626_Credentialing'; DriverName='HP Universal Printing PCL 6'; Location='Bldg 2626 Floor 2'; Comment='Credentialing office'; PortName='IP_10.26.26.100' }
        [pscustomobject]@{ Name='2626_Reception';     ShareName='2626_Reception';     DriverName='HP Universal Printing PCL 6'; Location='Bldg 2626 Lobby';   Comment='Front desk';            PortName='IP_10.26.26.101' }
        [pscustomobject]@{ Name='2626_Pharmacy';      ShareName='2626_Pharmacy';      DriverName='HP Universal Printing PCL 6'; Location='Bldg 2626 Floor 1'; Comment='Pharmacy';              PortName='IP_10.26.26.102' }
        [pscustomobject]@{ Name='1010_Lab';           ShareName='1010_Lab';           DriverName='HP Universal Printing PCL 6'; Location='Bldg 1010 Floor 3'; Comment='Lab';                   PortName='IP_10.10.10.50'  }
        [pscustomobject]@{ Name='1010_Reception';     ShareName='1010_Reception';     DriverName='HP Universal Printing PCL 6'; Location='Bldg 1010 Lobby';   Comment='Front desk';            PortName='IP_10.10.10.51'  }
        [pscustomobject]@{ Name='HR-Color';           ShareName='HR-Color';           DriverName='Xerox WorkCentre';            Location='HR Office';         Comment='Color shared';          PortName='IP_10.5.5.5'     }
    )
}

function Get-SharedPrinterOnServer {
    param([string]$Server, [switch]$Test)
    if ($Test) { return Get-MockSharedPrinter }
    try {
        Get-Printer -ComputerName $Server -ErrorAction Stop |
            Where-Object { $_.Shared -eq $true } |
            Select-Object Name, ShareName, DriverName, Location, Comment, PortName
    } catch {
        Write-Warning "Get-Printer failed against '$Server': $($_.Exception.Message)"
        Write-Verbose "Falling back to WMI enumeration."
        Get-CimInstance -ClassName Win32_Printer -ComputerName $Server -ErrorAction Stop |
            Where-Object { $_.Shared -eq $true } |
            Select-Object @{N='Name';E={$_.Name}},
                          @{N='ShareName';E={$_.ShareName}},
                          @{N='DriverName';E={$_.DriverName}},
                          @{N='Location';E={$_.Location}},
                          @{N='Comment';E={$_.Comment}},
                          @{N='PortName';E={$_.PortName}}
    }
}

function Get-ExistingConnection {
    param([switch]$Test)
    if ($Test) { return @() }
    Get-Printer -ErrorAction SilentlyContinue |
        Where-Object { $_.Type -eq 'Connection' } |
        Select-Object -ExpandProperty Name
}

function Install-PrinterConnection {
    param(
        [string]$Server,
        [string]$Share,
        [switch]$Force,
        [switch]$Test
    )
    $connection = "\\$Server\$Share"
    if ($Test) {
        Write-Host "  [TEST] Would connect: $connection" -ForegroundColor Yellow
        return [pscustomobject]@{ Connection = $connection; Status = 'TestSimulated' }
    }
    $existing = Get-ExistingConnection
    if ($existing -contains $connection) {
        if ($Force) {
            Write-Verbose "Removing existing connection '$connection' before reinstall."
            try {
                Remove-Printer -Name $connection -ErrorAction Stop
            } catch {
                Write-Warning "Could not remove existing connection '$connection': $($_.Exception.Message)"
            }
        } else {
            Write-Host "  Already connected: $connection" -ForegroundColor DarkGray
            return [pscustomobject]@{ Connection = $connection; Status = 'AlreadyConnected' }
        }
    }
    try {
        Add-Printer -ConnectionName $connection -ErrorAction Stop
        Write-Host "  Connected: $connection" -ForegroundColor Green
        return [pscustomobject]@{ Connection = $connection; Status = 'Connected' }
    } catch {
        Write-Warning "  Failed to connect '$connection': $($_.Exception.Message)"
        return [pscustomobject]@{ Connection = $connection; Status = "Failed: $($_.Exception.Message)" }
    }
}

function Set-DefaultPrinterByConnection {
    param([string]$Connection, [switch]$Test)
    if ($Test) {
        Write-Host "[TEST] Would set default printer to: $Connection" -ForegroundColor Yellow
        return
    }
    try {
        $cim = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$($Connection.Replace('\','\\'))'" -ErrorAction Stop
        if (-not $cim) { throw "Printer '$Connection' not found on local system." }
        $result = Invoke-CimMethod -InputObject $cim -MethodName SetDefaultPrinter
        if ($result.ReturnValue -ne 0) {
            throw "SetDefaultPrinter returned code $($result.ReturnValue)."
        }
        Write-Host "Default printer set to: $Connection" -ForegroundColor Cyan
    } catch {
        Write-Warning "Could not set default printer '$Connection': $($_.Exception.Message)"
    }
}

# --- Main ---------------------------------------------------------------

if ($Test) {
    Write-Host "=== TEST MODE: no real print server will be contacted ===" -ForegroundColor Yellow
    if (-not $PSBoundParameters.ContainsKey('PrintServer')) {
        $PrintServer = 'TESTSRV01'
    }
}

if ([string]::IsNullOrWhiteSpace($PrintServer)) {
    throw "PrintServer is required. Pass -PrintServer <name> (or -Test for offline testing)."
}

$server = Resolve-PrintServerName -Name $PrintServer
Write-Host "Print server: $server" -ForegroundColor Cyan

if (-not $Test -and -not (Test-PrintServerReachable -Server $server)) {
    Write-Warning "Print server '$server' did not respond to ping. Continuing anyway (ICMP may be blocked)."
}

Write-Host "Enumerating shared printers..." -ForegroundColor Cyan
$shared = @(Get-SharedPrinterOnServer -Server $server -Test:$Test)
if (-not $shared -or $shared.Count -eq 0) {
    throw "No shared printers were found on '$server'. Check permissions and the server name."
}
Write-Host ("Found {0} shared printer(s)." -f $shared.Count) -ForegroundColor Cyan

# Pick which printers to install
$selected = switch ($PSCmdlet.ParameterSetName) {
    'All'         { $shared }
    'ByName'      {
        $nameMatches = foreach ($pattern in $PrinterName) {
            $shared | Where-Object { $_.ShareName -like $pattern -or $_.Name -like $pattern }
        }
        $nameMatches | Sort-Object ShareName -Unique
    }
    'Auto'        {
        $prefix =
            if ($BuildingPrefix) {
                $BuildingPrefix
            } elseif ($PromptForBuilding) {
                $entered = Read-Host 'Enter building number (e.g. 2626)'
                if ([string]::IsNullOrWhiteSpace($entered)) {
                    throw "Building number is required."
                }
                $entered.Trim()
            } elseif ($LocalIP) {
                Get-LocalBuildingNumber -FromIP $LocalIP
            } else {
                Get-LocalBuildingNumber
            }
        Write-Host "Building prefix: $prefix" -ForegroundColor Cyan
        $buildingMatches = $shared | Where-Object {
            $_.ShareName -like "$prefix*" -or $_.Name -like "$prefix*"
        }
        if (-not $buildingMatches) {
            throw "No printers on '$server' have a name starting with '$prefix'."
        }
        $buildingMatches | Sort-Object ShareName
    }
    default {
        $shared |
            Sort-Object ShareName |
            Out-GridView -Title "Select printers on $server to install" -OutputMode Multiple
    }
}

if (-not $selected -or @($selected).Count -eq 0) {
    Write-Warning "No printers selected. Nothing to do."
    return
}

Write-Host ("Installing {0} printer connection(s)..." -f @($selected).Count) -ForegroundColor Cyan
$results = foreach ($p in $selected) {
    $share = if ($p.ShareName) { $p.ShareName } else { $p.Name }
    Install-PrinterConnection -Server $server -Share $share -Force:$Force -Test:$Test
}

if ($SetDefault) {
    $defaultShare = $SetDefault.TrimStart('\').TrimEnd('\')
    if ($defaultShare -like '*\*') {
        $defaultConnection = "\\$defaultShare"
    } else {
        $defaultConnection = "\\$server\$defaultShare"
    }
    Set-DefaultPrinterByConnection -Connection $defaultConnection -Test:$Test
}

$results | Format-Table -AutoSize
