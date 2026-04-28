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
    "PRINTSRV01.contoso.local", or "\\PRINTSRV01".

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

.PARAMETER Force
    Reinstall a connection even if it already exists.

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
#>

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$PrintServer,

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

    [string]$SetDefault,

    [switch]$Force
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
    # Pick the active 10.x.x.x IPv4 address and concatenate octets 2 and 3.
    $candidates = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.IPAddress -like '10.*' -and $_.AddressState -eq 'Preferred' }
    if (-not $candidates) {
        throw "No active 10.x.x.x IPv4 address found on this computer. Use -BuildingPrefix to specify one manually."
    }
    $chosen = @($candidates)[0]
    if (@($candidates).Count -gt 1) {
        $list = ($candidates | ForEach-Object { $_.IPAddress }) -join ', '
        Write-Warning "Multiple 10.x addresses found ($list). Using $($chosen.IPAddress). Override with -BuildingPrefix if wrong."
    }
    $octets = $chosen.IPAddress -split '\.'
    Write-Host "Local IP: $($chosen.IPAddress)" -ForegroundColor Cyan
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

function Get-SharedPrinterOnServer {
    param([string]$Server)
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
    Get-Printer -ErrorAction SilentlyContinue |
        Where-Object { $_.Type -eq 'Connection' } |
        Select-Object -ExpandProperty Name
}

function Install-PrinterConnection {
    param(
        [string]$Server,
        [string]$Share,
        [switch]$Force
    )
    $connection = "\\$Server\$Share"
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
    param([string]$Connection)
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

$server = Resolve-PrintServerName -Name $PrintServer
Write-Host "Print server: $server" -ForegroundColor Cyan

if (-not (Test-PrintServerReachable -Server $server)) {
    Write-Warning "Print server '$server' did not respond to ping. Continuing anyway (ICMP may be blocked)."
}

Write-Host "Enumerating shared printers..." -ForegroundColor Cyan
$shared = @(Get-SharedPrinterOnServer -Server $server)
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
        $prefix = if ($BuildingPrefix) { $BuildingPrefix } else { Get-LocalBuildingNumber }
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
    Install-PrinterConnection -Server $server -Share $share -Force:$Force
}

if ($SetDefault) {
    $defaultShare = $SetDefault.TrimStart('\').TrimEnd('\')
    if ($defaultShare -like '*\*') {
        $defaultConnection = "\\$defaultShare"
    } else {
        $defaultConnection = "\\$server\$defaultShare"
    }
    Set-DefaultPrinterByConnection -Connection $defaultConnection
}

$results | Format-Table -AutoSize
