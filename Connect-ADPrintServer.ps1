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

.PARAMETER Force
    Reinstall a connection even if it already exists.

.EXAMPLE
    .\Connect-ADPrintServer.ps1 -PrintServer PRINTSRV01 -Interactive

.EXAMPLE
    .\Connect-ADPrintServer.ps1 -PrintServer PRINTSRV01 -All

.EXAMPLE
    .\Connect-ADPrintServer.ps1 -PrintServer PRINTSRV01.contoso.local `
        -PrinterName 'HR-Color','Finance-*' -SetDefault 'HR-Color'
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
        $matches = foreach ($pattern in $PrinterName) {
            $shared | Where-Object { $_.ShareName -like $pattern -or $_.Name -like $pattern }
        }
        $matches | Sort-Object ShareName -Unique
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
