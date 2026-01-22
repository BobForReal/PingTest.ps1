Import-Module ImportExcel -ErrorAction Stop

function Get-HostState {
    param(
        [Parameter(Mandatory)]
        [string]$Ip
    )

    if (Test-Connection -ComputerName $Ip -Count 1 -Quiet -ErrorAction SilentlyContinue) {
        "Up"
    } else {
        "Down"
    }
}

function Resolve-Hostname {
    param(
        [Parameter(Mandatory)]
        [string]$Ip
    )

    try {
        ([System.Net.Dns]::GetHostEntry($Ip)).HostName
    } catch {
        $null
    }
}

function Main {
    $scriptDir   = $PSScriptRoot
    $serversPath = Join-Path $scriptDir "servers.txt"

    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $xlsxPath  = Join-Path $scriptDir "results-$timestamp.xlsx"

    $results =
        Get-Content -Path $serversPath |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ } |
        ForEach-Object {
            $ip = $_
            [PSCustomObject]@{
                IP           = $ip
                ComputerName = Resolve-Hostname -Ip $ip
                State        = Get-HostState   -Ip $ip
                Timestamp    = Get-Date
            }
        }

    $results | Export-Excel -Path $xlsxPath -AutoSize

    Invoke-Item -Path $xlsxPath
}

Main
