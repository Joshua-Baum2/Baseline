function Get-NetworkBaseline {
    <#
    .SYNOPSIS
    Collects a baseline of network performance and system status metrics and saves output to CSV files.

    .DESCRIPTION
    This function gathers system information, network adapter details, IP configuration,
    network statistics, performs a connectivity check, and retrieves open network ports, along with identifying applications using open ports. It saves different data sections to separate CSV files for easy viewing.

    .OUTPUTS
    CSV files with collected data.

    .EXAMPLE
    Get-NetworkBaseline -OutputDirectory "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports"
    #>

    [CmdletBinding()]
    param (
        [string]$OutputDirectory = "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports"
    )

    # Ensure output directory exists
    if (-not (Test-Path -Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
    }

    # Create a hashtable to store collected data
    $report = @{
        ReportMetadata = @(
            [PSCustomObject]@{
                GeneratedOn    = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                GeneratedBy    = $env:USERNAME
                ComputerName   = $env:COMPUTERNAME
                OS             = (Get-CimInstance -ClassName Win32_OperatingSystem).Caption
            }
        )
        Data = @{}
    }

    try {
        # Gathering basic system info
        Write-Host "Collecting System Information..."
        $systemInfo = Get-ComputerInfo | Select-Object CsName, WindowsVersion, WindowsBuildLabEx, CsManufacturer, CsModel
        $report['Data']['SystemInformation'] = $systemInfo
        $systemInfo | Export-Csv -Path (Join-Path $OutputDirectory "SystemInformation.csv") -NoTypeInformation -Force
    } catch {
        Write-Warning "Failed to collect system information: $_"
    }

    try {
        # Gathering Network Adapter details with friendly names
        Write-Host "Collecting Network Adapter Information..."
        $adapters = Get-NetAdapter | Select-Object Name, Status, MacAddress, LinkSpeed, MediaType
        foreach ($adapter in $adapters) {
            # Add friendly device name (if available)
            $adapterName = Get-NetAdapter -Name $adapter.Name | Select-Object -ExpandProperty InterfaceDescription
            $adapter | Add-Member -MemberType NoteProperty -Name FriendlyName -Value $adapterName
        }
        $report['Data']['NetworkAdapters'] = $adapters
        $adapters | Export-Csv -Path (Join-Path $OutputDirectory "NetworkAdapters.csv") -NoTypeInformation -Force
    } catch {
        Write-Warning "Failed to collect network adapter information: $_"
    }

    try {
        # Gathering IP configuration details
        Write-Host "Collecting IP Configuration Details..."
        $ipConfigs = Get-NetIPAddress | Select-Object InterfaceAlias, IPAddress, AddressFamily, PrefixLength, Type
        foreach ($ipConfig in $ipConfigs) {
            # Resolve device names using the IP address (if possible)
            if ($ipConfig.IPAddress -match '(\d{1,3}\.){3}\d{1,3}') {
                try {
                    $hostname = [System.Net.Dns]::GetHostEntry($ipConfig.IPAddress).HostName
                    $ipConfig | Add-Member -MemberType NoteProperty -Name HostName -Value $hostname
                } catch {
                    $ipConfig | Add-Member -MemberType NoteProperty -Name HostName -Value "N/A"
                }
            } else {
                $ipConfig | Add-Member -MemberType NoteProperty -Name HostName -Value "N/A"
            }
        }
        $report['Data']['IPConfiguration'] = $ipConfigs
        $ipConfigs | Export-Csv -Path (Join-Path $OutputDirectory "IPConfiguration.csv") -NoTypeInformation -Force
    } catch {
        Write-Warning "Failed to collect IP configuration details: $_"
    }

    try {
        # Checking network interface statistics
        Write-Host "Collecting Network Interface Statistics..."
        $networkStats = Get-NetAdapterStatistics | Select-Object Name, ReceivedBytes, SentBytes, ReceivedUnicastPackets, SentUnicastPackets
        $report['Data']['NetworkStatistics'] = $networkStats
        $networkStats | Export-Csv -Path (Join-Path $OutputDirectory "NetworkStatistics.csv") -NoTypeInformation -Force
    } catch {
        Write-Warning "Failed to collect network interface statistics: $_"
    }

    try {
        # Performing a basic network connectivity test
        Write-Host "Pinging Google DNS to check connectivity..."
        $connectivityTest = Test-Connection -ComputerName 8.8.8.8 -Count 4 -ErrorAction SilentlyContinue | Select-Object Address, ResponseTime, StatusCode, TimeToLive
        $report['Data']['ConnectivityTest'] = $connectivityTest
        $connectivityTest | Export-Csv -Path (Join-Path $OutputDirectory "ConnectivityTest.csv") -NoTypeInformation -Force
    } catch {
        Write-Warning "Failed to perform connectivity test: $_"
    }

    try {
        # Checking open network ports and identifying associated applications
        Write-Host "Collecting Open Network Ports Information..."
        $openPorts = Get-NetTCPConnection | Where-Object { $_.State -eq 'Listen' } | Select-Object LocalAddress, LocalPort, OwningProcess
        foreach ($port in $openPorts) {
            try {
                $hostname = [System.Net.Dns]::GetHostEntry($port.LocalAddress).HostName
                $port | Add-Member -MemberType NoteProperty -Name HostName -Value $hostname
            } catch {
                $port | Add-Member -MemberType NoteProperty -Name HostName -Value "N/A"
            }

            # Retrieve the process name associated with the OwningProcess (PID)
            try {
                $process = Get-Process -Id $port.OwningProcess -ErrorAction SilentlyContinue
                if ($process) {
                    $port | Add-Member -MemberType NoteProperty -Name ProcessName -Value $process.ProcessName
                } else {
                    $port | Add-Member -MemberType NoteProperty -Name ProcessName -Value "N/A"
                }
            } catch {
                $port | Add-Member -MemberType NoteProperty -Name ProcessName -Value "N/A"
            }
        }
        $report['Data']['OpenPorts'] = $openPorts
        $openPorts | Export-Csv -Path (Join-Path $OutputDirectory "OpenPorts.csv") -NoTypeInformation -Force
    } catch {
        Write-Warning "Failed to collect open network ports: $_"
    }

    Write-Host "Network Baseline Collection Completed."
}

# Example usage:
Get-NetworkBaseline -OutputDirectory "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports"


