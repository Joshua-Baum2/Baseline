function Get-HostApplications {

    [CmdletBinding()]
    param (
        [string]$OutputDirectory = "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports"
    )

    # Ensure output directory exists
    if (-not (Test-Path -Path $OutputDirectory)) {
        try {
            New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        } catch {
            Write-Error "Failed to create output directory at $OutputDirectory. Error: $_"
            return
        }
    }

    try {
        Write-Verbose "Retrieving installed applications from registry..."

        # Querying both 32-bit and 64-bit registry paths for installed applications
        $applications = Get-ItemProperty -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
                                          'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' |
                        Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

        # Filtering out entries with no DisplayName
        $filteredApplications = $applications | Where-Object { $null -ne $_.DisplayName }

        Write-Verbose "Exporting applications list to CSV..."
        $csvPath = Join-Path $OutputDirectory "Applications.csv"
        $filteredApplications | Export-Csv -Path $csvPath -NoTypeInformation -Force

        Write-Host "Applications list exported to: $csvPath"
    } catch {
        Write-Error "An error occurred while retrieving or exporting applications. Error: $_"
    }
}
