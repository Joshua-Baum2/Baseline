function Set-AuditPolicy {
    [CmdletBinding()]
    param (
        [string]$ExcelFilePath = "C:\Baseline\Reports\Baseline.xlsx",
        [string]$SheetName = "Policy Analyzer"  
    )

    # Load ImportExcel module if available
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Error "The ImportExcel module is required. Install it by running: Install-Module -Name ImportExcel"
        return
    }

    try {
        # Read data from Excel
        $auditData = Import-Excel -Path $ExcelFilePath -WorksheetName $SheetName

        foreach ($row in $auditData) {
            $policyType = $row.'Policy Type'
            $policyGroup = $row.'Policy Group or Registry Key'
            $policySetting = $row.'Policy Setting'
            $baseline = $row.'Baseline(s)'

            if ($policyType -eq 'Audit Policy') {
                # Handle Audit Policies
                $baselineSuccess = $null
                $baselineFailure = $null

                if ($baseline -eq 'Success and Failure') {
                    $baselineSuccess = 'enable'
                    $baselineFailure = 'enable'
                } elseif ($baseline -match "Success:(\w+)") {
                    $baselineSuccess = $matches[1]
                } elseif ($baseline -match "Failure:(\w+)") {
                    $baselineFailure = $matches[1]
                } elseif ($baseline -eq 'Success') {
                    $baselineSuccess = 'enable'
                } elseif ($baseline -eq 'Failure') {
                    $baselineFailure = 'enable'
                }

                # Get the current effective state of the policy
                $effectivePolicy = (auditpol /get /subcategory:"$policySetting") | Out-String
                $effectiveSuccess = if ($effectivePolicy -match "Success\s+:\s+(\w+)") { $matches[1] } else { $null }
                $effectiveFailure = if ($effectivePolicy -match "Failure\s+:\s+(\w+)") { $matches[1] } else { $null }

                # Check if the effective state matches the baseline
                $successNeedsUpdate = $baselineSuccess -and ($effectiveSuccess -ne $baselineSuccess)
                $failureNeedsUpdate = $baselineFailure -and ($effectiveFailure -ne $baselineFailure)

                if ($successNeedsUpdate -or $failureNeedsUpdate) {
                    Write-Host "Applying policy for: $policySetting"

                    try {
                        # Construct the auditpol command based on needs
                        $command = "auditpol /set /subcategory:""$policySetting"""
                        if ($successNeedsUpdate) {
                            $command += " /success:$baselineSuccess"
                        }
                        if ($failureNeedsUpdate) {
                            $command += " /failure:$baselineFailure"
                        }

                        # Log the command
                        Write-Host "Debug: Command to Execute: $command"

                        # Execute the command
                        Invoke-Expression $command

                        Write-Host "Policy applied successfully for: $policySetting"
                    } catch {
                        Write-Error "Failed to apply policy for: $policySetting. Error: $_"
                    }
                } else {
                    Write-Host "Policy already compliant for: $policySetting"
                }
            } elseif ($policyType -eq 'Registry Policy') {
                # Handle Registry modifications, checking if path exists first
                if (Test-Path $policyGroup) {
                    try {
                        $registryPath = $policyGroup
                        $registryValueName = $policySetting
                        $valueToSet = $baseline

                        Write-Host "Applying registry setting for: $registryPath -> $registryValueName = $valueToSet"
                        Set-ItemProperty -Path $registryPath -Name $registryValueName -Value $valueToSet -ErrorAction Stop
                        Write-Host "Registry setting applied successfully for: $registryPath -> $registryValueName = $valueToSet"
                    } catch {
                        Write-Error "Failed to set registry value for: $registryPath. Error: $_"
                    }
                } else {
                    Write-Warning "Registry path does not exist: $policyGroup. Skipping entry."
                }
            } else {
                # Skip unsupported entries with a warning
                Write-Warning "Skipping unsupported entry: $policyGroup"
            }
        }

    } catch {
        Write-Error "An error occurred while processing the audit policies. Error: $_"
    }
}

Set-AuditPolicy -ExcelFilePath "C:\Baseline\Reports\Baseline.xlsx" -SheetName "Policy Analyzer"

#auditpol /set /category:"Account Logon" /failure:enable /success:enable
#auditpol /set /category:"Detailed Tracking" /failure:disable /success:enable
#auditpol /set /subcategory:"Account Lockout" /failure:enable /success:disable
#auditpol /set /subcategory:"Group Membership" /failure:disable /success:enable
