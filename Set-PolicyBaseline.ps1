#Install-Module -Name ImportExcel
#Install-Module -Name ImportExcel -Scope CurrentUser

function Set-PolicyBaseline {
    [CmdletBinding()]
    param (
        [string]$ExcelFilePath = "C:\Baseline\Reports\Baseline.xlsx",
        [string]$SheetName = "Policy Analyzer"
    )
    
    if (-not(Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Error "The ImportExcel module is required. Install it by running: Install-Module -Name ImportExcel"
        return
    }

    try {
        # Import data from the specified Excel Sheet
        Write-Verbose "Reading data from Excel file: $ExcelFilePath"
        $baselineData = Import-Excel -Path $ExcelFilePath 

        foreach ($policy in $baselineData) {
            # Extract relevant fields
            $policyType = $policy.'Policy Type'
            $registryPath = $policy.'Policy Group or Registry Key'
            $registryValueName = $policy."Policy Setting"
            $expectedValue = $policy.Baseline

            # Check if the registry key exists
            if (-not (Test-Path -Path $registryPath)) {
                Write-Warning "Registry path $registryPath does not exist. Skipping Policy $registryValueName"
                continue
            }
            
            # Check if the current registry value matches the baseline
            $currentValue = (Get-ItemProperty -Path $registryPath -Name $registryValueName -ErrorAction SilentlyContinue).$registryValueName
            
            if ($currentValue -ne $expectedValue) {
                # Apply the baseline value if there is a mismatch
                Write-Host "Applying policy: $policyType (Current: $currentValue, Expected: $expectedValue)"
                try {
                    Set-ItemProperty -Path $registryPath -Name $registryValueName -Value $expectedValue -Credential .\vboxuser 
                    Write-Host "Successfully applied: $policyType"
                } catch {
                    Write-Error "Failed to apply policy: $policyType. Error: $_"
                }
            } else {
                Write-Host "Policy already compliant: $policyType"
            }
        } 
            } catch {
                Write-Error "An error occurred while processing the policy baseline. Error: $_"
    }
}

Set-PolicyBaseline -ExcelFilePath "C:\Baseline\Reports\Baseline.xlsx" 
