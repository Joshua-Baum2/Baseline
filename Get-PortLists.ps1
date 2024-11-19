function Get-PortLists {
    param (
        [string]$CsvFilePath1 = "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports\service-names-port-numbers.csv",
        [string]$CsvFilePath2 = "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports\OpenPorts.csv",
        [string]$ComparedOutputFile = "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports\PortComparisonReport.html",
        [string]$Csv1OutputFile = "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports\CsvPortsReport1.csv",
        [string]$Csv2OutputFile = "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports\CsvPortsReport2.csv"
    )

    # Read the first CSV file
    try {
        $csvData1 = Import-Csv -Path $CsvFilePath1
        $csvPorts1 = $csvData1 | Select-Object 'Port Number', 'Service Name', 'Transport Protocol', 'Description'
        # Save selected columns from the first CSV file to a new CSV file
        $csvPorts1 | Export-Csv -Path $Csv1OutputFile -NoTypeInformation -Force
        Write-Host "`nPorts from CSV File 1 saved to $Csv1OutputFile"
    } catch {
        Write-Warning "Failed to read or parse CSV File 1: $_"
        return
    }

    # Read the second CSV file
    try {
        $csvData2 = Import-Csv -Path $CsvFilePath2
        $csvPorts2 = $csvData2 | Select-Object @{Name='Port Number'; Expression={$_.'LocalPort'}}, 'OwningProcess', 'ProcessName', 'Description'
        # Save selected columns from the second CSV file to a new CSV file
        $csvPorts2 | Export-Csv -Path $Csv2OutputFile -NoTypeInformation -Force
        Write-Host "`nPorts from CSV File 2 saved to $Csv2OutputFile"
    } catch {
        Write-Warning "Failed to read or parse CSV File 2: $_"
        return
    }

    # Convert port numbers to integers for comparison, filtering out invalid values
    $csvPorts1Filtered = $csvPorts1 | Where-Object { [int]::TryParse($_.'Port Number', [ref]$null) }
    $csvPorts2Filtered = $csvPorts2 | Where-Object { [int]::TryParse($_.'Port Number', [ref]$null) }

    # Find matching ports
    $matchingPorts = $csvPorts1Filtered | Where-Object {
        $port = [int]$_.‘Port Number’
        $csvPorts2Filtered | Where-Object { [int]$_.‘Port Number’ -eq $port }
    }

    # Find non-matching ports from csvPorts2
    $nonMatchingPorts = $csvPorts2Filtered | Where-Object {
        $port = [int]$_.‘Port Number’
        -not ($csvPorts1Filtered | Where-Object { [int]$_.‘Port Number’ -eq $port })
    }

    # Create HTML report with matching ports and non-matching ports highlighted
    $htmlContent = @"
<html>
<head>
    <style>
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid black; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        .highlight { background-color: #ffcccc; }
    </style>
</head>
<body>
    <h2>Port Comparison Report</h2>
    <h3>Matching Ports</h3>
    <table>
        <tr>
            <th>Port Number</th>
            <th>Service Name</th>
            <th>Transport</th>
            <th>Description</th>
        </tr>
"@

    # Add matching ports to the HTML table
    foreach ($port in $matchingPorts) {
        $htmlContent += "<tr><td>$($port.'Port Number')</td><td>$($port.'Service Name')</td><td>$($port.'Transport Protocol')</td><td>$($port.Description)</td></tr>"
    }

    $htmlContent += @"
    </table>
    <h3>Non-Matching Ports (from CSV File 2)</h3>
    <table>
        <tr>
            <th>Port Number</th>
            <th>Service Name</th>
            <th>Transport</th>
            <th>Description</th>
        </tr>
"@

    # Add non-matching ports to the HTML table with highlighting
    foreach ($port in $nonMatchingPorts) {
        $htmlContent += "<tr class='highlight'><td>$($port.'Port Number')</td><td>$($port.OwningProcess)</td><td>$($port.ProcessName)</td><td>$($port.Description)</td></tr>"
    }

    $htmlContent += @"
    </table>
</body>
</html>
"@

    # Save the HTML report
    try {
        $htmlContent | Out-File -FilePath $ComparedOutputFile -Force
        Write-Host "`nComparison report saved to $ComparedOutputFile"
    } catch {
        Write-Warning "Failed to write the comparison report: $_"
    }
}

# File export
Compare-PortLists -CsvFilePath1 "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports\service-names-port-numbers.csv" -CsvFilePath2 "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports\OpenPorts.csv" -ComparedOutputFile "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports\PortComparisonReport.html" -Csv1OutputFile "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports\CsvPortsReport1.csv" -Csv2OutputFile "C:\Users\joshu\Documents\Cyber Security\Baseline\Reports\CsvPortsReport2.csv"



