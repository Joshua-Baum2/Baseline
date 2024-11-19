#Query Security Log for Event ID 4625 - Failed login attempts
$failedLogins = Get-WinEvent -LogName Security -FilterXPath "*[System[EventID=4625]]" 

#Display results
$failedLogins | Select-Object TimeCreated, Id, Message | Format-Table -AutoSize