# Get last system boot time
$time = (Get-WinEvent -LogName System -MaxEvents 1 -FilterXPath "*[System[EventID=12]]" | Select-Object -First 1).TimeCreated
Write-Host -ForegroundColor Yellow $time
pause