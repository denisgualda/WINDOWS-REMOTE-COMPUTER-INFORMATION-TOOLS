#$Computer = "lt2ba19391"
#$Computer = "10.129.21.228"
$Computer = $args[0]

ForEach-Object { [wmi]($_.Dependent) }

$HWDocking = Get-WmiObject Win32_USBControllerDevice -Computername $Computer -ErrorAction Stop | ForEach-Object { [wmi]($_.Dependent) } | Where-Object {$_.name -like "USB\VID_17EF&PID_A38F" -or $_.DeviceID -like "*USB\VID_17EF&PID_A38F\*"} | select DeviceID
($HWDocking.deviceid -split "\\")[-1]



