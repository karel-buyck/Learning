#https://blogs.msdn.microsoft.com/cloud_solution_architect/2017/01/03/move-the-azure-temporary-disk-to-a-different-drive-letter-on-windows-server/
#To disable the Windows page file, we use “gwmi win32_pagefilesetting” which uses WMI to first check if the page file is enabled or not. If it is, we use this script to delete it and restart the VM:

=================================

#Script 1

gwmi win32_pagefilesetting
$pf=gwmi win32_pagefilesetting
$pf.Delete()
Restart-Computer –Force

==============================================

#Script 2

Get-Partition -DriveLetter "D" | Set-Partition -NewDriveLetter "T"
Set-WMIInstance -Class Win32_PageFileSetting -Arguments @{ Name = "T:\pagefile.sys"; MaximumSize = 0; }

#Arg Maximsize = 0  zal de config op system managed plaatsen, is best practice  

==============================

$TempDriveLetter = "T"
$drive = Get-WmiObject -Class win32_volume -Filter “DriveLetter = '$TempDriveLetter'”


Set-WMIInstance -Class Win32_PageFileSetting -Arguments @{ Name = "T:\pagefile.sys"; MaximumSize = 0; }
Restart-Computer -Force

Get-Partition -DriveLetter "D" | Set-Partition -NewDriveLetter $TempDriveLetter
$TempDriveLetter = "T"
$drive = Get-WmiObject -Class win32_volume -Filter “DriveLetter = '$TempDriveLetter'”




