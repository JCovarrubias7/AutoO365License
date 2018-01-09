#############################################################################
# Author  : Jorge E. Covarrubias
# Website : 
# LinkedIn  : https://www.linkedin.com/in/jorge-e-covarrubias-973217141/
#
# Original script : https://gallery.technet.microsoft.com/office/DirSync-Timer-for-O365-c432dfc7
#
# Version   : 1.0
# Created   : 9/14/2017
# Modified  :
# 11/29/2017  - Adding heading to script
#             - To recap, added Staff and Student Functions to script.
#             - Added the count variables to keep count on how many times it has ran.
#             - Adjusted $toGetCST time from 5 to 6 due to daylights saving time.
#             - General clean up of code.
# 
#
# Purpose : This script will run continuously to check if new accounts are in Office365 and assign them licenses
#           according to the Title associated with the account.
#           The script is set to check every 10 minutes. This can be changed by changing the Start-Sleep to how
#           ever many seconds you would like.
#
#############################################################################

$pshost = Get-Host              # Get the PowerShell Host.
$pswindow = $pshost.UI.RawUI    # Get the PowerShell Host's UI.

$newsize = $pswindow.windowsize
$newsize.height = 6
$newsize.width = 55
$pswindow.windowsize = $newsize

$signature = @’
[DllImport("user32.dll")]
public static extern bool SetWindowPos(
    IntPtr hWnd,
    IntPtr hWndInsertAfter,
    int X,
    int Y,
    int cx,
    int cy,
    uint uFlags);
‘@
 
$type = Add-Type -MemberDefinition $signature -Name SetWindowPosition -Namespace SetWindowPos -Using System.Text -PassThru

$handle = (Get-Process -id $Global:PID).MainWindowHandle
$alwaysOnTop = New-Object -TypeName System.IntPtr -ArgumentList (-1)
$type::SetWindowPos($handle, $alwaysOnTop, 0, 0, 0, 0, 0x0003)

Import-Module MSonline
$O365Cred=Get-Credential
Set-ExecutionPolicy Unrestricted
Connect-MsolService -Credential $O365Cred
#Import-Module LyncOnlineConnector
$staffLicensedCount = 0
$studentLicenseCount = 0 
Clear-Host

Function StaffLocationandLicenses(){
		# Set Usage Location
		$staffLocationUsers = Get-MsolUser -All -Title "Staff" | where {$_.UsageLocation -eq $null}; 
		If ($staffLocationUsers) {
			$staffLocationUsers | foreach {Set-MsolUser -UserPrincipalName $_.UserPrincipalName -UsageLocation "US"}
		} Else {
		
		}
		
		Start-Sleep -Seconds 15
		
		# Set Licenses
		$staffUnlicensedUsers = Get-MsolUser -All -Title "Staff" -UnlicensedUsersOnly; 
		If ($staffUnlicensedUsers) {
			$staffUnlicensedUsers | foreach {Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses "sd104:STANDARDWOFFPACK_IW_FACULTY"}
			$Script:staffLicensedCount++
		} Else {
		
		}
}

Function StudentLocationandLicenses(){
		# Set Usage Location
		$studentLocationUsers = Get-MsolUser -All -Title "Student" | where {$_.UsageLocation -eq $null}; 
		If ($studentLocationUsers) {
			$studentLocationUsers | foreach {Set-MsolUser -UserPrincipalName $_.UserPrincipalName -UsageLocation "US"}
		} Else {
		
		}
		
		Start-Sleep -Seconds 30

		# Set Licenses
		$studentUnlicensedUsers = Get-MsolUser -All -Title "Student" -UnlicensedUsersOnly;
		If ($studentUnlicensedUsers) {
			$studentUnlicensedUsers | foreach {Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses "sd104:STANDARDWOFFPACK_IW_STUDENT"}
			$Script:studentLicenseCount++
		} Else {
		
		}
}

While($true)
{
    Clear-Host

    $time = Get-Date

    Write-Host "Script Last Checked DirSync at: " $time  -ForegroundColor Red
	  Write-Host "Times Executed : Staff = $staffLicensedCount Students = $studentLicenseCount" `r`n
    
    Start-Sleep -Seconds 2

    $lastDirSyncTimeUMT = Get-MsolCompanyInformation | select -ExpandProperty LastDirSyncTime

	#converts UMT to Central Standard Time = CST
	$toGetCST = New-TimeSpan -Hours 6
	#time set to by server to sync to 365
	$syncTimer = New-TimeSpan -Minutes 30

	$CST = ($lastDirSyncTimeUMT) - $toGetCST 
	$nextSyncTime = ($CST) + $syncTimer
    
    Write-Host "Last DirSync occurred at: " $CST -ForegroundColor Cyan
    Write-Host "Next DirSync will occur at: "  $nextSyncTime -ForegroundColor Cyan

	StaffLocationandLicenses
	StudentLocationandLicenses
	
    Start-Sleep -Seconds 600
    
} 