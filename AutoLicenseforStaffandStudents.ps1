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
			$staffUnlicensedUsers | foreach {Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses "sd104:STANDARDWOFFPACK_IW_FACULTY","sd104:CLASSDASH_PREVIEW"}
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
			$studentUnlicensedUsers | foreach {Set-MsolUserLicense -UserPrincipalName $_.UserPrincipalName -AddLicenses "sd104:STANDARDWOFFPACK_IW_STUDENT","sd104:CLASSDASH_PREVIEW"}
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