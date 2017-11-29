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
$O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSSession $O365Session -AllowClobber
Set-ExecutionPolicy Unrestricted
Connect-MsolService -Credential $O365Cred
#Import-Module LyncOnlineConnector
Clear-Host


While($true)
{
    Clear-Host

    $t = Get-Date

    Write-Host "DIR SYNC TIMER - Last Run at: " $t  -ForegroundColor Red `r`n
    
    Start-Sleep -Seconds 2

    $x = Get-MsolCompanyInformation | select -ExpandProperty LastDirSyncTime

    $y = New-TimeSpan -Hours 4
    $y2 = New-TimeSpan -Hours 3

    $z = ($x) - $y 
    $z2 = ($z) + $y2

    
    Write-Host "Last DirSync occurred at: " $z -ForegroundColor Cyan `r`n
    Write-Host "Next DirSync will occur at: "  $z2 -ForegroundColor Cyan


    Start-Sleep -Seconds 600
    
    
      
} 