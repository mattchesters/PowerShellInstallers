$redist = $PSScriptRoot + "\redist"

## Enable dotNet 3 for Windows 10
Import-Module DISM
$NetFx3Feature = Get-WindowsCapability -Online -Name NetFx3~~~~
if ($NetFx3Feature.State -ne "Installed"){
    Write-Host "Installing NetFx3 feature"
    $NetFx3Install = Add-WindowsCapability –Online -Name NetFx3~~~~
    Start-Sleep 5
    if ($NetFx3Feature.State -eq "Installed"){
        Write-Host "Successfully installed NetFx3 feature"
        } else {
        Write-Host "Failed to install NetFx3 feature"
        }
    } else {
    Write-Host "NetFx3 feature detected"
    }

##  dotNet 4.7 install
Write-Host "Trying Microsoft .NET Framework 4.6 (Offline Installer)"
Start-Process -Wait -FilePath $redist\NDP46-KB3045557-x86-x64-AllOS-ENU.exe -ArgumentList "/q /norestart"

##  2005 VC++ install
$2005Exists = Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{710f4c1c-cc18-4c49-8cbf-51240c89a1a2}'
if ($2005Exists){
    Write-Host "Detected Microsoft Visual C++ 2005 Redistributable (x86)"
    } else {
    Write-Host "Attempting Microsoft Visual C++ 2005 Redistributable (x86)"
    Start-Process -Wait -FilePath $redist\vcredist2005_x86.exe -ArgumentList "/Q"
    }

##  2010 VC++ install
$2010Exists = Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{F0C3E5D1-1ADE-321E-8167-68EF0DE699A5}'
if ($2010Exists){
    Write-Host "Detected Microsoft Visual C++ 2010 Redistributable (x86)"
    } else {
    Write-Host "Attempting Microsoft Visual C++ 2010 Redistributable (x86)"
    Start-Process -Wait -FilePath $redist\vcredist2010_x86.exe -ArgumentList "/q /norestart"
    }

##  2013 VC++ install
$2013Exists = Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{13A4EE12-23EA-3371-91EE-EFB36DDFFF3E}'
if ($2013Exists){
    Write-Host "Detected Microsoft Visual C++ 2013 Redistributable (x86) - 12.0.21005"
    } else {
    Write-Host "Attempting Microsoft Visual C++ 2013 Redistributable (x86) - 12.0.21005"
    Start-Process -Wait -FilePath $redist\vcredist2013_x86.exe -ArgumentList "/install /quiet /norestart /log %TEMP%\vcredist_2013_x86.log"
    }

##  2015 VC++ install
$2015Exists = Test-Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{282975d8-55fe-4991-bbbb-06a72581ce58}'
if ($2015Exists){
    Write-Host "Detected Microsoft Visual C++ 2015 Redistributable (x86)"
    } else {
    Write-Host "Attempting Microsoft Visual C++ 2015 Redistributable (x86) - 14.0.23026"
    Start-Process -Wait -FilePath $redist\vc_redist.2015.x86.exe -ArgumentList "/q /norestart"
    }

## Install Accounts Dataset Manager
Write-Host "Installing Accountants Dataset Manager"
$ADMArgs = @(
    "/i "
    "`"$PSScriptRoot\Accountants Dataset Manager.msi`""
    "/quiet"
    "/norestart"
 )
$ADM = Start-Process msiexec.exe -Wait -ArgumentList $ADMArgs
Write-Host $ADM.ExitCode

## Install AccountsV24
$v24ODBC_x64Args = @(
    "/i "
    "`"$PSScriptRoot\AccountsV24\packages\Sage50Accounts_ODBC_x64.msi`""
    "/quiet"
    "/norestart"
 )
$v24DataAccessArgs = @(
    "/i "
    "`"$PSScriptRoot\AccountsV24\packages\Sage50Accounts_DataAccess.msi`""
    "/quiet"
    "/norestart"
 )
$v24ReportPackArgs = @(
    "/i "
    "`"$PSScriptRoot\AccountsV24\packages\Sage50Accounts_ReportPack.msi`""
    "/quiet"
    "/norestart"
 )
$v24ClientArgs = @(
    "/i "
    "`"$PSScriptRoot\AccountsV24\packages\Sage50Accounts_Client.msi`""
    "/quiet"
    "/norestart"
 )
Write-Host "Installing v24 ODBC"
$v24ODBC_x64 = Start-Process msiexec.exe -Wait -ArgumentList $v24ODBC_x64Args
Write-Host $v24ODBC_x64.ExitCode
Write-Host "Installing v24 DataAccess"
$v24DataAccess = Start-Process msiexec.exe -Wait -ArgumentList $v24DataAccessArgs
Write-Host $v24DataAccess.ExitCode
Write-Host "Installing v24 ReportPack"
$v24ReportPack = Start-Process msiexec.exe -Wait -ArgumentList $v24ReportPackArgs
Write-Host $v24ReportPack.ExitCode
Write-Host "Installing v24 Client"
$v24Client = Start-Process msiexec.exe -Wait -ArgumentList $v24ClientArgs
Write-Host $v24Client.ExitCode

## Install AccountsV25
$v25ODBC_x64Args = @(
    "/i "
    "`"$PSScriptRoot\AccountsV25\packages\Sage50Accounts_ODBC_x64.msi`""
    "/quiet"
    "/norestart"
 )
$v25DataAccessArgs = @(
    "/i "
    "`"$PSScriptRoot\AccountsV25\packages\Sage50Accounts_DataAccess.msi`""
    "/quiet"
    "/norestart"
 )
$v25ReportPackArgs = @(
    "/i "
    "`"$PSScriptRoot\AccountsV25\packages\Sage50Accounts_ReportPack.msi`""
    "/quiet"
    "/norestart"
 )
$v25ClientArgs = @(
    "/i "
    "`"$PSScriptRoot\AccountsV25\packages\Sage50Accounts_Client.msi`""
    "/quiet"
    "/norestart"
 )
Write-Host "Installing v25 ODBC"
$v25ODBC_x64 = Start-Process msiexec.exe -Wait -ArgumentList $v25ODBC_x64Args
Write-Host $v25ODBC_x64.ExitCode
Write-Host "Installing v25 DataAccess"
$v25DataAccess = Start-Process msiexec.exe -Wait -ArgumentList $v25DataAccessArgs
Write-Host $v25DataAccess.ExitCode
Write-Host "Installing v25 ReportPack"
$v25ReportPack = Start-Process msiexec.exe -Wait -ArgumentList $v25ReportPackArgs
Write-Host $v25ReportPack.ExitCode
Write-Host "Installing v25 Client"
$v25Client = Start-Process msiexec.exe -Wait -ArgumentList $v25ClientArgs
Write-Host $v25Client.ExitCode

## Install AccountsV26
$v26ODBC_x64Args = @(
    "/i "
    "`"$PSScriptRoot\AccountsV26\packages\Sage50Accounts_ODBC_x64.msi`""
    "/quiet"
    "/norestart"
 )
$v26DataAccessArgs = @(
    "/i "
    "`"$PSScriptRoot\AccountsV26\packages\Sage50Accounts_DataAccess.msi`""
    "/quiet"
    "/norestart"
 )
$v26ReportPackArgs = @(
    "/i "
    "`"$PSScriptRoot\AccountsV26\packages\Sage50Accounts_ReportPack.msi`""
    "/quiet"
    "/norestart"
 )
$v26ClientArgs = @(
    "/i "
    "`"$PSScriptRoot\AccountsV26\packages\Sage50Accounts_Client.msi`""
    "/quiet"
    "/norestart"
 )
Write-Host "Installing v26 ODBC"
$v26ODBC_x64 = Start-Process msiexec.exe -Wait -ArgumentList $v26ODBC_x64Args
Write-Host $v26ODBC_x64.ExitCode
Write-Host "Installing v26 DataAccess"
$v26DataAccess = Start-Process msiexec.exe -Wait -ArgumentList $v26DataAccessArgs
Write-Host $v26DataAccess.ExitCode
Write-Host "Installing v26 ReportPack"
$v26ReportPack = Start-Process msiexec.exe -Wait -ArgumentList $v26ReportPackArgs
Write-Host $v26ReportPack.ExitCode
Write-Host "Installing v26 Client"
$v26Client = Start-Process msiexec.exe -Wait -ArgumentList $v26ClientArgs
Write-Host $v26Client.ExitCode


Write-Host "Complete"