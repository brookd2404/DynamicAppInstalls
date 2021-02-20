<#
    Author: David Brook
    Date: 17/02/2021

    Purpose: Install and/or detect the Git for Windows using the GitHub API
#>

param (
    [ValidateSet('64-bit','32-bit','ARM64')]
    [String]$Arch = '64-bit',
    [ValidateSet('Install','Uninstall',"Detect")]
    [string]$ExecutionType
)

function Get-LatestVersion {
    [String]$URL = "https://api.github.com/repos/git-for-windows/git/releases/latest"
    $WebResult = Invoke-RestMethod -Method GET -Uri $URL -ContentType "application/json" 
    
    $WebResult.name.split()[-1]

}

function Get-DownloadLink {
    [String]$URL = "https://api.github.com/repos/git-for-windows/git/releases/latest"
    $WebResult = Invoke-RestMethod -Method GET -Uri $URL -ContentType "application/json" 

    ($WebResult.assets | Where-Object {$_.Name -Match $EXEName}).browser_download_url
}

function Detect-GitForWindows {
    IF (((Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Git version*"}).DisplayVersion -Match $LatestVersion) -or ((Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Git version*"}).DisplayVersion -Match $LatestVersion))
    {
        $True
    }
}

function Install-GitForWindows {
    IF (!(Test-Path $DownloadPath)) {
        try {
            Write-Verbose "$DownloadPath Does not exist, Creating the folder"
            MKDIR $DownloadPath -ErrorAction Stop | Out-Null
        } catch {
            Write-Verbose "Failed to create folder $DownloadPath"
        }
    }
    
    try {
        Write-Verbose "Attempting client download"
        Invoke-WebRequest -Usebasicparsing -URI $DownloadLink -Outfile "$DownloadPath\$EXEName" -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to download RDInstaller"
    }
    
    try {
        "Installing Git for Windows $($LatestVersion)"
        Start-Process "$DownloadPath\$EXEName" -ArgumentList "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART" -Wait
    }
    catch {
        Write-Error "failed to Install Remote Desktop Client"
    }
}

function Uninstall-GitForWindows {
    try {
        $UninstallArgs = "/VERYSILENT /NORESTART"
        "Uninstalling Git for Windows"

        IF (Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Git version*"} -ErrorAction SilentlyContinue) {
            "Uninstall 64-Bit Git for Windows"
            $UninstallExe = (Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Git version*"}).UninstallString
            Start-Process $UninstallExe -ArgumentList $UninstallArgs -Wait
        }      
        
        IF (Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Git version*"} -ErrorAction SilentlyContinue) {
            "Uninstall 32-Bit Git for Windows" 
            $UninstallExe = (Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Git version*"}).UninstallString
            Start-Process $UninstallExe -ArgumentList $UninstallArgs -Wait
        } 

    } catch {
        Write-Error "failed to Uninstall Git for Windows"
    }
}

#If the latest version of the Remote Desktop application is not detected, Install it.
IF(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    $DownloadPath = "C:\Windows\Temp\GitInstaller\"
} ELSE {
    $DownloadPath = "$env:Temp\GitInstaller\"
}

$Global:UninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$Global:UninstallKeyWow6432Node = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
$global:LatestVersion = Get-LatestVersion
$global:DownloadLink = Get-DownloadLink
$global:EXEName = "Git-$LatestVersion-$Arch.exe"

switch ($ExecutionType) {
    Detect { 
        Detect-GitForWindows 
    }
    Uninstall {
        try {
                Unstall-GitForWindows -ErrorAction Stop
                "Uninstallation Complete"
        } 
        catch {
                Write-Error "Failed to Install Git for Windows"
        }
    }
    Default {
        IF (!(Detect-GitForWindows)) {
            try {
                "The latest version is not installed, Attempting install"
                Install-GitForWindows -ErrorAction Stop
                "Installation Complete"
            } catch {
                Write-Error "Failed to Install Git for Windows"
            }
        } ELSE {
            "The Latest Version is already installed"
        }
    }
}