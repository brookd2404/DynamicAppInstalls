<#
.SYNOPSIS
  This is a script to Dynamically Detect, Install and Uninstall the Git for Windows Client.

  https://gitforwindows.org/

.DESCRIPTION
  Use this script to detect, install or uninstall the Git for Windows client.

.PARAMETER Arch
    Select the architecture you would like to install, select from the following
    - 64-bit (Default)
    - 32-bit
    - ARM64

.PARAMETER ExecutionType
    Select the Execution type, this determines if you will be detecting, installing uninstalling the application.
    
    The options are as follows;
    - Install (Default)
    - Detect
    - Uninstall

.Parameter DownloadPath
    The location you would like the downloaded installer to go. 

    Default: $env:TEMP\GitInstall

.NOTES
  Version:        1.0
  Author:         David Brook
  Creation Date:  21/02/2021
  Purpose/Change: Initial script development
  
#>

param (
    [ValidateSet('64-bit','32-bit','ARM64')]
    [String]$Arch = '64-bit',
    [ValidateSet('Install','Uninstall',"Detect")]
    [string]$ExecutionType = "Detect",
    [string]$DownloadPath = "$env:Temp\GitInstaller\",
    [string]$GITPAC 
)

$TranscriptFile = "$env:SystemRoot\Logs\Software\GitForWindows_Dynamic_Install.Log"
Start-Transcript -Path $TranscriptFile

##############################################################################
########################## Application Detection #############################
##############################################################################
function Detect-Application {
    IF (((Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$DetectionString*"}).DisplayVersion -Match $LatestVersion) -or ((Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$DetectionString*"}).DisplayVersion -Match $LatestVersion))
    {
        Write-Output "$AppName is installed"
        $True
    }
}

##############################################################################
################## Application Installation/Uninstallation ###################
##############################################################################
function Install-Application {
    # If the Download Path does not exist, Then try and crate it. 
    IF (!(Test-Path $DownloadPath)) {
        try {
            Write-Verbose "$DownloadPath Does not exist, Creating the folder"
            New-Item -Path $DownloadPath -ItemType Directory -ErrorAction Stop | Out-Null
        } catch {
            Write-Verbose "Failed to create folder $DownloadPath"
        }
    }
    # Once the folder exists, download the installer
    try {
        Write-Verbose "Downloading Application Binaries for $AppName"
        Invoke-WebRequest -Usebasicparsing -URI $DownloadLink -Outfile "$DownloadPath\$EXEName" -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to download application binaries"
    }
    # Once Downloaded, Install the application
    try {
        "Installing $AppName $($LatestVersion)"
        Start-Process "$DownloadPath\$EXEName" -ArgumentList $InstallArgs -Wait
    }
    catch {
        Write-Error "Failed to Install $AppName, please check the transcript file ($TranscriptFile) for further details."
    }
}

function Uninstall-Application {
    try {
               
        IF (Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$DetectionString*"} -ErrorAction SilentlyContinue) {
            "Uninstalling $AppName"
            $UninstallExe = (Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$DetectionString*"}).UninstallString
            Start-Process $UninstallExe -ArgumentList $UninstallArgs -Wait
        }      
        
        IF (Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$DetectionString*"} -ErrorAction SilentlyContinue) {
            "Uninstalling $AppName" 
            $UninstallExe = (Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$DetectionString*"}).UninstallString
            Start-Process $UninstallExe -ArgumentList $UninstallArgs -Wait
        } 

    } catch {
        Write-Error "failed to Uninstall $AppName"
    }
}


##############################################################################
##################### Get the Information from the API #######################
##############################################################################
[String]$GitHubURI = "https://api.github.com/repos/git-for-windows/git/releases/latest"
IF ($GITPAC) {
    $RestResult = Invoke-RestMethod -Method GET -Uri $GitHubURI -ContentType "application/json" -Headers @{Authorization = "token $GITPAC"}
} ELSE {
    $RestResult = Invoke-RestMethod -Method GET -Uri $GitHubURI -ContentType "application/json"
}

##############################################################################
########################## Set Required Variables ############################
##############################################################################
$LatestVersion = $RestResult.name.split()[-1]
$EXEName = "Git-$LatestVersion-$Arch.exe"
$DownloadLink = ($RestResult.assets | Where-Object {$_.Name -Match $EXEName}).browser_download_url

##############################################################################
########################## Install/Uninstall Params ##########################
##############################################################################
$UninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$UninstallKeyWow6432Node = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
$DetectionString = "Git version"
$UninstallArgs = "/VERYSILENT /NORESTART"
$InstallArgs = "/SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART"
$AppName = "Git For Windows"


##############################################################################
############################# Do the Business ################################
##############################################################################
switch ($ExecutionType) {
    Detect { 
        Detect-Application 
    }
    Uninstall {
        try {
                Uninstall-Application -ErrorAction Stop
                "Uninstallation Complete"
        } 
        catch {
                Write-Error "Failed to Install $AppName"
        }
    }
    Default {
        IF (!(Detect-Application)) {
            try {
                "The latest version is not installed, Attempting install"
                Install-Application -ErrorAction Stop
                "Installation Complete"
            } catch {
                Write-Error "Failed to Install $AppName"
            }
        } ELSE {
            "The Latest Version is already installed"
        }
    }
}

Stop-Transcript