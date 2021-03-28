<#
    Author: David Brook
    Date: 03/02/01

    Purpose: Install and/or detect the Visual Studio Code application install
#>

param (
    [ValidateSet('64bit','32bit','ARM64')]
    [String]$Architecture = '64bit',
    [ValidateSet('System','User')]
    [string]$InstallAs = 'System',
    [ValidateSet('Install','Uninstall',"Detect")]
    [string]$ExecutionType,
    [String]$DownloadPath = "$env:Temp\VSCodeInstaller"
)

$TranscriptFile = "$env:SystemRoot\Logs\Software\VSCode_Dynamic_Install.Log"
IF (-Not ($ExecutionType -Match "Detect")) {
    Start-Transcript -Path $TranscriptFile
}

### Set the AppArchitecture Variable based on the Architecture passed in via the command line or header
switch ($Architecture) {
    64bit { $global:AppArchitecture = "win32-x64" }
    32bit { $global:AppArchitecture = "win32" }
    ARM64 { $global:AppArchitecture = "win32-arm64" }
}

### If the InstallAs Command Line Parameter is set to User, append the App Architecture with -User for the API Query
switch ($InstallAs) {
    User { $global:AppArchitecture = $AppArchitecture + "-user" }
}

##############################################################################
########################## Application Detection #############################
##############################################################################
function Detect-Application {
    ### If the application installation exists in the Registry, Return that the application is installed and a $True Value, 
    IF (((Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"}).DisplayVersion -Match $LatestVersion) -or `
     ((Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"}).DisplayVersion -Match $LatestVersion))
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
    IF (-not (Test-Path $DownloadPath)) {
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
        Invoke-WebRequest -Usebasicparsing -URI $DownloadLink -Outfile "$DownloadPath\$InstallerName" -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to download application binaries"
    }
    # Once Downloaded, Install the application
    try {
        "Installing $AppName v$($LatestVersion)"
        Start-Process -FilePath "$DownloadPath\$InstallerName" -ArgumentList $InstallArguments -Wait
    }
    catch {
        Write-Error "Failed to Install $AppName, please check the transcript file ($TranscriptFile) for further details."
    }
}

function Uninstall-Application {
    try {
        "Uninstalling $AppName"
        IF (Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"} -ErrorAction SilentlyContinue) {
            $UninstallEXE = (Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"}).UninstallString
            Start-Process ($UninstallEXE.Trim('"')) -ArgumentList $UninstallArguments -Wait
        }      
        IF (Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"} -ErrorAction SilentlyContinue) {
            $UninstallEXE = (Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"}).UninstallString
            $UninstallArguments = "/Silent"
            Start-Process ($UninstallEXE.Trim('"')) -ArgumentList $UninstallArguments -Wait
        } 

    } catch {
        Write-Error "failed to Uninstall $AppName"
    }
}

##############################################################################
##################### Get the Information from the API #######################
##############################################################################
$URL = "https://update.code.visualstudio.com/api/update/$AppArchitecture/stable/version"
$RestResult = Invoke-RestMethod -Method GET -Uri $URL -ContentType 'application/json'

##############################################################################
########################## Set Required Variables ############################
##############################################################################
$LatestVersion = $RestResult.productVersion
$DownloadLink = $RestResult.url

##############################################################################
########################## Install/Uninstall Params ##########################
##############################################################################
$InstallerName = "VSCode" + $AppArchitecture + $LatestVersion + ".exe"
$InstallArguments = "/verysilent /mergetasks=!runcode,addcontextmenufiles,addcontextmenufolders,associatewithfiles,addtopath"
$UninstallArguments = "/VerySilent"
$UninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$UninstallKeyWow6432Node = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
$AppName = "Microsoft Visual Studio Code"

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
            Write-Error "Failed to Uninstall $AppName"
        }
    }
    Default {
        IF (!(Detect-Application)) {
            try {
                "The latest version is not installed, Starting Installation"
                Install-Application -ErrorAction Stop
            } catch {
                Write-Error "Failed to Install $AppName"
            }
        } ELSE {
            "The Latest Version of $AppName is already installed"
        }
    }
}

IF (-Not ($ExecutionType -Match "Detect")) {
    Stop-Transcript
}
