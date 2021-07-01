<#
.SYNOPSIS
  This is a script to Dynamically Detect, Install and Uninstall the Microsoft Remote Desktop Client for Windows.
  
  https://docs.microsoft.com/en-us/windows-server/remote/remote-desktop-services/clients/windowsdesktop

.DESCRIPTION
  Use this script to detect, install or uninstall the Microsoft Remote Desktop client for Windows

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

    Default: $env:TEMP\RDInstaller

.NOTES
  Version:        1.2
  Author:         David Brook
  Creation Date:  21/02/2021
  Purpose/Change: Initial script development
  
#>

param (
    [ValidateSet('64-bit','32-bit','ARM64')]
    [String]$Arch = '64-bit',
    [ValidateSet('Install','Uninstall',"Detect")]
    [string]$ExecutionType,
    [string]$DownloadPath = "$env:Temp\RDInstaller\"
)

function Get-LatestVersion {
    [String]$URL = "https://docs.microsoft.com/en-us/windows-server/remote/remote-desktop-services/clients/windowsdesktop-whatsnew"
    $WebResult = Invoke-WebRequest -Uri $URL -UseBasicParsing 
    $WebResultHTML = $WebResult.RawContent

    $HTML = New-Object -Com "HTMLFile"
    $HTML.IHTMLDocument2_write($WebResultHTML)

    $Tables = @($html.all.tags('table'))
    $LatestVer = $null
    [System.Collections.ArrayList]$LatestVer = New-Object -TypeName psobject
    for ($i = 0; $i -le $tables.count; $i++) {
        $table = $tables[0] 
        $titles = @()
        $rows = @($table.Rows)
        ## Go through all of the rows in the table
        foreach ($row in $rows) {
            $cells = @($row.Cells)
    
            ## If we've found a table header, remember its titles
            if ($cells[0].tagName -eq "TH") {
                $titles = @($cells | ForEach-Object {
                        ("" + $_.InnerText).Trim()
                    })
                continue
            }
            $resultObject = [Ordered] @{}
            $counter = 0
            foreach ($cell in $cells) {
                $title = $titles[$counter]
                if (-not $title) { continue }
                $resultObject[$title] = ("" + $cell.InnerText).Trim()
                $Counter++
            }
            #$Version_Data = @()
            $Version_Data = [PSCustomObject]@{
                'LatestVersion'          = $resultObject.'Latest version'
            }
            $LatestVer.Add($Version_Data) | Out-null 
        }
    }
    $LatestVer
}
function Get-DownloadLink {
    $URL = "https://docs.microsoft.com/en-us/windows-server/remote/remote-desktop-services/clients/windowsdesktop"
    $WebResult = Invoke-WebRequest -Uri $URL -UseBasicParsing 
    $WebResultHTML = $WebResult.RawContent

    $HTML = New-Object -Com "HTMLFile"
    $HTML.IHTMLDocument2_write($WebResultHTML)

    ($HTML.links | Where-Object {$_.InnerHTMl -Like "*$Arch*"}).href        
}
function Detect-Application {
    IF (((Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"}).DisplayVersion -Match $LatestVersion) -or ((Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"}).DisplayVersion -Match $LatestVersion))
    {
        $True
    }
}
function Install-Application {
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
        Invoke-WebRequest -Usebasicparsing -URI $DownloadLink -Outfile "$DownloadPath\$InstallerName" -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to download $AppName"
    }
    
    try {
        "Installing $AppName v$($LatestVersion)"
        Start-Process "MSIEXEC.exe" -ArgumentList "/I ""$DownloadPath\$InstallerName"" /qn /norestart /l* ""$DownloadPath\RDINSTALL$(get-Date -format yyyy-MM-dd).log""" -Wait
    }
    catch {
        Write-Error "failed to Install $AppName"
    }
}

function Uninstall-Application {
    try {
        "Uninstalling $AppName"

        IF (Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"} -ErrorAction SilentlyContinue) {
            "Uninstalling $AppName"
            $UninstallGUID = (Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"}).PSChildName
            $UninstallArgs = "/X " + $UninstallGUID + " /qn"
            Start-Process "MSIEXEC.EXE" -ArgumentList $UninstallArgs -Wait
        }      
        
        IF (Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"} -ErrorAction SilentlyContinue) {
            "Uninstalling $AppName" 
            $UninstallGUID = (Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*$AppName*"}).UninstallString
            $UninstallArgs = "/X " + $UninstallGUID + " /qn"
            Start-Process "MSIEXEC.EXE" -ArgumentList $UninstallArgs -Wait
        } 

    } catch {
        Write-Error "failed to Uninstall $AppName"
    }
}

$UninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$UninstallKeyWow6432Node = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
$LatestVersion = ((Get-LatestVersion | Get-Unique | Sort-Object $_.LatestVersion)[0]).LatestVersion
$InstallerName = "RemoteDesktop-$LatestVersion-$Arch.msi"
$AppName = "Remote Desktop"
$DownloadLink = Get-DownloadLink

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
            "The Latest Version ($LatestVersion) of $AppName is already installed"
        }
    }
}