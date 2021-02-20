<#
    Author: David Brook
    Date: 03/02/01

    Purpose: Install and/or detect the Remote Desktop application install
#>

param (
    [ValidateSet('64-bit','32-bit','ARM64')]
    [String]$Arch = '64-bit',
    [ValidateSet('Install','Uninstall',"Detect")]
    [string]$ExecutionType
)

#Check if the logged in user is an admin or not, then set the Download path Accordingly 

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
function Detect-RemoteDesktop {
    IF (((Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Remote Desktop*"}).DisplayVersion -Match $LatestVersion) -or ((Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Remote Desktop*"}).DisplayVersion -Match $LatestVersion))
    {
        $True
    }
}
function Install-RemoteDesktop {
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
        Invoke-WebRequest -Usebasicparsing -URI $DownloadLink -Outfile "$DownloadPath\RemoteDesktop-$Arch.msi" -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to download RDInstaller"
    }
    
    try {
        "Installing Remote Desktop Client v$($LatestVersion)"
        Start-Process "MSIEXEC.exe" -ArgumentList "/I ""$DownloadPath\RemoteDesktop-$Arch.msi"" /qn /norestart /l* ""$DownloadPath\RDINSTALL$(get-Date -format yyyy-MM-dd).log""" -Wait
    }
    catch {
        Write-Error "failed to Install Remote Desktop Client"
    }
}

function Uninstall-RemoteDesktop {
    try {
        "Uninstalling Remote Desktop Client"

        IF (Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Remote Desktop*"} -ErrorAction SilentlyContinue) {
            "Uninstall 64-Bit Remote Desktop Client"
            $UninstallGUID = (Get-ChildItem -Path $UninstallKey | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Remote Desktop*"}).PSChildName
            $UninstallArgs = "/X " + $UninstallGUID + " /qn"
            Start-Process "MSIEXEC.EXE" -ArgumentList $UninstallArgs -Wait
        }      
        
        IF (Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Remote Desktop*"} -ErrorAction SilentlyContinue) {
            "Uninstall 32-Bit Remote Desktop Client" 
            $UninstallGUID = (Get-ChildItem -Path $UninstallKeyWow6432Node | Get-ItemProperty | Where-Object {$_.DisplayName -like "*Remote Desktop*"}).UninstallString
            $UninstallArgs = "/X " + $UninstallGUID + " /qn"
            Start-Process "MSIEXEC.EXE" -ArgumentList $UninstallArgs -Wait
        } 

    } catch {
        Write-Error "failed to Uninstall Remote Desktop Client"
    }
}

#If the latest version of the Remote Desktop application is not detected, Install it.
IF(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    $DownloadPath = "C:\Windows\Temp\RDInstaller\"
} ELSE {
    $DownloadPath = "$env:Temp\RDInstaller\"
}

$Global:UninstallKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$Global:UninstallKeyWow6432Node = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
$global:LatestVersion = ((Get-LatestVersion | Get-Unique | Sort-Object $_.LatestVersion)[0]).LatestVersion
$DownloadLink = Get-DownloadLink

switch ($ExecutionType) {
    Detect { 
        Detect-RemoteDesktop 
    }
    Uninstall {
        try {
                Uninstall-RemoteDesktop -ErrorAction Stop
                "Uninstallation Complete"
        } 
        catch {
            Write-Error "Failed to Install Remote Desktop Client"
        }
    }
    Default {
        IF (!(Detect-RemoteDesktop)) {
            try {
                "The latest version is not installed, Attempting install"
                Install-RemoteDesktop -ErrorAction Stop
                "Installation Complete"
            } catch {
                Write-Error "Failed to Install Remote Desktop Client"
            }
        } ELSE {
            "The Latest Version is already installed"
        }
    }
}