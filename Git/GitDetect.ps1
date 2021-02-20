$GitVer = "https://gitforwindows.org/"

IF (((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Git_is1\" -Name "DisplayVersion" -ErrorAction SilentlyContinue) -Match  (((Invoke-WebRequest $GitVer -UseBasicParsing).Links | Where-Object { $_.OuterHTML -Like "*Version*" }).Title.Split(' ')[1])) -or ((Get-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Git_is1\" -Name "DisplayVersion" -ErrorAction SilentlyContinue) -Match  (((Invoke-WebRequest $GitVer -UseBasicParsing).Links | Where-Object { $_.OuterHTML -Like "*Version*" }).Title.Split(' ')[1]))) {
    $True
}