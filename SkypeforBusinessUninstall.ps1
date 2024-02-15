# Uninstall File / Start Menu for Skype for Business
# Created by: Steven Rabago


Invoke-Command -ScriptBlock {
$ProgressPreference = "SilentlyContinue"


Write-Host "Closing Skype for Business (lync) and uninstalling..."

If ((get-process "lync" -ea SilentlyContinue) -eq $Null){ 
    echo "lync.exe is not running" 

    Get-ChildItem -Path "C:\Program Files\Microsoft Office\root\Office16\*lync*" -Recurse | Remove-Item
    Write-Host "Skype for Business (Lync) successfully uninstalled from C:\Program Files\Microsoft Office\root\Office16\"
}
else { 
    echo "lync.exe is running, closing"

    Stop-Process -Name lync -Force

    Get-ChildItem -Path "C:\Program Files\Microsoft Office\root\Office16\*lync*" -Recurse | Remove-Item
    Write-Host "Skype for Business (Lync) successfully uninstalled from C:\Program Files\Microsoft Office\root\Office16\"
}
Write-host ""


Write-Host "Closing Skype for Business Recording Manager (OcPubMgr) and uninstalling..."
If ((get-process "OcPubMgr" -ea SilentlyContinue) -eq $Null){ 
    echo "OcPubMgr.exe is not running" 

    Get-ChildItem -Path "C:\Program Files\Microsoft Office\root\Office16\*OcPubMgr*" -Recurse | Remove-Item
    Write-Host "Skype for Business Recording Manager (OcPubMgr) successfully uninstalled from C:\Program Files\Microsoft Office\root\Office16\"
}
else { 
    echo "OcPubMgr.exe is running, closing"

    Stop-Process -Name OcPubMgr -Force

    Get-ChildItem -Path "C:\Program Files\Microsoft Office\root\Office16\*OcPubMgr*" -Recurse | Remove-Item
    Write-Host "Skype for Business Recording Manager (OcPubMgr) successfully uninstalled from C:\Program Files\Microsoft Office\root\Office16\"
}
Write-host ""


$FileShortcut = "Skype for Business.lnk"
if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\") {
    Get-ChildItem -path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\" -recurse | Where-Object Name -like $FileShortcut | Remove-Item
    Write-Host "Skype for Business Start Menu Shortcut removed from C:\ProgramData\Microsoft\Windows\Start Menu\Programs\"
}
else {
    Write-host "The Skype for Business Start Menu shortcut does not exist!"
}
Write-host ""


$FileShortcut2 = "Skype for Business Recording Manager.lnk"
if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\") {
    Get-ChildItem -path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\" -recurse | Where-Object Name -like $FileShortcut2 | Remove-Item
    Write-Host "Skype for Business Recording Manager Start Menu Shortcut removed from C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office Tools\"
}
else {
    Write-host "The Skype for Business Recording Manager Start Menu shortcut does not exist!"
}
Write-host ""

if (Test-Path -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Lync") {
    Remove-Item -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Lync" -Recurse
    Write-Host "Lync successfully removed from Registry"

    # REG DELETE HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\15.0\Lync /F
}
else {
    Write-host "The Lync Registry Key does not exist!"
}





}