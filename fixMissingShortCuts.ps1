# Script Name: Fix-MissingShortcuts.ps1
# Author: Matthew Navarette
# Date Created: January 14, 2023
# Purpose: Replaces missing shortcuts with a template file containing the shortcut files.

$ImportedCSV =  .{Invoke-WebRequest https://raw.githubusercontent.com/matnav/MissingShortcuts/main/shortcuts.csv -OutFile C:\path\shortcuts.csv 
Import-Csv -Path C:\path\shortcuts.csv}
$AppTest = Test-Path -Path $ImportedCSV.Target -PathType Leaf
#Restore Shortcuts to Public desktop
    $ComObj = New-Object -ComObject WScript.Shell
    $ShortCut = $ComObj.CreateShortcut("C:\Users\Public\Desktop\Chrome.lnk")
    $ShortCut.TargetPath = "C:\Program Files (x86)\Google Chrome\Chrome.exe"
    $ShortCut.Description = "Google Chrome"
   #The IconLocation line is optional    
    $shortcut.IconLocation = "C:\path\to\icon.ico"
    $ShortCut.FullName 
    $ShortCut.WindowStyle = 1
    $ShortCut.Save()

if ($AppTest -eq $True){

	if(!(Test-Path -Path $ImportedCSV.ShortcutFull))
{
ForEach-Object {
write-host "Replacing missing shortcut for" $MissingShortcut.ShortcutName

$ComObj = New-Object -ComObject WScript.Shell
		$ShortCut = $ComObj.CreateShortcut($MissingShortcut.ShortcutFull)
		$ShortCut.TargetPath = $MissingShortcut.Target
		$ShortCut.Description = $MissingShortcut.ShortcutName
        $ShortCut.WorkingDirectory = $MissingShortcut.WorkingDirectory
		$ShortCut.FullName 
		$ShortCut.WindowStyle = 1
		$ShortCut.Save()
}}
}else{
write-host "Something went wrong... :("
}


