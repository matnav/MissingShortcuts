# Script Name: Fix-MissingShortcuts.ps1
# Author: Matthew Navarette
# Date Created: January 14, 2023
# Purpose: Replaces missing shortcuts with a template file containing the shortcut files.

# Import CSV file containing shortcuts
$ImportedCSV = Invoke-WebRequest https://raw.githubusercontent.com/matnav/MissingShortcuts/main/shortcuts.csv -OutFile C:\path\shortcuts.csv 
Import-Csv -Path C:\path\shortcuts.csv

# Test if the target path of the shortcuts exist
$AppTest = Test-Path -Path $ImportedCSV.Target -PathType Leaf

# Create a shortcut on the Public Desktop
	#$ComObj = New-Object -ComObject WScript.Shell
	#$ShortCut = $ComObj.CreateShortcut("C:\path\to\shortcut.lnk")
	#$ShortCut.TargetPath = "C:\path\to\actual\program.exe"
	#$ShortCut.Description = "Program Name"
	#$ShortCut.IconLocation = "C:\path\to\icon.ico"
	#$ShortCut.FullName 
	#$ShortCut.WindowStyle = 1
	#$ShortCut.Save()

$MissingShortcuts = 0
# Check if the target path exist, if true then check if the shortcut exist
# if the shortcut doesn't exist, create it and keep count of missing shortcuts
if ($AppTest -eq $True){
    ForEach ($MissingShortcut in $ImportedCSV) {
        if(!(Test-Path -Path $MissingShortcut.ShortcutFull)) {
            write-host "Replacing missing shortcut for" $MissingShortcut.ShortcutName
            $ComObj = New-Object -ComObject WScript.Shell
            $ShortCut = $ComObj.CreateShortcut($MissingShortcut.ShortcutFull)
            $ShortCut.TargetPath = $MissingShortcut.Target
            $ShortCut.Description = $MissingShortcut.ShortcutName
            $ShortCut.WorkingDirectory = $MissingShortcut.WorkingDirectory
            $ShortCut.FullName 
            $ShortCut.WindowStyle = 1
            $ShortCut.Save()
            $MissingShortcuts++
        }
    }
    if($MissingShortcuts -eq 0) {
        write-host "No missing shortcuts were found"
    }
} else {
    write-host "Something went wrong... :("
}
