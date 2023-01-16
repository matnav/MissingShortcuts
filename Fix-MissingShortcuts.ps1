# Script Name: Fix-MissingShortcuts.ps1
# Author: Matthew Navarette
# Date Created: January 14, 2023
# Purpose: Replaces missing shortcuts with a template file containing the shortcut files.

# Define the location of the folder where the CSV file will be stored
$folder = 'C:\matnav\'
# Define the URL of the CSV file containing the shortcuts
$hostedcsv = 'https://raw.githubusercontent.com/matnav/MissingShortcuts/main/shortcuts.csv'

# Check if the folder exists, if not create the folder
if (!(Test-Path -Path $folder)) {New-Item -ItemType Directory -Path $folder}

# Import CSV file containing shortcuts
# Download the CSV file from the specified URL and save it to the defined folder
# then import the CSV file into a variable
$ImportedCSV = .{Invoke-WebRequest $hostedcsv OutFile "$folder\shortcuts.csv"
Import-Csv -Path "$folder\shortcuts.csv"}

# Test if the application is installed
# Check if the target path specified in the CSV file exists
$AppTest = Test-Path -Path $ImportedCSV.Target -PathType Leaf

# Create a shortcut on the Public Desktop
# This block of code uses the WScript.Shell COM object to create a shortcut on the Public Desktop. 
# Please note that the Test-Path method doesn't work correctly with hidden files, so this is a quick workaround. 
# To use it, simply uncomment the code and change the target path, description and icon location to suit your needs.
	#$ComObj = New-Object -ComObject WScript.Shell
		#$ShortCut = $ComObj.CreateShortcut("C:\path\to\shortcut.lnk")
		#$ShortCut.TargetPath = "C:\path\to\program.exe"
		#$ShortCut.Description = "Program Name"
		#$ShortCut.IconLocation = "C:\path\to\icon.ico"
		#$ShortCut.FullName 
		#$ShortCut.WindowStyle = 1
		#$ShortCut.Save()

# Counter for missing shortcuts
$MissingShortcuts = 0

# Check if the target path exist, if true then check if the shortcut exist
# if the shortcut doesn't exist, create it and keep count of missing shortcuts
if ($AppTest -eq $True){
    ForEach ($MissingShortcut in $ImportedCSV) {
        if(!(Test-Path -Path $MissingShortcut.ShortcutFull)) {
            write-host "Replacing missing shortcut for " $MissingShortcut.ShortcutName
            # Create a new shortcut using the WScript.Shell COM object
            $ComObj = New-Object -ComObject WScript.Shell
            $ShortCut = $ComObj.CreateShortcut($MissingShortcut.ShortcutFull)
            # Set the target path, description, and working directory of the shortcut
            $ShortCut.TargetPath = $MissingShortcut.Target
            $ShortCut.Description = $MissingShortcut.ShortcutName
            $ShortCut.WorkingDirectory = $MissingShortcut.WorkingDirectory
            # Save the shortcut
            $ShortCut.FullName 
            $ShortCut.WindowStyle = 1
            $ShortCut.Save()
            # Increase the missing shortcuts counter
            $MissingShortcuts++
        }
    }
    # Check if any missing shortcuts were found
    if($MissingShortcuts -eq 0) {
        write-host "No missing shortcuts were found"
    }
} else {
    write-host "Something went wrong... :("
}

$files = Get-ChildItem -Path $folder
#Check if the folder only contains the file 'shortcuts.csv' or if it's empty, then delete the folder
if ($files.Count -eq 0 -or ($files.Count -eq 1 -and $files[0].Name -eq 'shortcuts.csv')) {
    Remove-Item -Path $folder -Recurse -Force
}
