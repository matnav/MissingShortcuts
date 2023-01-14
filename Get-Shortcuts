<#
    .SYNOPSIS
    This script is used to get all the shortcuts on the desktop and exports them to a CSV file, with the option to exclude certain paths.
    
    .DESCRIPTION
    The script uses the Get-ChildItem cmdlet to recursively search the Start Menu folder for all
    shortcuts, and then uses the WScript.Shell COM object to retrieve properties of the shortcuts.
    The properties are then stored in a custom object, which is added to an array. The array is then 
    exported to a CSV file.
    
    .PARAMETER ExcludePaths
    The paths that you want to exclude when searching for shortcuts. These can be partial paths, and will match any path that contains the specified string.
    
    .PARAMETER FilePath
    The path where the CSV file will be exported to.
    
    .PARAMETER SearchPath
    The path where the script looks for the shortcuts
    
    .PARAMETER ShowOutput
    Show the output of the script on console or not
    
    .EXAMPLE
    .\Get-DesktopShortcuts.ps1 -ExcludePaths "Accessibility", "Accessories", "Administrative" -FilePath "C:\matnav\shortcuts.csv" -SearchPath "C:\ProgramData\Microsoft\Windows\Start Menu" -ShowOutput $true
    This will run the script and export the shortcuts to a CSV file, excluding any paths that contain the specified strings, looking for shortcuts in "C:\ProgramData\Microsoft\Windows\Start Menu" and Showing the output on console.
    
    .NOTES
    Author: Matthew Navarette
    Date Created: January 14, 2023
#>

param (
    [string[]]$ExcludePaths = @(),
    [string]$FilePath = "c:\matnav\shortcuts.csv",
    [string]$SearchPath = "C:\ProgramData\Microsoft\Windows\Start Menu",
    [bool]$
