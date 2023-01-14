#
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
    .\Get-Shortcuts.ps1 -ExcludePaths "Accessibility", "Accessories", "Administrative" -FilePath "C:\matnav\shortcuts.csv" -SearchPath "C:\ProgramData\Microsoft\Windows\Start Menu" -ShowOutput $true
    This will run the script and export the shortcuts to a CSV file, excluding any paths that contain the specified strings, looking for shortcuts in "C:\ProgramData\Microsoft\Windows\Start Menu" and Showing the output on console.
    
    .NOTES
    Author: Matthew Navarette
    Date Created: January 14, 2023
#>

param (
    [string[]]$ExcludePaths = @(),
    [string]$FilePath = "c:\matnav\shortcuts.csv"
)

function Get-Shortcuts {
    # Define an array to store the shortcut objects
    $Shortcuts = @()

    # Get all the shortcuts in the Start Menu folder
    $AllShortcuts = Get-ChildItem -Recurse $SearchPath -Include *.lnk

    # Create a WScript.Shell object
    $Shell = New-Object -ComObject WScript.Shell

    # Loop through each shortcut
    foreach ($Shortcut in $AllShortcuts) {
        # Check if the shortcut path contains any of the excluded strings
        if ($ExcludePaths -notcontains ($Shortcut.DirectoryName | Select-String -SimpleMatch -Quiet)) {
            # Create a custom object to store the shortcut properties
            $Properties = @{
                ShortcutName     = $Shortcut.Name
                ShortcutFull     = $Shortcut.FullName
                ShortcutPath     = $shortcut.DirectoryName
                Target           = $Shell.CreateShortcut($Shortcut).targetpath
                WorkingDirectory = $Shell.CreateShortcut($Shortcut).WorkingDirectory
            }
            $Shortcuts += New-Object PSObject -Property $Properties
            if($ShowOutput){
                Write-Output "ShortcutName: $($Properties.ShortcutName) ShortcutFull: $($Properties.ShortcutFull) ShortcutPath: $($Properties.ShortcutPath) Target: $
