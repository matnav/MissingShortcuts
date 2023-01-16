<#
    .SYNOPSIS
    This script is used to collect all the shortcuts in the Start Menu and export them to a CSV file for replication or transfer. The script also includes the option to exclude certain paths from the search.
    
    .DESCRIPTION
    The script uses the Get-ChildItem cmdlet to recursively search the Start Menu folder for all shortcuts and also checks if the shortcut is accessible or not.
    Then it uses the WScript.Shell COM object to retrieve properties of the shortcuts. 
    The properties are then stored in a custom object, which is added to an array. 
    The array is then exported to a CSV file. This file can then be used to replicate the shortcuts on another computer or transfer them to another location.
    
    .PARAMETER ExcludePaths
    The paths that you want to exclude when searching for shortcuts. These can be partial paths, and will match any path that contains the specified string.
    
    .PARAMETER FilePath
    The path where the CSV file will be exported to.
    
    .PARAMETER SearchPath
    The path where the script looks for the shortcuts

    .PARAMETER CSVHeader
    If set to $True, the first line of the CSV file will contain headers for the columns.
    
    .EXAMPLE
    .\Get-Shortcuts.ps1 -ExcludePaths "Accessibility", "Accessories", "Administrative" -FilePath "C:\shortcuts.csv" -SearchPath "C:\ProgramData\Microsoft\Windows\Start Menu" -CSVHeader $True
    This will run the script and export the shortcuts to a CSV file, excluding any paths that contain the specified strings, looking for shortcuts in "C:\ProgramData\Microsoft\Windows\Start Menu" and adding headers in the first line of the CSV file.
    
    .NOTES
    Author: Matthew Navarette
    Date Created: January 14, 2023
    #>

param (
    [string[]]$ExcludePaths = @(),
    [string]$FilePath = "C:\shortcuts.csv",
    [string]$SearchPath = "C:\ProgramData\Microsoft\Windows\Start Menu",
    [bool]$CSVHeader = $True
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
            try {
                # Create a custom object to store the shortcut properties
                $Properties = @{
                    ShortcutName = $Shortcut.Name
                    ShortcutFull = $Shortcut.FullName
                    ShortcutPath = $shortcut.DirectoryName
                    Target = $Shell.CreateShortcut($Shortcut).targetpath
                    WorkingDirectory = $Shell.CreateShortcut($Shortcut).WorkingDirectory
}
$Shortcuts += New-Object PSObject -Property $Properties
} catch {
Write-Warning "Failed to get properties for shortcut $($Shortcut.FullName). Error: $_"
}
}
}
# Release the COM object
[Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null
# Export the array to a CSV file
if ($CSVHeader) {
$Shortcuts | Select-Object ShortcutName, ShortcutFull, ShortcutPath, Target, WorkingDirectory | Export-Csv $FilePath -NoTypeInformation -Force
} else {
$Shortcuts | Export-Csv $FilePath -NoTypeInformation -Force
}
# Return the number of shortcuts found
Write-Output "Found $($Shortcuts.Count) shortcuts"
}

Get-Shortcuts
