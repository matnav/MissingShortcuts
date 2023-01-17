param (
    [string[]]$ExcludePaths = @(),
    [string]$FilePath = "c:\matnav\shortcuts-uptodate.csv",
    [string]$SearchPath = "C:\ProgramData\Microsoft\Windows\Start Menu",
    [bool]$ShowOutput = $false
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
        if ($ExcludePaths.Length -eq 0) {
            # Create a custom object to store the shortcut properties
            $Properties = @{
                ShortcutName     = $Shortcut.Name
                ShortcutFull     = $Shortcut.FullName
                ShortcutPath     = $shortcut.DirectoryName
                Target           = $Shell.CreateShortcut($Shortcut).TargetPath
                Arguments        = $Shell.CreateShortcut($Shortcut).Arguments
                WorkingDirectory = $Shell.CreateShortcut($Shortcut).WorkingDirectory
            }
            $Shortcuts += New-Object PSObject -Property $Properties
            if($ShowOutput){
                Write-Output "ShortcutName: $($Properties.ShortcutName) ShortcutFull: $($Properties.ShortcutFull) ShortcutPath: $($Properties.ShortcutPath) Target: $($Properties.Target) Arguments: $($Properties.Arguments) WorkingDirectory: $($Properties.WorkingDirectory)"
            }
        }
        else {
            # Check if the shortcut path contains any of the excluded strings
            if ($ExcludePaths -notcontains ($Shortcut.DirectoryName | Select-String -SimpleMatch -Quiet)) {
                # Create a custom object to store the shortcut properties
                $Properties = @{
                    ShortcutName     = $Shortcut.Name
                    ShortcutFull     = $Shortcut.FullName
                    ShortcutPath     = $shortcut.DirectoryName
                    Target           = $Shell.CreateShortcut($Shortcut).TargetPath
                    Arguments        = $Shell.CreateShortcut($Shortcut).Arguments
                    WorkingDirectory = $Shell.CreateShortcut($Shortcut).WorkingDirectory
                }
                $Shortcuts += New-Object PSObject -Property $Properties
                if($ShowOutput){
                    Write-Output "ShortcutName: $($Properties.ShortcutName) ShortcutFull: $($Properties.ShortcutFull) ShortcutPath: $($Properties.ShortcutPath) Target: $($Properties.Target) Arguments: $($Properties.Arguments) WorkingDirectory: $($Properties.WorkingDirectory)"
                }
            }
        }
    }
    # Release the COM object
    [Runtime.InteropServices.Marshal]::ReleaseComObject($Shell) | Out-Null
        if (!(Test-Path -Path $FilePath)) {
        # Create the parent directory if it does not exist
        New-Item -ItemType Directory -Path (Split-Path $FilePath -Parent)  -ErrorAction SilentlyContinue| Out-Null
    }
    # Export the array to a CSV file
    $Shortcuts | Export-Csv -Path $FilePath -NoTypeInformation -Force
    # Return the number of shortcuts found
    Write-Output "Found $($Shortcuts.Count) shortcuts"
}

Get-Shortcuts
