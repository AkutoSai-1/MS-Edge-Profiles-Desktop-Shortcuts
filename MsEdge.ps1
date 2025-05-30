$EdgeProfiles = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Edge\User Data" | Where-Object { $_.Name -like "Profile *"}

ForEach($EdgeProfile in $EdgeProfiles.Name) {
    # Get Profile Name (Refresh Data Before Using)
    $Preferences = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\$EdgeProfile\Preferences"
    if (Test-Path $Preferences) {
        Start-Sleep -Seconds 1  # Ensure Preferences file isn't locked
        $Data = (ConvertFrom-Json (Get-Content $Preferences -Raw))
        $ProfileName = $Data.Profile.Name

        # Define Shortcut File
        $ShortcutFile = "$env:USERPROFILE\Desktop\$ProfileName - Edge.lnk"

        # Force Delete Shortcut Using CMD
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c del `"$ShortcutFile`"" -NoNewWindow -Wait

        Start-Sleep -Seconds 2  # Small delay after deletion

        # Create New Shortcut
        $TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
        $WScriptShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
        $Shortcut.IconLocation = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\$EdgeProfile\Edge Profile.ico, 0"
        $Shortcut.Arguments = "--profile-directory=`"$EdgeProfile`""
        $Shortcut.TargetPath = $TargetPath
        $Shortcut.Save()
    }
}

# Handle Default Profile (For Person 1 Fix)
$Preferences = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Preferences"
if (Test-Path $Preferences) {
    Start-Sleep -Seconds 1  # Ensure file refresh
    $Data = (ConvertFrom-Json (Get-Content $Preferences -Raw))
    $ProfileName = $Data.Profile.Name
    if ($ProfileName -eq "Person 1") { $ProfileName = "Default" }

    # Define Shortcut File
    $ShortcutFile = "$env:USERPROFILE\Desktop\$ProfileName - Edge.lnk"

    # Force Delete Shortcut Using CMD
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c del `"$ShortcutFile`"" -NoNewWindow -Wait

    Start-Sleep -Seconds 2  # Small delay after deletion

    # Create New Shortcut
    $TargetPath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.IconLocation = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Edge Profile.ico, 0"
    $Shortcut.Arguments = "--profile-directory=`"Default`""
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.Save()
}
