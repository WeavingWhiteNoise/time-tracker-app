# Create a Desktop Shortcut for Time Tracker
$DesktopPath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("Desktop"), "Time Tracker.lnk")
$ExePath = [System.IO.Path]::Combine($PSScriptRoot, "dist", "TimeTracker.exe")

$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($DesktopPath)
$Shortcut.TargetPath = $ExePath
$Shortcut.WorkingDirectory = [System.IO.Path]::Combine($PSScriptRoot, "dist")
$Shortcut.Description = "Time Tracker Application"
$Shortcut.Save()

Write-Host "✓ Desktop shortcut created successfully!" -ForegroundColor Green
Write-Host "You can now double-click 'Time Tracker' on your desktop to launch the app."
