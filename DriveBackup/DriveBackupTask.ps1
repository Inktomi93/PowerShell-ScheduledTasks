$path = "C:\Temp\DriveBackup"
If(test-path $path)
{
  Remove-Item $path -Force  -Recurse -ErrorAction SilentlyContinue
}

try{
  if (Get-ScheduledTask -TaskName "DriveBackup" -ErrorAction Stop)
  {
    Write-Output "Found Task"
    Unregister-ScheduledTask -TaskName "DriveBackup" -Confirm:$False
  }
}
Catch {
  Write-Output "No Previous Scheduled Task Found"
}

New-Item -Path "C:\Temp" -Name "DriveBackup" -ItemType "directory" -Force

Copy-Item ".\DriveBackupMonitor.ps1" -Destination "C:\Temp\DriveBackup"
Copy-Item ".\DriveBackupMonitor.xml" -Destination "C:\Temp\DriveBackup"

REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\FileSystem" /v LongPathsEnabled /t REG_DWORD /d 1 /F

# Trust PSGallery
Set-PSRepository -Name 'PSGallery' -installationPolicy Trusted

# NuGet
Install-PackageProvider -Name NuGet -Force | Out-Null

# Install HashCopy
if (Get-InstalledModule -Name "HashCopy" -ErrorAction Ignore){
Write-Output "Module Already Loaded"
}
else{
Install-Module -Name HashCopy -Scope AllUsers -Force -ErrorAction Ignore
}

schtasks /Create /XML "C:\Temp\DriveBackup\DriveBackupMonitor.xml" /tn DriveBackup

Write-output "Succesfully Installed"

Start-ScheduledTask -TaskName "DriveBackup"

Exit 0