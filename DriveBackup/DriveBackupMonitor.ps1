Start-Transcript -Path "C:\Temp\DriveBackup\VerboseBackup.log"
if (Get-Module -ListAvailable -Name 'HashCopy') {
    Import-Module -Name 'HashCopy'
} else {
    return $false
}
$Name = $env:COMPUTERNAME + " Backup"

$path = "G:\My Drive"

$GoogleWebAppURL = 'https://script.google.com/macros/s/ScriptAddress'  

#Checking For Presence of Google Drive File Stream
if(test-path -path $env:USERPROFILE\AppData\Local\Google\DriveFS){
    #Test if there is an AlphaNumeric Folder that signifies the user is signed into Google Drive
    if(Get-ChildItem -Path $env:USERPROFILE\AppData\Local\Google\DriveFS -Directory | Where-Object -Property Name -Match '\d'){
        $DriveFS = Get-Process GoogleDriveFS -ErrorAction SilentlyContinue
        if ($DriveFS){
            $Column5Name = 'Drive_Status'
            $Column5Value = "Drive is Running, Backing Up"
            if (!$Column5Value) {$Column5Value = "N/A"}
            if(!(Test-Path -Path $path)){
                $Column9Name = 'Error_Output'
                $Column9Value = "Drive not mapped to G Drive"
                if (!$Column9Value) {$Column9Value = "N/A"}
            }
            else{
                Copy-FileHash -Path $env:USERPROFILE\Desktop -Destination "G:\My Drive\$($Name)\Desktop" -Recurse -PassThru -Verbose -ErrorAction Continue -ErrorVariable +err
                Copy-FileHash -Path $env:USERPROFILE\Documents -Destination "G:\My Drive\$($Name)\Documents" -Recurse -PassThru -Verbose -ErrorAction Continue -ErrorVariable +err
                Copy-FileHash -Path $env:USERPROFILE\Pictures -Destination "G:\My Drive\$($Name)\Pictures" -Recurse -PassThru -Verbose -ErrorAction Continue -ErrorVariable +err

                $Column9Name = 'Error_Output'
                $Column9Value = ($err | Out-String).trim()
                if (!$Column9Value) {$Column9Value = "N/A"}

                Stop-Process -Name "GoogleDriveFS" -Force
                start-sleep -Seconds 1
                start-process -filepath "C:\Program Files\Google\Drive File Stream\launch.bat" -WindowStyle hidden
            }
            Stop-Transcript  
            Start-Sleep -Seconds 5  
            $Column10Name = 'Success_Output'
            $Column10Value = Get-Content -Path "C:\Temp\DriveBackup\VerboseBackup.log" -Raw            
            if (!$Column10Value) {$Column10Value = "N/A"}
        }
        else {
            $Column5Name = 'Drive_Status'
            $Column5Value = "Drive is not Running - Starting & Backing Up"
            if (!$Column5Value) {$Column5Value = "N/A"}
            Start-Process -Filepath "C:\Program Files\Google\Drive File Stream\launch.bat" -WindowStyle hidden
            Start-Sleep -Seconds 15
            if(!(Test-Path -Path $path)){
                $Column9Name = 'Error_Output'
                $Column9Value = "Drive not mapped to G Drive"
                if (!$Column9Value) {$Column9Value = "N/A"}
            }
            else{
                Copy-FileHash -Path $env:USERPROFILE\Desktop -Destination "G:\My Drive\$($Name)\Desktop" -Recurse -PassThru -Verbose -ErrorAction Continue -ErrorVariable +err
                Copy-FileHash -Path $env:USERPROFILE\Documents -Destination "G:\My Drive\$($Name)\Documents" -Recurse -PassThru -Verbose -ErrorAction Continue -ErrorVariable +err
                Copy-FileHash -Path $env:USERPROFILE\Pictures -Destination "G:\My Drive\$($Name)\Pictures" -Recurse -PassThru -Verbose -ErrorAction Continue -ErrorVariable +err
                
                $Column9Name = 'Error_Output'
                $Column9Value = ($err | Out-String).trim()
                if (!$Column9Value) {$Column9Value = "N/A"}
                Stop-Process -Name "GoogleDriveFS" -Force
                start-sleep -Seconds 1
                start-process -filepath "C:\Program Files\Google\Drive File Stream\launch.bat" -WindowStyle hidden
            }
            Stop-Transcript  
            Start-Sleep -Seconds 5  
            $Column10Name = 'Success_Output'
            $Column10Value = Get-Content -Path "C:\Temp\DriveBackup\VerboseBackup.log" -Raw            
            if (!$Column10Value) {$Column10Value = "N/A"}
        }
    }
    else {
        $Column5Name = 'Drive_Status'
        $Column5Value = "Not Signed Into Drive"
        if (!$Column5Value) {$Column5Value = "N/A"}

        $Column9Name = 'Error_Output'
        $Column9Value = "N/A"
        if (!$Column9Value) {$Column9Value = "N/A"}

        $Column10Name = 'Success_Output'
        $Column10Value = "N/A"
        if (!$Column10Value) {$Column10Value = "N/A"}
        Stop-Transcript
    }
}
else {
    $Column5Name = 'Drive_Status'
    $Column5Value = "There was an issue with the backup"
    if (!$Column5Value) {$Column5Value = "N/A"}
    
    $Column9Name = 'Error_Output'
    $Column9Value = "N/A"
    if (!$Column9Value) {$Column9Value = "N/A"}

    $Column10Name = 'Success_Output'
    $Column10Value = "N/A"
    if (!$Column10Value) {$Column10Value = "N/A"}
    Stop-Transcript
}

$Column2Name = 'Computer'
$Column2Value = ($env:COMPUTERNAME | Out-String).trim()
if (!$Column2Value) {$Column2Value = "N/A"}

$Column3Name = 'Model'
$Column3Value = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Model | Out-String).trim()
if (!$Column3Value) {$Column3Value = "N/A"}

$Column4Name = 'User'
$Column4Value = ((Get-ItemProperty Registry::\HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\CloudDomainJoin\JoinInfo\*).UserEmail | Out-String).trim()
if (!$Column4Value) {$Column4Value = "N/A"}

$DesktopFolders = (Get-ChildItem -Directory $env:USERPROFILE\Desktop -Recurse -Exclude *.lnk | Measure-Object).Count
$DesktopFiles = (Get-ChildItem -File $env:USERPROFILE\Desktop -Recurse -Exclude *.lnk | Measure-Object).Count

$Column6Name = 'Desktop_Items'
$Column6Value = "$($DesktopFolders) Folders and $($DesktopFiles) Files"
if (!$Column6Value) {$Column6Value = "N/A"}    

$DocumentFolders = (Get-ChildItem -Directory $env:USERPROFILE\Documents -Recurse -Exclude *.lnk | Measure-Object).Count
$DocumentFiles = (Get-ChildItem -File $env:USERPROFILE\Documents -Recurse -Exclude *.lnk | Measure-Object).Count

$Column7Name = 'Document_Items'
$Column7Value = "$($DocumentFolders) Folders and $($DocumentFiles) Files"
if (!$Column7Value) {$Column7Value = "N/A"}

$PicturesFolders = (Get-ChildItem -Directory $env:USERPROFILE\Pictures -Recurse -Exclude *.lnk | Measure-Object).Count
$PicturesFiles = (Get-ChildItem -File $env:USERPROFILE\Pictures -Recurse -Exclude *.lnk | Measure-Object).Count

$Column8Name = 'Picture_Items'
$Column8Value = "$($PicturesFolders) Folders and $($PicturesFiles) Files"
if (!$Column8Value) {$Column8Value = "N/A"} 

$postParams2 = @{$Column2Name=$Column2Value;$Column3Name=$Column3Value;$Column4Name=$Column4Value;$Column5Name=$Column5Value;$Column6Name=$Column6Value;$Column7Name=$Column7Value;$Column8Name=$Column8Value;$Column9Name=$Column9Value;$Column10Name=$Column10Value}

# Find keys of `$null` values
$nullKeys = $postParams2.Keys | Where-Object { $postParams2[$_] -eq $null }

# Populate appropriate indices with 0
$nullKeys | ForEach-Object { $postParams2[$_] = "N/A" }

#Post to the KPI-Test Sheet

Invoke-WebRequest -UseBasicParsing -Uri $GoogleWebAppURL -Method POST -Body $postParams2
Start-Sleep -Seconds 15
Remove-Item -Path "C:\Temp\DriveBackup\VerboseBackup.log" -Force            