# PowerShell script to create labinfo.ini, create a shortcut, and delete specific files

# Step 1: Delete labInfo.ini from each user's OneDrive directory and from C:\
$RelativeFilePath = "OneDrive - IGNE\Documents\LabReports\labInfo.ini"

# Get all user directories in C:\Users
$UserDirectories = Get-ChildItem -Path "C:\Users" -Directory

foreach ($UserDir in $UserDirectories) {
    $FullFilePath = Join-Path -Path $UserDir.FullName -ChildPath $RelativeFilePath
    if (Test-Path $FullFilePath) {
        Write-Host "Deleting: $FullFilePath"
        Remove-Item $FullFilePath -Force
    } else {
        Write-Host "File not found: $FullFilePath"
    }
}

# Also delete labinfo.ini from C:\
$GlobalLabInfoPath = "C:\labinfo.ini"
if (Test-Path $GlobalLabInfoPath) {
    Remove-Item $GlobalLabInfoPath -Force
}

# Step 2: Create labinfo.ini with specified content
# Define the URL of the LabinfoAirdrie.ini file
$url = "https://raw.githubusercontent.com/bastardsandcheats/TempLR/main/LabinfoAston.ini"
# Define the destination path where the file should be saved
$destinationPath = "C:\labinfo.ini"
# Use Invoke-WebRequest to download the file
Invoke-WebRequest -Uri $url -OutFile $destinationPath
Write-Host "The file has been downloaded and saved to $destinationPath"

# Step 3: Create a shortcut to Excel with arguments
$shortcutPath = [System.Environment]::GetFolderPath('CommonDesktopDirectory') + "\Lab Reports on prem.lnk"
$excelPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
$excelArguments = '/e "Z:\Lab\Lab Report\Installation\UserFunctions.xla"'

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $excelPath
$Shortcut.Arguments = $excelArguments
$Shortcut.IconLocation = "C:\Windows\System32\imageres.dll,99"
$Shortcut.Save()

Write-Host "Operation completed successfully."
