# PowerShell script to create labinfo.ini, create a shortcut, and delete specific files

# Step 1: Delete the LabReports folder from each user's OneDrive directory
$RelativeFolderPath = "OneDrive - IGNE\Documents\LabReports"

# Get all user directories in C:\Users
$UserDirectories = Get-ChildItem -Path "C:\Users" -Directory

foreach ($UserDir in $UserDirectories) {
    $FullFolderPath = Join-Path -Path $UserDir.FullName -ChildPath $RelativeFolderPath
    if (Test-Path $FullFolderPath) {
        Write-Host "Deleting: $FullFolderPath"
        Remove-Item $FullFolderPath -Recurse -Force
    } else {
        Write-Host "Folder not found: $FullFolderPath"
    }
}

# Also delete labinfo.ini from C:\
$GlobalLabInfoPath = "C:\labinfo.ini"
if (Test-Path $GlobalLabInfoPath) {
    Remove-Item $GlobalLabInfoPath -Force
}

# Step 2: Create labinfo.ini with specified content
$labInfoPath = "C:\labinfo.ini"
$iniContent = @"
[PATH]
ROOTPATH=z:\Lab\Lab Report\Installation\
"@

# Ensure UTF-8 encoding without BOM
[System.IO.File]::WriteAllText($labInfoPath, $iniContent, [System.Text.Encoding]::UTF8)


# Step 3: Create a shortcut to Excel with arguments
$shortcutPath = [System.Environment]::GetFolderPath('CommonDesktopDirectory') + "\Lab Reports on prem.lnk"
$excelPath = "C:\Program Files (x86)\Microsoft Office\OFFICE11\EXCEL.EXE"
$excelArguments = '/e "z:\Lab Reports\System\Installation\UserFunctions.xla"'

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $excelPath
$Shortcut.Arguments = $excelArguments
$Shortcut.IconLocation = "C:\Windows\System32\imageres.dll,99"
$Shortcut.Save()

Write-Host "Operation completed successfully."
