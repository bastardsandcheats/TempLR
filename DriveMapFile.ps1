# Define the URL of the file to download
$url = "https://raw.githubusercontent.com/bastardsandcheats/TempLR/main/MapLabreportsAirdrie.bat"

# Define the destination path (Public Desktop)
$publicDesktop = [System.Environment]::GetFolderPath('CommonDesktopDirectory')
$destination = Join-Path -Path $publicDesktop -ChildPath "MapLabreportsAirdrie.bat"

# Use Invoke-WebRequest to download the file
Invoke-WebRequest -Uri $url -OutFile $destination

Write-Host "File downloaded to $destination"
