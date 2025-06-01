# Ensure the script can run with elevated privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as an Administrator!"
    break
}

# Function to test internet connectivity
function Test-InternetConnection {
    try {
        $testConnection = Test-Connection -ComputerName www.google.com -Count 1 -ErrorAction Stop
        return $true
    }
    catch {
        Write-Warning "Internet connection is required but not available. Please check your connection."
        return $false
    }
}

# Function to install CodingFonts
function Install-CodingFonts {
    param (
        [string]$FontName = "CodingFonts",
        [string]$FontDisplayName = "Coding Fonts"
    )

    try {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        $fontFamilies = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
        if ($fontFamilies -notcontains "${FontDisplayName}") {
            $fontZipUrl = "https://github.com/o9-9/files/raw/main/${FontName}.zip"
            $zipFilePath = "$env:TEMP\${FontName}.zip"
            $extractPath = "$env:TEMP\${FontName}"

            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFileAsync((New-Object System.Uri($fontZipUrl)), $zipFilePath)

            while ($webClient.IsBusy) {
                Start-Sleep -Seconds 2
            }

            Expand-Archive -Path $zipFilePath -DestinationPath $extractPath -Force
            $destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
            Get-ChildItem -Path $extractPath -Recurse -Filter "*.ttf" | ForEach-Object {
                If (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {
                    $destination.CopyHere($_.FullName, 0x10)
                }
            }

            Remove-Item -Path $extractPath -Recurse -Force
            Remove-Item -Path $zipFilePath -Force
        } else {
            Write-Host "Font ${FontDisplayName} already installed"
        }
    }
    catch {
        Write-Error "Failed to download or install ${FontDisplayName} font. Error: $_"
    }
}

# Check for internet connectivity before proceeding
if (-not (Test-InternetConnection)) {
    break
}

# Font Install
Install-CodingFonts -FontName "CodingFonts" -FontDisplayName "Coding Fonts"

# Installs the latest VS Code
winget install --id Microsoft.VisualStudioCode --scope machine --accept-package-agreements --accept-source-agreements

Write-Host -ForegroundColor Green "✔ VS Code Installed."

# Define GitHub repository and VS Code paths
$repoUrl = "https://raw.githubusercontent.com/o9-9/vscode-setup/main"
$vsCodeUserPath = "$env:APPDATA\Code\User"

# Ensure the VS Code settings directory exists
if (!(Test-Path $vsCodeUserPath)) {
    New-Item -ItemType Directory -Path $vsCodeUserPath -Force
   Write-Host -ForegroundColor Green "✔ Created VS Code settings directory."
}

# Download and copy settings.json
Invoke-WebRequest -Uri "$repoUrl/settings.json" -OutFile "$vsCodeUserPath\settings.json"
Write-Host -ForegroundColor Green "✔ Copied settings.json to VS Code."

# Download and copy keybindings.json
Invoke-WebRequest -Uri "$repoUrl/keybindings.json" -OutFile "$vsCodeUserPath\keybindings.json"
Write-Host -ForegroundColor Green "✔ Copied keybindings.json to VS Code."

# Refresh environment variables
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
Write-Host -ForegroundColor Green "✔ Environment variables refreshed."

# Download and install extensions
$extensionsJson = "$env:TEMP\extensions.json"
Invoke-WebRequest -Uri "$repoUrl/extensions.json" -OutFile $extensionsJson
$extensions = (Get-Content $extensionsJson | ConvertFrom-Json).extensions
$extensions | ForEach-Object {
    code --install-extension $_
    Write-Host -ForegroundColor Cyan "✔ Installed $_"
}
Remove-Item $extensionsJson

Write-Host -ForegroundColor Green "✔ VS Code setup complete."

# create reg file
$MultilineComment = @"
Windows Registry Editor Version 5.00

; Open files with VS Code
[HKEY_CLASSES_ROOT\*\shell\Open with VS Code]
@="Edit with VS Code"
"Icon"="C:\\Program Files\\Microsoft VS Code\\Code.exe,0"

[HKEY_CLASSES_ROOT\*\shell\Open with VS Code\command]
@="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\" \"%1\""

; Open folder as VS Code Project (right-click ON a folder)
[HKEY_CLASSES_ROOT\Directory\shell\vscode]
@="Open Folder as VS Code Project"
"Icon"="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\",0"

[HKEY_CLASSES_ROOT\Directory\shell\vscode\command]
@="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\" \"%1\""

; Open folder as VS Code Project (right-click INSIDE a folder)
[HKEY_CLASSES_ROOT\Directory\Background\shell\vscode]
@="Open Folder as VS Code Project"
"Icon"="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\",0"

[HKEY_CLASSES_ROOT\Directory\Background\shell\vscode\command]
@="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\" \"%V\""
"@
Set-Content -Path "$env:TEMP\VS Code Context Menu.reg" -Value $MultilineComment -Force
# edit reg file
$path = "$env:TEMP\VS Code Context Menu.reg"
(Get-Content $path) -replace "\?","$" | Out-File $path
# disable optimize drives
schtasks /Change /DISABLE /TN "\Microsoft\Windows\Defrag\ScheduledDefrag" | Out-Null
# import reg file
Regedit.exe /S "$env:TEMP\VS Code Context Menu.reg"

Write-Host -ForegroundColor Green "✔ Add VS Code Context Menu complete."
