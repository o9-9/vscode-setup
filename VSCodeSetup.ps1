# Ensure Script Runs With Administrator Privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please Run Setup as an Administrator"
    exit
}

# Global Variables
$repoUrl = "https://raw.githubusercontent.com/o9-9/vscode-setup/main"
$vsCodeUserPath = "$env:APPDATA\Code\User"
$vsCodeExtensionsPath = "$env:USERPROFILE\.vscode\extensions"

function Install-Config {
    <#
    .SYNOPSIS
    Installs VS Code and Applies User Settings, Keybindings, Extensions, and Context Menu Entries.
    .DESCRIPTION
    Installs the Latest Visual Studio Code, Updates the User Settings and Keybindings from a GitHub Repo,
    installs the Listed Extensions, Refreshes Environment Variables, and Sets Up the VS Code context Menu Via Registry.
    #>
    Write-Host -ForegroundColor White "Installing Visual Studio Code..."
    winget install --id Microsoft.VisualStudioCode --scope machine --accept-package-agreements --accept-source-agreements | Out-Null
    Write-Host -ForegroundColor Green "✔ VS Code Installed."

    Write-Host -ForegroundColor White "Configuring VS Code User Settings..."
    if (!(Test-Path $vsCodeUserPath)) {
        New-Item -ItemType Directory -Path $vsCodeUserPath -Force | Out-Null
        Write-Host -ForegroundColor Cyan "✔ Created VS Code User Settings Directory."
    }

    # Download settings.json and keybindings.json from the Repository
    Invoke-WebRequest -Uri "$repoUrl/settings.json" -OutFile "$vsCodeUserPath\settings.json" -UseBasicParsing
    Write-Host -ForegroundColor Yellow "✔ Copied Settings.json."
    Invoke-WebRequest -Uri "$repoUrl/keybindings.json" -OutFile "$vsCodeUserPath\keybindings.json" -UseBasicParsing
    Write-Host -ForegroundColor Yellow "✔ Copied Keybindings.json."

    # Refresh Environment Variables to Ensure New Tools are on the PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path", "User")
    Write-Host -ForegroundColor Magenta "✔ Environment Variables Refreshed."

    # Install Extensions listed in the Extensions.json from the Repo
    $extensionsJson = "$env:TEMP\extensions.json"
    Invoke-WebRequest -Uri "$repoUrl/extensions.json" -OutFile $extensionsJson -UseBasicParsing
    $extensions = (Get-Content $extensionsJson -Raw | ConvertFrom-Json).extensions
    foreach ($ext in $extensions) {
        code --install-extension $ext | Out-Null
        Write-Host -ForegroundColor Cyan "✔ Installed Extension: $ext"
    }
    Remove-Item -Path $extensionsJson -Force

    # Configure VS Code Context Menu Via Registry Entries
    Write-Host -ForegroundColor White "Configuring VS Code Context Menu..."
    $regFilePath = "$env:TEMP\VSCodeContextMenu.reg"
    $regContent = @"
Windows Registry Editor Version 5.00

; Open files with VS Code
[HKEY_CLASSES_ROOT\*\shell\Open with VS Code]
@="\"Edit with VS Code\""
"Icon"="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\",0"

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
    Set-Content -Path $regFilePath -Value $regContent -Force

    # Import Registry Entries Silently
    Regedit.exe /S $regFilePath
    Remove-Item -Path $regFilePath -Force
    Write-Host -ForegroundColor Green "✔ VS Code Context Menu Configured."

    Write-Host -ForegroundColor Green "✔ VS Code Setup Complete."
}

function Install-CodingFonts {
    <#
    .SYNOPSIS
    Downloads and Installs Fonts from a Zip Archive on GitHub.
    .PARAMETER FontName
    The Base Name of the Fonts Zip File (Without Extension) on the Remote Repo.
    .PARAMETER FontDisplayName
    The Display Name of the Fonts to Verify Installation.
    #>
    param (
        [string]$FontName = "CodingFonts",
        [string]$FontDisplayName = "Coding Fonts"
    )
    try {
        Add-Type -AssemblyName System.Drawing
        $installedFonts = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
        if ($installedFonts -notcontains $FontDisplayName) {
            $fontZipUrl   = "https://github.com/o9-9/vscode-setup/raw/main/$($FontName).zip"
            $zipFilePath  = "$env:TEMP\$($FontName).zip"
            $extractPath  = "$env:TEMP\$($FontName)"

            # Download the font Zip File
            Invoke-WebRequest -Uri $fontZipUrl -OutFile $zipFilePath -UseBasicParsing

            # Extract and Install Fonts Files
            Expand-Archive -LiteralPath $zipFilePath -DestinationPath $extractPath -Force
            $shellApp   = New-Object -ComObject Shell.Application
            $fontFolder = $shellApp.Namespace(0x14)
            Get-ChildItem -Path $extractPath -Filter "*.ttf" -Recurse | ForEach-Object {
                if (!(Test-Path "C:\Windows\Fonts\$($_.Name)")) {
                    $fontFolder.CopyHere($_.FullName, 0x10)
                }
            }

            # Clean Up
            Remove-Item -Path $extractPath -Recurse -Force
            Remove-Item -Path $zipFilePath -Force
            Write-Host -ForegroundColor Green "✔ Installed Font"
        }
        else {
            Write-Host -ForegroundColor Cyan "Font Already Installed"
        }
    } catch {
        Write-Error "Failed to Install Fonts"
    }
}

function Install-Fonts {
    Write-Host -ForegroundColor White "Installing coding fonts..."
    Install-CodingFonts -FontName "CodingFonts" -FontDisplayName "Coding Fonts"
}

function Install-Theme {
    <#
    .SYNOPSIS
    Downloads the o9 Theme from GitHub and Installs it Into the .VSCode Extensions Folder.
    #>
    Write-Host -ForegroundColor White "Installing o9 Theme..."
    $zipUrl      = "https://github.com/o9-9/vscode-setup/archive/refs/heads/main.zip"
    $zipPath     = "$env:TEMP\vscode-setup-main.zip"
    $extractDir  = "$env:TEMP\vscode-setup"
    
    # Download and Extract the Repository Zip
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractDir -Force

    # The Theme Folder Inside the Extracted Content
    $themeSource = Join-Path $extractDir "o9-Theme"
    $themeDest   = Join-Path $vsCodeExtensionsPath "o9-Theme"

    if (Test-Path $themeDest) {
        Remove-Item -Path $themeDest -Recurse -Force
    }
    Move-Item -Path $themeSource -Destination $themeDest

    # Clean Up
    Remove-Item -Path $extractDir -Recurse -Force
    Remove-Item -Path $zipPath -Force

    Write-Host -ForegroundColor Green "✔ o9 Theme Installed."
}

# Main Interactive Menu Loop
while ($true) {
    Write-Host ""
    Write-Host "Select Option:"
    Write-Host "1. Install Config"
    Write-Host "2. Install Fonts"
    Write-Host "3. Install Theme"
    Write-Host "4. Back"

    $choice = Read-Host "Enter choice (1-4)"
    switch ($choice) {
        "1" { Install-Config }
        "2" { Install-Fonts }
        "3" { Install-Theme }
        "4" {
            Write-Host -ForegroundColor Green "Goodbye"
            break
        }
        default { Write-Warning "Please Enter 1, 2, 3, or 4." }
    }
}
