# Ensure the script can run with elevated privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as an Administrator!"
    break
}

function Install-o9Theme {
    param (
        [string]$ThemeUrl = "https://github.com/o9-9/vscode-setup/releases/download/9/o9-Theme.zip",
        [string]$ZipPath = "$env:TEMP\o9-Theme.zip",
        [string]$ExtractPath = "$env:TEMP\o9-Theme",
        [string]$DestinationPath = "C:\Program Files\Microsoft VS Code\resources\app\extensions",
        [string]$SevenZipPath = "C:\Program Files\7-Zip\7z.exe"
    )

    Write-Host "`n[1/4] Downloading o9 Theme..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $ThemeUrl -OutFile $ZipPath -UseBasicParsing
        Write-Host "[OK] Downloaded: $ZipPath" -ForegroundColor Green
    } catch {
        Write-Error "Failed to download o9 Theme: $_"
        return
    }

    Write-Host "`n[2/4] Extracting zip with 7-Zip..." -ForegroundColor Cyan
    if (!(Test-Path $SevenZipPath)) {
        Write-Error "7-Zip not found at: $SevenZipPath"
        return
    }

    if (Test-Path $ExtractPath) {
        Remove-Item -Path $ExtractPath -Recurse -Force
    }

    try {
        & "$SevenZipPath" x $ZipPath -o"$ExtractPath" -y | Out-Null
        Write-Host "[OK] Extracted to: $ExtractPath" -ForegroundColor Green
    } catch {
        Write-Error "Extraction failed: $_"
        return
    }

    Write-Host "`n[3/4] Moving theme to VS Code extensions directory..." -ForegroundColor Cyan
    try {
        $folderName = Get-ChildItem -Path $ExtractPath | Where-Object { $_.PSIsContainer } | Select-Object -First 1
        $targetFolder = Join-Path -Path $DestinationPath -ChildPath $folderName.Name

        if (Test-Path $targetFolder) {
            Remove-Item -Path $targetFolder -Recurse -Force
        }

        Move-Item -Path $folderName.FullName -Destination $DestinationPath
        Write-Host "[OK] o9 Theme installed at: $targetFolder" -ForegroundColor Green
    } catch {
        Write-Error "Failed to move theme folder: $_"
        return
    }

    Write-Host "`n[4/4] Installation complete." -ForegroundColor Green
}

function Install-JetBrainsMono {
    param (
        [string]$FontName = "JetBrainsMono",
        [string]$FontDisplayName = "JetBrainsMono NF",
        [string]$Version = "3.2.1"
    )

    try {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        $fontFamilies = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
        if ($fontFamilies -notcontains "${FontDisplayName}") {
            $fontZipUrl = "https://github.com/ryanoasis/nerd-fonts/releases/download/v${Version}/${FontName}.zip"
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

function Install-CodeFonts {
    param (
        [Parameter(Mandatory=$false)]
        [string]$ZipUrl = "https://github.com/o9-9/vscode-setup/raw/main/CodeFonts.zip"
    )

    # Ensure running as administrator
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "You must run this script as an administrator."
        return
    }

    # Create temp directory
    $TempPath = [System.IO.Path]::Combine($env:TEMP, "CodeFonts_" + [System.Guid]::NewGuid())
    New-Item -Path $TempPath -ItemType Directory -Force | Out-Null

    $ZipPath = [System.IO.Path]::Combine($TempPath, "CodeFonts.zip")

    try {
        Write-Host "`n[1/4] Downloading fonts zip from GitHub..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $ZipUrl -OutFile $ZipPath

        Write-Host "`n[2/4] Extracting fonts..." -ForegroundColor Cyan
        Expand-Archive -Path $ZipPath -DestinationPath $TempPath -Force

        # Find all font files (ttf/otf/woff/woff2)
        $FontFiles = Get-ChildItem -Path $TempPath -Recurse -Include *.ttf, *.otf, *.woff, *.woff2

        if ($FontFiles.Count -eq 0) {
            throw "No font files found in the archive."
        }

        Write-Host "`n[3/4] Installing fonts..." -ForegroundColor Cyan

        foreach ($FontFile in $FontFiles) {
            $FontDest = Join-Path -Path "$env:SystemRoot\Fonts" -ChildPath $FontFile.Name

            # Copy font to system fonts folder
            if (-not (Test-Path $FontDest)) {
                Copy-Item -Path $FontFile.FullName -Destination $FontDest -Force
            }

            # Register font in the Registry
            $FontName = [System.IO.Path]::GetFileName($FontFile.Name)
            $FontExt = $FontFile.Extension.ToLower()

            switch ($FontExt) {
                ".ttf" { $RegVal = $FontName }
                ".otf" { $RegVal = $FontName }
                ".woff" { $RegVal = $FontName }
                ".woff2" { $RegVal = $FontName }
                default { $RegVal = $FontName }
            }

            # Try to extract display name from the font file (optional, fallback to file name)
            try {
                Add-Type -AssemblyName System.Drawing
                $fontCol = New-Object System.Drawing.Text.PrivateFontCollection
                $fontCol.AddFontFile($FontFile.FullName)
                if ($fontCol.Families.Count -gt 0) {
                    $DisplayName = $fontCol.Families[0].Name
                } else {
                    $DisplayName = $FontName
                }
            } catch {
                $DisplayName = $FontName
            }

            $RegPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
            $RegValueName = "$DisplayName ($FontExt)"
            Set-ItemProperty -Path $RegPath -Name $RegValueName -Value $FontName -Force
        }

        Write-Host "`n[4/4] Fonts installed successfully." -ForegroundColor Cyan

        # Refresh font cache (no perfect way, but send WM_FONTCHANGE to all windows)
        Add-Type @"
using System;
using System.Runtime.InteropServices;
public class NativeMethods {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern int SendMessageTimeout(
        IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);
}
"@

        $HWND_BROADCAST = [intptr]0xffff
        $WM_FONTCHANGE = 0x001D
        $SMTO_ABORTIFHUNG = 0x0002
        $result = [intptr]::Zero

        [void][NativeMethods]::SendMessageTimeout($HWND_BROADCAST, $WM_FONTCHANGE, [intptr]0, [intptr]0, $SMTO_ABORTIFHUNG, 1000, [ref]$result)

        Write-Host "Font cache refreshed. Fonts should now be available in all applications." -ForegroundColor Cyan

    } catch {
        Write-Error "An error occurred: $_"
    } finally {
        Write-Host "Cleaning up temporary files..."
        Remove-Item -Path $TempPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Install-Config {
    # Installs the latest VS Code
    Write-Host "`n[1/5] Installing VS Code..." -ForegroundColor Cyan
    winget install --id Microsoft.VisualStudioCode --scope machine --accept-package-agreements --accept-source-agreements
    Write-Host "✔ VS Code Installed." -ForegroundColor Green

    # Define GitHub Repository and VS Code Paths
    Write-Host "`n[2/5] Installing GitHub Repository..." -ForegroundColor Cyan
    $repoUrl = "https://raw.githubusercontent.com/o9-9/vscode-setup/main"
    $vsCodeUserPath = "$env:APPDATA\Code\User"
    Write-Host "✔ GitHub Repository Installed." -ForegroundColor Green

    # Ensure the VS Code Settings Directory Exists
    if (!(Test-Path $vsCodeUserPath)) {
        New-Item -ItemType Directory -Path $vsCodeUserPath -Force
       Write-Host "✔ Created VS Code Settings Directory." -ForegroundColor Green
    }

    # Download and copy settings.json
    Invoke-WebRequest -Uri "$repoUrl/settings.json" -OutFile "$vsCodeUserPath\settings.json"
    Write-Host "✔ Copied Settings.json to VS Code." -ForegroundColor Green

    # Download and copy keybindings.json
    Invoke-WebRequest -Uri "$repoUrl/keybindings.json" -OutFile "$vsCodeUserPath\keybindings.json"
    Write-Host "✔ Copied Keybindings.json to VS Code." -ForegroundColor Green

    # Refresh environment variables
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    Write-Host "✔ Environment Variables Refreshed." -ForegroundColor Green

    # Download and install extensions
    Write-Host "`n[3/5] Installing Extensions..." -ForegroundColor Cyan
    $extensionsJson = "$env:TEMP\extensions.json"
    Invoke-WebRequest -Uri "$repoUrl/extensions.json" -OutFile $extensionsJson
    $extensions = (Get-Content $extensionsJson | ConvertFrom-Json).extensions
    $extensions | ForEach-Object {
        code --install-extension $_
        Write-Host "✔ Installed $_" -ForegroundColor Cyan
    }
    Remove-Item $extensionsJson

    # Adding VS Code to Right-Click Context Menu
    Write-Host "`n[4/5] Adding VS Code to Context Menu..." -ForegroundColor Cyan
    $MultilineComment = @"
Windows Registry Editor Version 5.00

[HKEY_CLASSES_ROOT\*\shell\Open with VS Code]
@="Edit with VS Code"
"Icon"="C:\\Program Files\\Microsoft VS Code\\Code.exe,0"
[HKEY_CLASSES_ROOT\*\shell\Open with VS Code\command]
@="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\" \"%1\""

[HKEY_CLASSES_ROOT\Directory\shell\vscode]
@="Open Folder as VS Code Project"
"Icon"="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\",0"
[HKEY_CLASSES_ROOT\Directory\shell\vscode\command]
@="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\" \"%1\""

[HKEY_CLASSES_ROOT\Directory\Background\shell\vscode]
@="Open Folder as VS Code Project"
"Icon"="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\",0"
[HKEY_CLASSES_ROOT\Directory\Background\shell\vscode\command]
@="\"C:\\Program Files\\Microsoft VS Code\\Code.exe\" \"%V\""
"@
    $regFile = "$env:TEMP\VSCodeContextMenu.reg"
    Set-Content -Path $regFile -Value $MultilineComment -Force
    Regedit.exe /S $regFile

    Write-Host "✔ VS Code Context Menu Entries Added." -ForegroundColor Green

    Write-Host "`n[5/5] Configuration Complete." -ForegroundColor Green
}

function Show-Menu {
    Clear-Host
    Write-Host "==========================" -ForegroundColor Yellow
    Write-Host "     o9 Setup Installer   " -ForegroundColor Cyan
    Write-Host "==========================" -ForegroundColor Yellow
    Write-Host "1. Install o9 Theme"
    Write-Host "2. Install JetBrains Mono Font"
    Write-Host "3. Install Code Fonts"
    Write-Host "4. Install VS Code Config"
    Write-Host "5. Exit"
    Write-Host "--------------------------"
}

function Start-Installer {
    do {
        Show-Menu
        $choice = Read-Host "`nSelect an option (1-5)"
        switch ($choice) {
            '1' {
                Install-o9Theme
                Pause
            }
            '2' {
                Install-JetBrainsMono -FontName "JetBrainsMono" -FontDisplayName "JetBrainsMono NF"
                Pause
            }
            '3' {
                Install-CodeFonts
                Pause
            }
            '4' {
                Install-Config
                Pause
            }
            '5' {
                Write-Host "`nExiting installer..." -ForegroundColor DarkGray
            }
            default {
                Write-Host "Invalid option. Try again." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    } while ($choice -ne '5')
}

# Start the script
Start-Installer
