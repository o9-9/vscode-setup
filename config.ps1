if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Output "o9 Config to be run as Administrator."
    $argList = @()

    $PSBoundParameters.GetEnumerator() | ForEach-Object {
        $argList += if ($_.Value -is [switch] -and $_.Value) {
            "-$($_.Key)"
        } elseif ($_.Value -is [array]) {
            "-$($_.Key) $($_.Value -join ',')"
        } elseif ($_.Value) {
            "-$($_.Key) '$($_.Value)'"
        }
    }

    $script = if ($PSCommandPath) {
        "& { & `'$($PSCommandPath)`' $($argList -join ' ') }"
    } else {
        "&([ScriptBlock]::Create((irm https://github.com/o9-9/vscode-setup/releases/download/9/config.ps1))) $($argList -join ' ')"
    }

    $powershellCmd = if (Get-Command pwsh -ErrorAction SilentlyContinue) { "pwsh" } else { "powershell" }
    $processCmd = if (Get-Command wt.exe -ErrorAction SilentlyContinue) { "wt.exe" } else { "$powershellCmd" }

    if ($processCmd -eq "wt.exe") {
        Start-Process $processCmd -ArgumentList "$powershellCmd -ExecutionPolicy Bypass -NoProfile -Command `"$script`"" -Verb RunAs
    } else {
        Start-Process $processCmd -ArgumentList "-ExecutionPolicy Bypass -NoProfile -Command `"$script`"" -Verb RunAs
    }

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
    $fontUrl = "https://github.com/ryanoasis/nerd-fonts/releases/download/v3.2.1/JetBrainsMono.zip"
    $zipPath = "$env:TEMP\JetBrainsMono.zip"
    $extractPath = "$env:TEMP\JetBrainsMono"
    $fontsPath = "$env:WINDIR\Fonts"

    Write-Host "`n[1/3] Downloading JetBrainsMono Nerd Font..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $fontUrl -OutFile $zipPath -UseBasicParsing
        Write-Host "[OK] Downloaded font zip." -ForegroundColor Green
    } catch {
        Write-Error "Failed to download fonts: $_"
        return
    }

    Write-Host "`n[2/3] Extracting font files..." -ForegroundColor Cyan
    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force
    }
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractPath -Force

    Write-Host "`n[3/3] Installing fonts..." -ForegroundColor Cyan
    try {
        $fonts = Get-ChildItem "$extractPath" -Include *.ttf -Recurse
        foreach ($font in $fonts) {
            $destFontPath = Join-Path -Path $fontsPath -ChildPath $font.Name
            Copy-Item -Path $font.FullName -Destination $destFontPath -Force

            $regFontName = [System.IO.Path]::GetFileNameWithoutExtension($font.Name)
            New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" `
                -Name "$regFontName (TrueType)" -Value $font.Name -PropertyType String -Force | Out-Null
        }
        Write-Host "[OK] JetBrainsMono Nerd Fonts installed successfully!" -ForegroundColor Green
    } catch {
        Write-Error "Font installation failed: $_"
    }
}

function Install-CodeFonts {
    $fontUrl = "https://github.com/o9-9/vscode-setup/releases/download/9/CodeFonts.zip"
    $zipPath = "$env:TEMP\CodeFonts.zip"
    $extractPath = "$env:TEMP\CodeFonts"
    $fontsPath = "$env:WINDIR\Fonts"

    Write-Host "`n[1/3] Downloading Code Fonts..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $fontUrl -OutFile $zipPath -UseBasicParsing
        Write-Host "[OK] Downloaded font zip." -ForegroundColor Green
    } catch {
        Write-Error "Failed to download fonts: $_"
        return
    }

    Write-Host "`n[2/3] Extracting font files..." -ForegroundColor Cyan
    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force
    }
    Expand-Archive -LiteralPath $zipPath -DestinationPath $extractPath -Force

    Write-Host "`n[3/3] Installing fonts..." -ForegroundColor Cyan
    try {
        $fonts = Get-ChildItem "$extractPath" -Include *.ttf -Recurse
        foreach ($font in $fonts) {
            $destFontPath = Join-Path -Path $fontsPath -ChildPath $font.Name
            Copy-Item -Path $font.FullName -Destination $destFontPath -Force

            $regFontName = [System.IO.Path]::GetFileNameWithoutExtension($font.Name)
            New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" `
                -Name "$regFontName (TrueType)" -Value $font.Name -PropertyType String -Force | Out-Null
        }
        Write-Host "[OK] Code Fonts installed successfully!" -ForegroundColor Green
    } catch {
        Write-Error "Font installation failed: $_"
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
    Write-Host -ForegroundColor Green "✔ Copied Settings.json to VS Code." -ForegroundColor Green

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
        Write-Host -ForegroundColor Cyan "✔ Installed $_"
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
                Install-JetBrainsMono
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
