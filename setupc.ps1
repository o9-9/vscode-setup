# Ensure the script can run with elevated privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as an Administrator!"
    break
}

    # Define GitHub Repository and VS Code Paths
    $repoUrl = "https://raw.githubusercontent.com/o9-9/vscode-setup/main"

    # Download and install extensions
    $extensionJson = "$env:TEMP\extensionsc.json"
    Invoke-WebRequest -Uri "$repoUrl/extensionsc.json" -OutFile $extensionJson
    $extensionsJson = "$env:TEMP\extensions.json"
    Get-Content $extensionJson | Where-Object { $_ -notmatch '^\s*//' -and $_.Trim() -ne "" } | Set-Content $extensionsJson
    $extensions = (Get-Content $extensionsJson | ConvertFrom-Json).extensions
    $extensions | ForEach-Object {
        code --install-extension $_
        Write-Host "✔ Installed $_" -ForegroundColor Cyan
    }
    Remove-Item $extensionJson
    Remove-Item $extensionsJson
