# ===============================
# Office Deployment Tool Installer (Native XML)
# Improved Version
# ===============================

# ---------- Auto-run as Administrator ----------
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $IsAdmin) {
    Write-Host "The script does not have Administrator rights. I am trying to relaunch it..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    # Obține calea către scriptul curent
    $scriptPath = $MyInvocation.MyCommand.Definition

    # Relansează scriptul ca Administrator
    Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    
    exit 0  # Oprește instanța curentă
}


Clear-Host
Write-Host "============================================" -ForegroundColor DarkCyan
Write-Host "      Office Deployment Tool Installer     " -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor DarkCyan
Write-Host ""

# ---------- Data ----------
$products = [ordered]@{
    "1" = @{ Name="Office 365 Enterprise"; ID="O365ProPlusEEANoTeamsRetail"; Channel="Current" }
    "2" = @{ Name="Office 365 Business";   ID="O365BusinessEEANoTeamsRetail"; Channel="Current" }
    "3" = @{ Name="Office LTSC 2024 ProPlus"; ID="ProPlus2024Volume"; Channel="PerpetualVL2024" }
}

$languages = [ordered]@{
    "1" = "MatchOS"
    "2" = "en-US"
    "3" = "ro-RO"
}

$apps = [ordered]@{
    "1"="Access"; "2"="Publisher"; "3"="Outlook"; "4"="PowerPoint"
}

$appSettingsMap = [ordered]@{
    "Word"      = @{ Key="software\microsoft\office\16.0\word\options"; Name="defaultformat"; Value=""; Type="REG_SZ"; App="word16"; Id="L_SaveWordfilesas" }
    "Excel"     = @{ Key="software\microsoft\office\16.0\excel\options"; Name="defaultformat"; Value="51"; Type="REG_DWORD"; App="excel16"; Id="L_SaveExcelfilesas" }
    "PowerPoint"= @{ Key="software\microsoft\office\16.0\powerpoint\options"; Name="defaultformat"; Value="27"; Type="REG_DWORD"; App="ppt16"; Id="L_SavePowerPointfilesas" }
    "Outlook"   = @{ Key="software\microsoft\office\16.0\outlook\options"; Name="cachedmode"; Value="0"; Type="REG_DWORD"; App="ol16"; Id="L_CachedMode" }
    "Access"    = @{ Key="software\microsoft\office\16.0\access\options"; Name="showstatusbar"; Value="1"; Type="REG_DWORD"; App="acc16"; Id="L_ShowStatusBar" }
    "Publisher" = @{ Key="software\microsoft\office\16.0\publisher\options"; Name="pubdefaultview"; Value="1"; Type="REG_DWORD"; App="pub16"; Id="L_DefaultView" }
}

# ---------- Functions ----------
function Show-Menu($title, $options) {
    Write-Host ""
    Write-Host $title -ForegroundColor Cyan
    Write-Host "--------------------------------------------"
    foreach ($key in $options.Keys) {
        if ($options[$key] -is [string]) {
            Write-Host "$key) $($options[$key])"
        } else {
            Write-Host "$key) $($options[$key].Name)"
        }
    }
    Write-Host "--------------------------------------------"
}

function Get-Choice($options, $prompt) {
    $choice = Read-Host -Prompt $prompt
    if ($options.Keys -contains $choice) { return $choice }
    Write-Host "Invalid choice, exiting..." -ForegroundColor Red
    exit 1
}

# ---------- Step 1: Product ----------
Show-Menu "[1/3] Select product:" $products
$prodChoice = Get-Choice $products "Enter option (1-$($products.Count))"
$productID  = $products[$prodChoice].ID
$productChannel  = $products[$prodChoice].Channel
$productDisplayName = $products[$prodChoice].Name

# ---------- Step 2: Language ----------
Show-Menu "[2/3] Select language:" $languages
$langChoice = Get-Choice $languages "Enter option (1-$($languages.Count))"
$language = $languages[$langChoice]

# ---------- Step 3: Apps ----------
Show-Menu "[3/3] Select apps to exclude (comma separated, or Enter for none):" $apps
$excludeChoice = Read-Host -Prompt "Enter option (ex: 1,3)"
$excludeApps = @()

if ($excludeChoice) {
    foreach ($key in $excludeChoice.Split(",") | ForEach-Object { $_.Trim() }) {
        if ($apps.Keys -contains $key) {
            $excludeApps += $apps[$key]
        } else {
            Write-Host "Warning: Invalid app selection '$key' ignored." -ForegroundColor Yellow
        }
    }
}

# ---------- Step 4: Build native XML ----------
$xml = New-Object System.Xml.XmlDocument

# Root Configuration
$configNode = $xml.CreateElement("Configuration")
$xml.AppendChild($configNode) | Out-Null

# Add node
$addNode = $xml.CreateElement("Add")
$addNode.SetAttribute("OfficeClientEdition", "64")
$addNode.SetAttribute("Channel", $productChannel)
$configNode.AppendChild($addNode) | Out-Null

# Product node
$productNode = $xml.CreateElement("Product")
$productNode.SetAttribute("ID", $productID)
$addNode.AppendChild($productNode) | Out-Null

# Language node
$langNode = $xml.CreateElement("Language")
$langNode.SetAttribute("ID", $language)
$productNode.AppendChild($langNode) | Out-Null

# Excluded apps
foreach ($app in $excludeApps) {
    $excludeNode = $xml.CreateElement("ExcludeApp")
    $excludeNode.SetAttribute("ID", $app)
    $productNode.AppendChild($excludeNode) | Out-Null
}

# Always exclude these
foreach ($app in "Groove","OneDrive","Lync","OneNote") {
    $excludeNode = $xml.CreateElement("ExcludeApp")
    $excludeNode.SetAttribute("ID", $app)
    $productNode.AppendChild($excludeNode) | Out-Null
}

# Display node
$displayNode = $xml.CreateElement("Display")
$displayNode.SetAttribute("Level", "Full")
$displayNode.SetAttribute("AcceptEULA", "TRUE")
$configNode.AppendChild($displayNode) | Out-Null

# Updates node
$updatesNode = $xml.CreateElement("Updates")
$updatesNode.SetAttribute("Enabled", "TRUE")
$configNode.AppendChild($updatesNode) | Out-Null

# RemoveMSI node
$removeNode = $xml.CreateElement("RemoveMSI")
$configNode.AppendChild($removeNode) | Out-Null

# AppSettings node
$appSettingsNode = $xml.CreateElement("AppSettings")
$configNode.AppendChild($appSettingsNode) | Out-Null

foreach ($app in $appSettingsMap.Keys) {
    if ($excludeApps -notcontains $app) {
        $setting = $appSettingsMap[$app]
        $userNode = $xml.CreateElement("User")
        $userNode.SetAttribute("Key", $setting.Key)
        $userNode.SetAttribute("Name", $setting.Name)
        $userNode.SetAttribute("Value", $setting.Value)
        $userNode.SetAttribute("Type", $setting.Type)
        $userNode.SetAttribute("App", $setting.App)
        $userNode.SetAttribute("Id", $setting.Id)
        $appSettingsNode.AppendChild($userNode) | Out-Null
    }
}

# ========== SUMMARY ==========
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "        SUMMARY OF YOUR SELECTIONS        " -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ("Product:      {0} (ID: {1})" -f $productDisplayName, $productID)
Write-Host "Language:     $language"
Write-Host "Channel:      $productChannel"
Write-Host "Excluded Apps: $($excludeApps -join ', ')"
Write-Host "============================================" -ForegroundColor Green
Write-Host ""

# Save XML
$folderPath = "$env:TEMP\OfficeODT"
if (-not (Test-Path $folderPath)) { New-Item -Path $folderPath -ItemType Directory | Out-Null }
$configPath = "$folderPath\config.xml"
$xml.Save($configPath)
Write-Host "Configuration XML created: $configPath" -ForegroundColor Green

# ---------- Step 5: Download setup ----------
$setupUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
$setupPath = "$folderPath\setup.exe"
$downloadSuccess = $false

try {
    # Creează fișier gol
    $fileStream = [System.IO.File]::Create($setupPath)

    # Obține dimensiunea totală
    $response = Invoke-WebRequest -Uri $setupUrl -Method Head
    $totalBytes = [int64]$response.Headers['Content-Length']

    # Deschide stream pentru citire
    $reader = [System.Net.HttpWebRequest]::Create($setupUrl).GetResponse().GetResponseStream()
    $buffer = New-Object byte[] 8192
    $totalRead = 0
    $bytesRead = 0
    $barLength = 50

    Write-Host "Downloading setup.exe..." -NoNewline

    while (($bytesRead = $reader.Read($buffer, 0, $buffer.Length)) -gt 0) {
        $fileStream.Write($buffer, 0, $bytesRead)
        $totalRead += $bytesRead

        $percent = [math]::Round(($totalRead / $totalBytes) * 100)
        $filled = [math]::Round($barLength * $percent / 100)
        $empty = $barLength - $filled
        $bar = ("*" * $filled) + ("-" * $empty)

        Write-Host "`r[$bar] $percent% " -NoNewline
    }

    $reader.Close()
    $fileStream.Close()
    Write-Host "`nDownload complete: $setupPath" -ForegroundColor Green
    $downloadSuccess = $true
}
catch {
    Write-Host "`nERROR: Download failed - $_" -ForegroundColor Red
    exit 1
}

# ---------- Step 6: Install ----------
$installSuccess = $false
if ($downloadSuccess) {
    try {
        Write-Host "Installing Office..." -ForegroundColor Yellow
        & $setupPath /configure $configPath
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Office installation complete!" -ForegroundColor Green
            $installSuccess = $true
        } else {
            Write-Host "Office installation failed! (Exit code $LASTEXITCODE)" -ForegroundColor Red
            $installSuccess = $false
        }
    }
    catch {
        Write-Host "ERROR: Installation failed - $_" -ForegroundColor Red
        $installSuccess = $false
    }
}

# ---------- Step 7: Cleanup ----------
if ($installSuccess) {
    try {
        Remove-Item -Path $folderPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Temporary files deleted." -ForegroundColor Cyan
    } catch {
        Write-Host "Warning: Could not delete temporary files. Manual cleanup may be needed." -ForegroundColor Yellow
    }
} else {
    Write-Host "Temporary files NOT deleted for troubleshooting." -ForegroundColor Yellow
}

# ---------- Step 8: Info ----------
Write-Host ""
Write-Host "============================================" -ForegroundColor Yellow
Write-Host "NOTE: Classic Outlook is part of the Office suite." -ForegroundColor Cyan
Write-Host "      New Outlook for Windows is NOT included." -ForegroundColor Cyan
Write-Host ""
Write-Host "For the new Outlook, install from Microsoft Store" -ForegroundColor Green
Write-Host "or quickly with winget using the command:" -ForegroundColor Green
Write-Host ""
Write-Host "    winget install -e --id=9NRX63209R7B --source=msstore --accept-package-agreements" -ForegroundColor White
Write-Host "============================================" -ForegroundColor Yellow
