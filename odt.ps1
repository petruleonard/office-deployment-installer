
# =============================================
# Office Deployment Tool Installer (Native XML)
# =============================================

param(
    [switch]$elevated
)
[System.Console]::CursorVisible = $false

# ---------- Initial Checks ----------

# Check and run the script as admin if required
$myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator

if (-not $myWindowsPrincipal.IsInRole($adminRole)) {
    Write-Host "Restarting script with administrator privileges..."
        Start-Sleep -Seconds 1
cls 
    try {
        $newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"
        $newProcess.Arguments = $myInvocation.MyCommand.Definition
        $newProcess.Verb = "runas"
        [System.Diagnostics.Process]::Start($newProcess) | Out-Null
        exit
    } catch {
        Write-Host "Administrator rights were not granted. The script cannot continue." -ForegroundColor Red
        Start-Sleep -Seconds 2
        exit 1
    }
}

# ---------- The normal script continues here ----------

Clear-Host
Write-Host "============================================" -ForegroundColor DarkCyan
Write-Host "      Office Deployment Tool Installer     " -ForegroundColor Yellow
Write-Host "============================================" -ForegroundColor DarkCyan
Write-Host ""

# ---------- UI Functions ----------
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
    while ($true) {
        $choice = Read-Host -Prompt $prompt
        if ($options.Keys -contains $choice) { return $choice }
        [Console]::SetCursorPosition(0, [Console]::CursorTop -1)
        Write-Host (" " * ($prompt.Length + $choice.Length + 5)) -NoNewline
        [Console]::SetCursorPosition(0, [Console]::CursorTop)
    }
}

function Get-MultiChoice($options, $prompt) {
    $validKeys = $options.Keys | ForEach-Object { [string]$_ }
    while ($true) {
        $raw = Read-Host -Prompt $prompt
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $parts = $raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        $invalid = $parts | Where-Object { $validKeys -notcontains $_ }
        if (-not $invalid) {
            return $parts | ForEach-Object { $options[$_] }
        }
        [Console]::SetCursorPosition(0, [Console]::CursorTop -1)
        Write-Host (" " * ($prompt.Length + $raw.Length + 5)) -NoNewline
        [Console]::SetCursorPosition(0, [Console]::CursorTop)
    }
}

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

# ---------- Steps ----------
Show-Menu "[1/3] Select product:" $products
$prodChoice = Get-Choice $products "Enter option (1-$($products.Count))"
$productID  = $products[$prodChoice].ID
$productChannel  = $products[$prodChoice].Channel
$productDisplayName = $products[$prodChoice].Name

Show-Menu "[2/3] Select language:" $languages
$langChoice = Get-Choice $languages "Enter option (1-$($languages.Count))"
$language = $languages[$langChoice]

Show-Menu "[3/3] Select apps to exclude (Word and Excel will be installed by default):" $apps
$excludeApps = Get-MultiChoice $apps "Enter apps to exclude (e.g.1,3) or press Enter for none"


# ---------- Build XML ----------
$xml = New-Object System.Xml.XmlDocument
$configNode = $xml.CreateElement("Configuration")
$xml.AppendChild($configNode) | Out-Null

$addNode = $xml.CreateElement("Add")
$addNode.SetAttribute("OfficeClientEdition", "64")
$addNode.SetAttribute("Channel", $productChannel)
$configNode.AppendChild($addNode) | Out-Null

$productNode = $xml.CreateElement("Product")
$productNode.SetAttribute("ID", $productID)
$addNode.AppendChild($productNode) | Out-Null

$langNode = $xml.CreateElement("Language")
$langNode.SetAttribute("ID", $language)
$productNode.AppendChild($langNode) | Out-Null

foreach ($app in $excludeApps + @("Groove","OneDrive","Lync","OneNote")) {
    $excludeNode = $xml.CreateElement("ExcludeApp")
    $excludeNode.SetAttribute("ID", $app)
    $productNode.AppendChild($excludeNode) | Out-Null
}

$displayNode = $xml.CreateElement("Display")
$displayNode.SetAttribute("Level", "Full")
$displayNode.SetAttribute("AcceptEULA", "TRUE")
$configNode.AppendChild($displayNode) | Out-Null

$updatesNode = $xml.CreateElement("Updates")
$updatesNode.SetAttribute("Enabled", "TRUE")
$configNode.AppendChild($updatesNode) | Out-Null

$configPath = "$env:TEMP\OfficeODT\config.xml"
if (-not (Test-Path "$env:TEMP\OfficeODT")) { New-Item -Path "$env:TEMP\OfficeODT" -ItemType Directory | Out-Null }
$xml.Save($configPath)
Write-Host "Configuration XML created: $configPath" -ForegroundColor yellow

# ---------- Download setup with colorful progress bar ----------
$setupUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
$setupPath = "$env:TEMP\OfficeODT\setup.exe"
$downloadSuccess = $false

# Ensure folder exists
$folderPath = "$env:TEMP\OfficeODT"
if (-not (Test-Path $folderPath)) { New-Item -Path $folderPath -ItemType Directory | Out-Null }

try {
    # Get total size
    $response = Invoke-WebRequest -Uri $setupUrl -Method Head
    $totalBytes = [int64]$response.Headers['Content-Length']

    $reader = [System.Net.HttpWebRequest]::Create($setupUrl).GetResponse().GetResponseStream()
    $fileStream = [System.IO.File]::Create($setupPath)
    $buffer = New-Object byte[] 8192
    $totalRead = 0
    $bytesRead = 0

    Write-Host "Downloading setup.exe..."

while (($bytesRead = $reader.Read($buffer, 0, $buffer.Length)) -gt 0) {
    $fileStream.Write($buffer, 0, $bytesRead)
    $totalRead += $bytesRead

     $percent = [math]::Round(($totalRead / $totalBytes) * 100)

# Determine console width dynamically
     $width = $Host.UI.RawUI.WindowSize.Width
# Reserve ~10 chars for percent text and brackets
     $barLength = [math]::Max(10, $width - 10)

     $filled = [math]::Round($barLength * $percent / 100)
     $empty  = $barLength - $filled

     $barFilled = "*" * $filled
     $barEmpty  = "-" * $empty

# Draw colorful progress bar
        Write-Host "`r[" -NoNewline
        Write-Host $barFilled -NoNewline -ForegroundColor Green
        Write-Host $barEmpty -NoNewline -ForegroundColor DarkGray
        Write-Host "] $percent% " -NoNewline
}

    $reader.Close()
    $fileStream.Close()
    Write-Host "`nDownload complete: $setupPath" -ForegroundColor Yellow
    $downloadSuccess = $true
}
catch {
    Write-Host "`nERROR: Download failed - $_" -ForegroundColor Red
    exit 1
}

# ========== SUMMARY ==========
$includedApps = $apps.Values | Where-Object { $excludeApps -notcontains $_ }
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "         SUMMARY OF YOUR SELECTIONS        " -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host "Product:       $productDisplayName"
Write-Host "Language:      $language"
Write-Host "Channel:       $productChannel"
Write-Host "Included Apps: Word Excell $($includedApps -join ', ')"
Write-Host "============================================" -ForegroundColor Green
Write-Host ""

$response = Read-Host "Do you wish to continue? (y/n)"

if ($response -eq "y" -or $response -eq "Y") {
    Write-Host "Continuing script execution..."


# ---------- Install ----------
if ($downloadSuccess) {
    try {
        Write-Host "Installing Office..." -ForegroundColor Yellow
         & $setupPath /configure $configPath
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Office installation complete!" -ForegroundColor Green
        } else {
            Write-Host "Office installation failed! (Exit code $LASTEXITCODE)" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "ERROR: Installation failed - $_" -ForegroundColor Red
    }
}

} else {
    Write-Host "Script execution has been stopped."
}

# ---------- Cleanup ----------
try {
    Remove-Item -Path $folderPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Temporary files deleted." -ForegroundColor Cyan
} catch {
    Write-Host "Temporary files were not deleted. Manual cleanup may be needed." -ForegroundColor Yellow
}
Start-Sleep -Seconds 2
