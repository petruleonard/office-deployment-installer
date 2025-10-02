
# =============================================
# Office Deployment Tool Installer (PS 5.1, sobru & corporate)
# =============================================

param(
    [switch]$elevated
)
[System.Console]::CursorVisible = $false

# ---------- Initial Checks ----------
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
        Write-Host "Administrator rights were not granted. Script cannot continue." -ForegroundColor Red
        Start-Sleep -Seconds 2
        exit 1
    }
}

# ---------- Header ----------
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

# ---------- Progress Bar ----------
function Show-ProgressBar {
    param (
        [int]$Percent
    )

    $barLength = 44  # Bara fixă aliniată cu meniurile
    $filledLength = [math]::Round($Percent / 100 * $barLength)
    $emptyLength = $barLength - $filledLength

    $filled = ('*' * $filledLength)
    $empty  = ('-' * $emptyLength)

    Write-Host -NoNewline "`r[" -ForegroundColor DarkCyan
    Write-Host -NoNewline $filled -ForegroundColor Green
    Write-Host -NoNewline $empty -ForegroundColor DarkGray
    Write-Host -NoNewline "] " -ForegroundColor DarkCyan
    Write-Host -NoNewline ("{0,3}%" -f $Percent) -ForegroundColor Cyan
}

# ---------- Data ----------

$fixedApps = @("Word","Excel")

$products = [ordered]@{
    "1" = @{ Name="Office 365 Enterprise"; ID="O365ProPlusEEANoTeamsRetail"; PidKey="H8DN8-Y2YP3-CR9JT-DHDR9-C7GP3"; Channel="Current" }
    "2" = @{ Name="Office 365 Business";   ID="O365BusinessEEANoTeamsRetail"; PidKey="Y9NF9-M2QWD-FF6RJ-QJW36-RRF2T"; Channel="Current" }
    "3" = @{ Name="Office LTSC 2024 ProPlus"; ID="ProPlus2024Volume"; PidKey="4YV2J-VNG7W-YGTP3-443TK-TF8CP"; Channel="PerpetualVL2024" }
}

$languages = [ordered]@{
    "1" = "MatchOS"
    "2" = "en-US"
    "3" = "ro-RO"
}

$apps = [ordered]@{
    "1"="Access"; "2"="Publisher"; "3"="Outlook"; "4"="PowerPoint"
}

# ---------- Steps ----------
Show-Menu "[1/3] Select product:" $products
$prodChoice = Get-Choice $products "Enter option (1-$($products.Count))"
$productID  = $products[$prodChoice].ID
$productChannel  = $products[$prodChoice].Channel
$productPidKey = $products[$prodChoice].PidKey
$productDisplayName = $products[$prodChoice].Name

Show-Menu "[2/3] Select language:" $languages
$langChoice = Get-Choice $languages "Enter option (1-$($languages.Count))"
$language = $languages[$langChoice]

Show-Menu "[3/3] Select apps to exclude (Word and Excel will be installed by default):" $apps
$excludeApps = Get-MultiChoice $apps "Enter apps to exclude (e.g.1,3) or press Enter for none"

# ---------- Build XML ----------
$tempFolder = "$env:TEMP\OfficeODT"
$configPath = "$tempFolder\config.xml"
if (-not (Test-Path $tempFolder)) { New-Item -Path $tempFolder -ItemType Directory | Out-Null }

$xml = New-Object System.Xml.XmlDocument
$configNode = $xml.CreateElement("Configuration")
$xml.AppendChild($configNode) | Out-Null

$addNode = $xml.CreateElement("Add")
$addNode.SetAttribute("OfficeClientEdition", "64")
$addNode.SetAttribute("Channel", $productChannel)
$addNode.SetAttribute("PIDKEY", $productPidKey)
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

$xml.Save($configPath)
Write-Host ""
Write-Host "Configuration XML created: $configPath" -ForegroundColor Yellow

# ---------- Download setup.exe ----------
$setupUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
$setupPath = "$tempFolder\setup.exe"
$downloadSuccess = $false

Write-Host "`nDownloading setup.exe..." -ForegroundColor Cyan
Write-Host ""

$req = [System.Net.HttpWebRequest]::Create($setupUrl)
$resp = $req.GetResponse()
$totalBytes = $resp.ContentLength
$stream = $resp.GetResponseStream()
$fileStream = [System.IO.File]::Create($setupPath)

$buffer = New-Object byte[] 8192
$totalRead = 0
$bytesRead = 0

while (($bytesRead = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
    $fileStream.Write($buffer, 0, $bytesRead)
    $totalRead += $bytesRead

    $percent = [math]::Round(($totalRead / $totalBytes) * 100)
    Show-ProgressBar -Percent $percent
}

$fileStream.Close()
$stream.Close()
$downloadSuccess = $true

Write-Host ""
Write-Host "`nDownload complete: $setupPath" -ForegroundColor Yellow


# ========== SUMMARY ==========
$includedApps = $fixedApps + ($apps.Values | Where-Object { $excludeApps -notcontains $_ })
Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "         SUMMARY OF YOUR SELECTIONS        " -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host "Product:       $productDisplayName"
Write-Host "Language:      $language"
Write-Host "Channel:       $productChannel"
Write-Host "Included Apps: $($includedApps -join ', ')"
Write-Host "=============================================" -ForegroundColor Green
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
if (Test-Path $tempFolder) {
    try {
        Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction Stop
        Write-Host "Temporary files deleted automatically." -ForegroundColor Cyan
    } catch {
        Write-Host "ERROR: Temporary files could not be deleted automatically. Manual cleanup may be needed." -ForegroundColor Yellow
    }
} else {
    Write-Host "Temporary folder does not exist. Nothing to delete." -ForegroundColor Green
}

Start-Sleep -Seconds 3
