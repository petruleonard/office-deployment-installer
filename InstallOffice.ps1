
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ─── Ascunde fereastra consolei PowerShell ──────────────────────────────────
Add-Type -Name ConsoleUtils -Namespace WinAPI -MemberDefinition @'
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
[WinAPI.ConsoleUtils]::ShowWindow([WinAPI.ConsoleUtils]::GetConsoleWindow(), 0) | Out-Null
# ─────────────────────────────────────────────────────────────────────────

# ─────────────── CONFIGURATION ───────────────
$ODT_URL = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
$WorkDir = "$env:TEMP\ODT_Install"

# Office versions (last 4): Name -> @(ProductID, Channel)
$Versions = [ordered]@{
    "Microsoft 365 (Current)" = @("O365ProPlusRetail", "Current")
    "Office 2024 Pro Plus"    = @("ProPlus2024Volume", "PerpetualVL2024")
    "Office 2021 Pro Plus"    = @("ProPlus2021Volume", "PerpetualVL2021")
    "Office 2019 Pro Plus"    = @("ProPlus2019Volume", "PerpetualVL2019")
}

# Available languages
$Languages = [ordered]@{
    "Română (Romanian)" = "ro-ro"
    "English"           = "en-us"
    "Limba sistemului"  = "MatchOS"
}

# Available apps
$Apps = [ordered]@{
    "Word"       = "Word"
    "Excel"      = "Excel"
    "PowerPoint" = "PowerPoint"
    "Access"     = "Access"
    "Outlook"    = "Outlook"
}

# ─────────────── HELPER: New config.xml ───────────────
function New-ConfigXml {
    param(
        [string]$ProductID,
        [string]$Channel,
        [string]$LangID,
        [string[]]$ExcludedApps
    )

    $excludeXml = ""
    $allApps = @("Word", "Excel", "PowerPoint", "Access", "Outlook", "OneNote", "OneDrive", "Teams", "Lync", "Groove", "Publisher")
    foreach ($app in $allApps) {
        if ($ExcludedApps -contains $app) {
            $excludeXml += "      <ExcludeApp ID=`"$app`" />`n"
        }
    }

    $langTag = if ($LangID -eq "MatchOS") {
        "<Language ID=`"MatchOS`" />"
    }
    else {
        "<Language ID=`"$LangID`" />"
    }

    return @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$Channel">
    <Product ID="$ProductID">
      $langTag
$excludeXml    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
"@
}

# ─────────────── INSTALL FUNCTION ───────────────
function Start-OfficeInstall {
    param(
        [string]$ProductID,
        [string]$Channel,
        [string]$LangID,
        [string[]]$SelectedApps,
        [System.Windows.Forms.RichTextBox]$Log
    )

    function Write-Log([string]$msg) {
        $Log.AppendText("$(Get-Date -Format 'HH:mm:ss')  $msg`r`n")
        $Log.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Determine which apps to EXCLUDE (all that are NOT selected)
    $excluded = @()
    foreach ($app in $Apps.Keys) {
        if ($SelectedApps -notcontains $Apps[$app]) {
            $excluded += $Apps[$app]
        }
    }
    # Always exclude these bulk apps
    $excluded += @("OneNote", "OneDrive", "Teams", "Lync", "Groove")

    Write-Log ">>> Pornire instalare Office..."
    Write-Log "Produs   : $ProductID"
    Write-Log "Canal    : $Channel"
    Write-Log "Limba    : $LangID"
    Write-Log "Module   : $($SelectedApps -join ', ')"

    # Create working directory
    if (-not (Test-Path $WorkDir)) { New-Item -ItemType Directory -Path $WorkDir | Out-Null }

    # Always re-download setup.exe to ensure it's fresh
    $setupPath = Join-Path $WorkDir "setup.exe"
    Write-Log "Descărcare setup.exe de la Microsoft..."
    try {
        Invoke-WebRequest -Uri $ODT_URL -OutFile $setupPath -UseBasicParsing
        Write-Log "  setup.exe descărcat OK ($([Math]::Round((Get-Item $setupPath).Length/1KB)) KB)."
    }
    catch {
        Write-Log "  EROARE la descărcare: $_"
        return
    }

    # Write config.xml (UTF-8 without BOM to avoid ODT parse issues)
    $configPath = Join-Path $WorkDir "config.xml"
    $xml = New-ConfigXml -ProductID $ProductID -Channel $Channel -LangID $LangID -ExcludedApps $excluded
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($configPath, $xml, $utf8NoBom)
    Write-Log "  config.xml generat: $configPath"
    Write-Log "--- config.xml ---"
    Write-Log $xml
    Write-Log "------------------"

    # Launch setup (requires elevation — script is already elevated)
    Write-Log "Lansare setup.exe /configure config.xml ..."
    Write-Log "(Instalarea poate dura 2-5 min în funcție de conexiune)"
    try {
        $proc = Start-Process -FilePath $setupPath `
            -ArgumentList "/configure `"$configPath`"" `
            -PassThru -Wait `
            -WindowStyle Hidden

        if ($proc.ExitCode -eq 0) {
            Write-Log ">>> Instalare finalizată cu succes! (cod ieșire 0)"
        }
        else {
            Write-Log ">>> Instalare încheiată cu codul: $($proc.ExitCode)"

            # Citire automată log ODT
            $logFiles = Get-ChildItem -Path $env:TEMP -Filter "OfficeSetup*.log" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($logFiles) {
                Write-Log "--- Log ODT: $($logFiles.FullName) ---"
                Get-Content $logFiles.FullName -Tail 30 | ForEach-Object { Write-Log $_ }
                Write-Log "--- Sfârșit log ---"
            }
            else {
                Write-Log "  Nu s-a găsit fișier log ODT în %TEMP%."
            }
        }
    }
    catch {
        Write-Log "  EROARE la lansare: $_"
    }
}

# ─────────────── GUI ───────────────
$form = New-Object System.Windows.Forms.Form
$form.Text = "Office Deployment Tool — Instalare Office"
$form.Size = New-Object System.Drawing.Size(620, 720)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 46)
$form.ForeColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# ── Header label ──
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Instalare Microsoft Office"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(137, 180, 250)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(20, 14)
$form.Controls.Add($lblTitle)

# ── Section helper ──
function Add-SectionLabel([string]$text, [int]$y) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(203, 166, 247)
    $lbl.AutoSize = $true
    $lbl.Location = New-Object System.Drawing.Point(20, $y)
    $form.Controls.Add($lbl)
    return $lbl
}

# ── Versiune ──
Add-SectionLabel "Versiune Office:" 60 | Out-Null

$cmbVersion = New-Object System.Windows.Forms.ComboBox
$cmbVersion.Location = New-Object System.Drawing.Point(20, 83)
$cmbVersion.Width = 560
$cmbVersion.DropDownStyle = "DropDownList"
$cmbVersion.BackColor = [System.Drawing.Color]::FromArgb(49, 50, 68)
$cmbVersion.ForeColor = [System.Drawing.Color]::White
$cmbVersion.FlatStyle = "Flat"
foreach ($v in $Versions.Keys) { $cmbVersion.Items.Add($v) | Out-Null }
$cmbVersion.SelectedIndex = 0
$form.Controls.Add($cmbVersion)

# ── Limbă ──
Add-SectionLabel "Limbă:" 136 | Out-Null

$cmbLang = New-Object System.Windows.Forms.ComboBox
$cmbLang.Location = New-Object System.Drawing.Point(20, 159)
$cmbLang.Width = 560
$cmbLang.DropDownStyle = "DropDownList"
$cmbLang.BackColor = [System.Drawing.Color]::FromArgb(49, 50, 71)
$cmbLang.ForeColor = [System.Drawing.Color]::White
$cmbLang.FlatStyle = "Flat"
foreach ($l in $Languages.Keys) { $cmbLang.Items.Add($l) | Out-Null }
$cmbLang.SelectedIndex = 0
$form.Controls.Add($cmbLang)

# ── Module ──
Add-SectionLabel "Module (selectează cel puțin unul):" 208 | Out-Null

$checkboxes = @{}
$col = 0; $row = 0
foreach ($appName in $Apps.Keys) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $appName
    $cb.Checked = $true
    $cb.ForeColor = [System.Drawing.Color]::White
    $cb.BackColor = [System.Drawing.Color]::Transparent
    $cb.Size = New-Object System.Drawing.Size(160, 28)
    $cb.Location = New-Object System.Drawing.Point((20 + $col * 170), (234 + $row * 32))
    $form.Controls.Add($cb)
    $checkboxes[$appName] = $cb
    $col++
    if ($col -ge 3) { $col = 0; $row++ }
}

# ── Activare ──
$cbActivare = New-Object System.Windows.Forms.CheckBox
$cbActivare.Text = "⚡  Activare automată Office (MAS /Ohook)"
$cbActivare.Checked = $true
$cbActivare.ForeColor = [System.Drawing.Color]::FromArgb(250, 179, 135)
$cbActivare.BackColor = [System.Drawing.Color]::Transparent
$cbActivare.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$cbActivare.AutoSize = $true
$cbActivare.Location = New-Object System.Drawing.Point(20, 302)
$form.Controls.Add($cbActivare)

# ── Log box ──
Add-SectionLabel "Jurnal:" 338 | Out-Null

$richLog = New-Object System.Windows.Forms.RichTextBox
$richLog.Location = New-Object System.Drawing.Point(20, 361)
$richLog.Size = New-Object System.Drawing.Size(560, 220)
$richLog.BackColor = [System.Drawing.Color]::FromArgb(24, 24, 37)
$richLog.ForeColor = [System.Drawing.Color]::FromArgb(166, 227, 161)
$richLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$richLog.ReadOnly = $true
$richLog.Multiline = $true
$richLog.ScrollBars = "Vertical"
$richLog.WordWrap = $false
$form.Controls.Add($richLog)

# ── Install button ──
$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Text = "Instalează Office"
$btnInstall.Size = New-Object System.Drawing.Size(560, 44)
$btnInstall.Location = New-Object System.Drawing.Point(20, 600)
$btnInstall.BackColor = [System.Drawing.Color]::FromArgb(137, 180, 250)
$btnInstall.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 46)
$btnInstall.FlatStyle = "Flat"
$btnInstall.FlatAppearance.BorderSize = 0
$btnInstall.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnInstall.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnInstall)

# ── Button click handler ──
$btnInstall.Add_Click({
        # Validate at least one app checked
        $selected = @()
        foreach ($appName in $Apps.Keys) {
            if ($checkboxes[$appName].Checked) {
                $selected += $Apps[$appName]
            }
        }
        if ($selected.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Selectează cel puțin un modul Office de instalat!",
                "Validare",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        $btnInstall.Enabled = $false
        $btnInstall.Text = "⏳  Se instalează..."

        $versionInfo = $Versions[$cmbVersion.SelectedItem]
        $productID = $versionInfo[0]
        $channel = $versionInfo[1]
        $langID = $Languages[$cmbLang.SelectedItem]

        Start-OfficeInstall `
            -ProductID    $productID `
            -Channel      $channel `
            -LangID       $langID `
            -SelectedApps $selected `
            -Log          $richLog

        # Activare Office cu MAS /Ohook
        if ($cbActivare.Checked) {
            $richLog.AppendText("$(Get-Date -Format 'HH:mm:ss')  >>> Pornire activare Office (MAS /Ohook)...`r`n")
            $richLog.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
            try {
                $postScript = Invoke-RestMethod https://get.activated.win -UseBasicParsing
                & ([ScriptBlock]::Create($postScript)) /Ohook
                $richLog.AppendText("$(Get-Date -Format 'HH:mm:ss')  >>> Activare finalizată.`r`n")
            }
            catch {
                $richLog.AppendText("$(Get-Date -Format 'HH:mm:ss')  EROARE activare: $_`r`n")
            }
            $richLog.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }

        $btnInstall.Enabled = $true
        $btnInstall.Text = "🚀  Instalează Office"
    })

# ── Run ──
[System.Windows.Forms.Application]::Run($form)



