Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "SilentlyContinue"

# =================== DOUBLE BUFFERING ===================
Add-Type @"
using System;
using System.Windows.Forms;
public class DoubleBufferedDGV : DataGridView {
    public DoubleBufferedDGV() {
        this.DoubleBuffered = true;
        this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
    }
}
public class DoubleBufferedPanel : Panel {
    public DoubleBufferedPanel() {
        this.DoubleBuffered = true;
        this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
    }
}
"@ -ReferencedAssemblies System.Windows.Forms, System.Drawing

# =================== GLOBAL STATE ===================
$script:isRunning = $false
$script:mailMap = @{}
$script:allEmails = [System.Collections.ArrayList]::new()
$script:exclusionFilters = [System.Collections.ArrayList]::new()
$script:matchRules = [System.Collections.ArrayList]::new()
$script:alertClearMap = @{}
$script:allDomains = @{}
$script:processedIds = @{}
$script:folderScanLog = [System.Collections.ArrayList]::new()
$script:selectedDomains = @()

# Stats
$script:pieCritical = 0
$script:pieAttention = 0
$script:pieTrouble = 0
$script:pieCleared = 0

# =================== THEMES ===================
$script:themes = @{
    "Cyber Night" = @{
        name        = "Cyber Night"
        bgDeep      = [System.Drawing.Color]::FromArgb(10, 14, 23)
        bgMain      = [System.Drawing.Color]::FromArgb(15, 20, 32)
        bgSidebar   = [System.Drawing.Color]::FromArgb(12, 17, 28)
        bgCard      = [System.Drawing.Color]::FromArgb(22, 28, 42)
        bgCardHover = [System.Drawing.Color]::FromArgb(28, 36, 54)
        bgInput     = [System.Drawing.Color]::FromArgb(18, 24, 38)
        border      = [System.Drawing.Color]::FromArgb(40, 55, 85)
        borderLight = [System.Drawing.Color]::FromArgb(55, 75, 110)
        textPrimary = [System.Drawing.Color]::FromArgb(235, 240, 250)
        textSecond  = [System.Drawing.Color]::FromArgb(140, 160, 190)
        textMuted   = [System.Drawing.Color]::FromArgb(75, 95, 125)
        accent      = [System.Drawing.Color]::FromArgb(0, 195, 255)
        accentGlow  = [System.Drawing.Color]::FromArgb(0, 150, 200)
        critical    = [System.Drawing.Color]::FromArgb(255, 71, 87)
        attention   = [System.Drawing.Color]::FromArgb(255, 195, 0)
        trouble     = [System.Drawing.Color]::FromArgb(255, 135, 50)
        clear       = [System.Drawing.Color]::FromArgb(46, 213, 115)
        gridAlt     = [System.Drawing.Color]::FromArgb(18, 24, 38)
    }
    "Purple Haze" = @{
        name        = "Purple Haze"
        bgDeep      = [System.Drawing.Color]::FromArgb(15, 10, 25)
        bgMain      = [System.Drawing.Color]::FromArgb(22, 15, 38)
        bgSidebar   = [System.Drawing.Color]::FromArgb(18, 12, 32)
        bgCard      = [System.Drawing.Color]::FromArgb(35, 25, 55)
        bgCardHover = [System.Drawing.Color]::FromArgb(45, 32, 70)
        bgInput     = [System.Drawing.Color]::FromArgb(28, 20, 45)
        border      = [System.Drawing.Color]::FromArgb(70, 50, 100)
        borderLight = [System.Drawing.Color]::FromArgb(90, 65, 130)
        textPrimary = [System.Drawing.Color]::FromArgb(245, 235, 255)
        textSecond  = [System.Drawing.Color]::FromArgb(175, 155, 200)
        textMuted   = [System.Drawing.Color]::FromArgb(100, 80, 130)
        accent      = [System.Drawing.Color]::FromArgb(190, 90, 255)
        accentGlow  = [System.Drawing.Color]::FromArgb(150, 70, 210)
        critical    = [System.Drawing.Color]::FromArgb(255, 71, 87)
        attention   = [System.Drawing.Color]::FromArgb(255, 195, 0)
        trouble     = [System.Drawing.Color]::FromArgb(255, 135, 50)
        clear       = [System.Drawing.Color]::FromArgb(46, 213, 115)
        gridAlt     = [System.Drawing.Color]::FromArgb(28, 20, 45)
    }
    "Matrix" = @{
        name        = "Matrix"
        bgDeep      = [System.Drawing.Color]::FromArgb(5, 15, 8)
        bgMain      = [System.Drawing.Color]::FromArgb(8, 22, 12)
        bgSidebar   = [System.Drawing.Color]::FromArgb(6, 18, 10)
        bgCard      = [System.Drawing.Color]::FromArgb(15, 35, 20)
        bgCardHover = [System.Drawing.Color]::FromArgb(20, 45, 28)
        bgInput     = [System.Drawing.Color]::FromArgb(10, 28, 15)
        border      = [System.Drawing.Color]::FromArgb(30, 80, 45)
        borderLight = [System.Drawing.Color]::FromArgb(45, 110, 60)
        textPrimary = [System.Drawing.Color]::FromArgb(180, 255, 200)
        textSecond  = [System.Drawing.Color]::FromArgb(100, 200, 130)
        textMuted   = [System.Drawing.Color]::FromArgb(50, 120, 70)
        accent      = [System.Drawing.Color]::FromArgb(0, 255, 120)
        accentGlow  = [System.Drawing.Color]::FromArgb(0, 200, 95)
        critical    = [System.Drawing.Color]::FromArgb(255, 71, 87)
        attention   = [System.Drawing.Color]::FromArgb(255, 195, 0)
        trouble     = [System.Drawing.Color]::FromArgb(255, 135, 50)
        clear       = [System.Drawing.Color]::FromArgb(0, 255, 120)
        gridAlt     = [System.Drawing.Color]::FromArgb(10, 28, 15)
    }
    "Light Mode" = @{
        name        = "Light Mode"
        bgDeep      = [System.Drawing.Color]::FromArgb(242, 244, 248)
        bgMain      = [System.Drawing.Color]::FromArgb(250, 251, 254)
        bgSidebar   = [System.Drawing.Color]::FromArgb(255, 255, 255)
        bgCard      = [System.Drawing.Color]::FromArgb(255, 255, 255)
        bgCardHover = [System.Drawing.Color]::FromArgb(245, 247, 252)
        bgInput     = [System.Drawing.Color]::FromArgb(255, 255, 255)
        border      = [System.Drawing.Color]::FromArgb(218, 222, 232)
        borderLight = [System.Drawing.Color]::FromArgb(200, 205, 218)
        textPrimary = [System.Drawing.Color]::FromArgb(25, 30, 45)
        textSecond  = [System.Drawing.Color]::FromArgb(85, 95, 115)
        textMuted   = [System.Drawing.Color]::FromArgb(150, 160, 175)
        accent      = [System.Drawing.Color]::FromArgb(25, 120, 220)
        accentGlow  = [System.Drawing.Color]::FromArgb(20, 95, 180)
        critical    = [System.Drawing.Color]::FromArgb(220, 53, 69)
        attention   = [System.Drawing.Color]::FromArgb(230, 160, 0)
        trouble     = [System.Drawing.Color]::FromArgb(220, 110, 30)
        clear       = [System.Drawing.Color]::FromArgb(40, 167, 69)
        gridAlt     = [System.Drawing.Color]::FromArgb(248, 249, 252)
    }
}

$script:currentTheme = "Cyber Night"
$theme = $script:themes[$script:currentTheme]

# =================== MATCH TYPES ===================
$script:matchFieldTypes = @(
    "Combined (All Fields)",
    "IP Address",
    "Device Name",
    "Serial Number", 
    "Category",
    "Hostname",
    "Subject Pattern",
    "Custom Regex"
)

# =================== SOUND FUNCTION ===================
function Play-SeveritySound {
    param([string]$Severity)
    if (-not $script:chkSound -or -not $script:chkSound.Checked) { return }
    try {
        switch ($Severity) {
            "CRITICAL" { 
                [Console]::Beep(1000, 350); Start-Sleep -Milliseconds 80
                [Console]::Beep(1000, 350); Start-Sleep -Milliseconds 80
                [Console]::Beep(1200, 450)
            }
            "ATTENTION" { 
                [Console]::Beep(800, 280); Start-Sleep -Milliseconds 70
                [Console]::Beep(900, 320)
            }
            "TROUBLE" { [Console]::Beep(600, 220) }
        }
    } catch {}
}

# =================== UTILITY FUNCTIONS ===================
function Get-EmailDomain {
    param([string]$Email)
    if ([string]::IsNullOrWhiteSpace($Email)) { return "unknown" }
    if ($Email -match "@(.+)$") { return $matches[1].ToLower().Trim() }
    if ($Email -match "^/O=([^/]+)") { return ("exchange_" + $matches[1]).ToLower() }
    $clean = $Email -replace '[^a-zA-Z0-9]', ''
    if ($clean.Length -gt 15) { $clean = $clean.Substring(0,15) }
    return "other_$clean".ToLower()
}

function Extract-IPAddresses {
    param([string]$Text)
    $ips = @()
    $matches = [regex]::Matches($Text, '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})')
    foreach ($m in $matches) {
        $ip = $m.Groups[1].Value
        $octets = $ip -split '\.'
        $valid = $true
        foreach ($o in $octets) { if ([int]$o -gt 255) { $valid = $false; break } }
        if ($valid -and $ip -notin $ips) { $ips += $ip }
    }
    return $ips
}

function Extract-DeviceName {
    param([string]$Text)
    $devices = @()
    $patterns = @('Device\s*(?:Name)?\s*[:=]\s*([^\r\n,;]+)', 'Hostname\s*[:=]\s*([^\r\n,;]+)')
    foreach ($p in $patterns) {
        $ms = [regex]::Matches($Text, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($m in $ms) {
            $d = $m.Groups[1].Value.Trim()
            if ($d -and $d.Length -gt 1 -and $d -notin $devices) { $devices += $d }
        }
    }
    return $devices
}

function Extract-SerialNumber {
    param([string]$Text)
    $serials = @()
    $patterns = @('Serial\s*(?:Number|No\.?)?\s*[:=]\s*([A-Za-z0-9\-]+)', 'S/?N\s*[:=]\s*([A-Za-z0-9\-]+)')
    foreach ($p in $patterns) {
        $ms = [regex]::Matches($Text, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($m in $ms) {
            $s = $m.Groups[1].Value.Trim()
            if ($s -and $s.Length -gt 3 -and $s -notin $serials) { $serials += $s }
        }
    }
    return $serials
}

function Extract-Category {
    param([string]$Text)
    $cats = @()
    $patterns = @('Category\s*[:=]\s*([^\r\n,;]+)', 'Type\s*[:=]\s*([^\r\n,;]+)')
    foreach ($p in $patterns) {
        $ms = [regex]::Matches($Text, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($m in $ms) {
            $c = $m.Groups[1].Value.Trim()
            if ($c -and $c.Length -gt 1 -and $c -notin $cats) { $cats += $c }
        }
    }
    return $cats
}

function Extract-CombinedKey {
    param([string]$Text)
    $keys = @()
    foreach ($ip in (Extract-IPAddresses -Text $Text)) { $keys += "IP:$($ip.ToLower())" }
    foreach ($d in (Extract-DeviceName -Text $Text)) { $keys += "DEV:$($d.ToLower())" }
    foreach ($s in (Extract-SerialNumber -Text $Text)) { $keys += "SN:$($s.ToLower())" }
    foreach ($c in (Extract-Category -Text $Text)) { $keys += "CAT:$($c.ToLower())" }
    return $keys
}

function Test-CombinedMatch {
    param([array]$Keys1, [array]$Keys2, [int]$MinMatches = 2)
    if ($Keys1.Count -eq 0 -or $Keys2.Count -eq 0) { return $false }
    $count = 0
    foreach ($k in $Keys1) { if ($Keys2 -contains $k) { $count++ } }
    return ($count -ge $MinMatches)
}

function Get-MatchKeyForEmail {
    param($Email, [string]$MatchType, [string]$CustomPattern)
    $text = "$($Email.Subject) $($Email.Body)"
    switch ($MatchType) {
        "Combined (All Fields)" { return Extract-CombinedKey -Text $text }
        "IP Address" { 
            $ips = Extract-IPAddresses -Text $text
            return ($ips | ForEach-Object { "IP:$($_.ToLower())" })
        }
        "Device Name" { 
            $devs = Extract-DeviceName -Text $text
            return ($devs | ForEach-Object { "DEV:$($_.ToLower())" })
        }
        "Serial Number" { 
            $sns = Extract-SerialNumber -Text $text
            return ($sns | ForEach-Object { "SN:$($_.ToLower())" })
        }
        "Category" { 
            $cats = Extract-Category -Text $text
            return ($cats | ForEach-Object { "CAT:$($_.ToLower())" })
        }
        "Custom Regex" {
            if ($CustomPattern) {
                try {
                    $ms = [regex]::Matches($text, $CustomPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    return ($ms | ForEach-Object { "CUSTOM:$($_.Value.ToLower())" })
                } catch { return @() }
            }
            return @()
        }
        default { return @() }
    }
}

function Get-MatchRuleForDomain {
    param([string]$Domain, [string]$Severity)
    foreach ($rule in $script:matchRules) {
        $domainMatch = ($rule.Domain -eq "-- ALL DOMAINS --") -or ($rule.Domain -eq $Domain)
        $sevMatch = ($rule.Severity -eq "ALL") -or ($rule.Severity -eq $Severity)
        if ($domainMatch -and $sevMatch) { return $rule }
    }
    return $null
}

function Test-EmailExcluded {
    param($Email)
    foreach ($f in $script:exclusionFilters) {
        $domMatch = ($f.Domain -eq "-- ALL --") -or ($f.Domain -eq $Email.Domain)
        if ($domMatch) {
            $content = "$($Email.Subject) $($Email.Sender)".ToLower()
            if ($content.Contains($f.Word.ToLower())) { return $true }
        }
    }
    return $false
}

# =================== OUTLOOK INIT ===================
try {
    $script:outlook = New-Object -ComObject Outlook.Application
    $script:namespace = $script:outlook.GetNamespace("MAPI")
} catch {
    [System.Windows.Forms.MessageBox]::Show("Cannot connect to Outlook!", "Error", "OK", "Error")
    exit
}

# =================== MAIN FORM ===================
$form = New-Object System.Windows.Forms.Form
$form.Text = "TICKET MONITOR PRO"
$form.BackColor = $theme.bgMain
$form.ForeColor = $theme.textPrimary
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1600, 950)
$form.MinimumSize = New-Object System.Drawing.Size(1200, 700)


# =================== SIDEBAR (Left Panel - 320px) ===================
$sidebar = New-Object System.Windows.Forms.Panel
$sidebar.Width = 320
$sidebar.Dock = "Left"
$sidebar.BackColor = $theme.bgSidebar
$sidebar.Padding = New-Object System.Windows.Forms.Padding(15, 15, 15, 15)
$form.Controls.Add($sidebar)

# Vertical separator
$sidebarLine = New-Object System.Windows.Forms.Panel
$sidebarLine.Width = 1
$sidebarLine.Dock = "Left"
$sidebarLine.BackColor = $theme.border
$form.Controls.Add($sidebarLine)

# --- LOGO ---
$logoBox = New-Object System.Windows.Forms.Panel
$logoBox.Dock = "Top"
$logoBox.Height = 70
$logoBox.BackColor = [System.Drawing.Color]::Transparent
$sidebar.Controls.Add($logoBox)

$lblLogo1 = New-Object System.Windows.Forms.Label
$lblLogo1.Text = "TICKET"
$lblLogo1.Font = New-Object System.Drawing.Font("Segoe UI", 22, [System.Drawing.FontStyle]::Bold)
$lblLogo1.ForeColor = $theme.accent
$lblLogo1.Location = New-Object System.Drawing.Point(0, 5)
$lblLogo1.AutoSize = $true
$logoBox.Controls.Add($lblLogo1)

$lblLogo2 = New-Object System.Windows.Forms.Label
$lblLogo2.Text = "MONITOR PRO"
$lblLogo2.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$lblLogo2.ForeColor = $theme.textSecond
$lblLogo2.Location = New-Object System.Drawing.Point(2, 42)
$lblLogo2.AutoSize = $true
$logoBox.Controls.Add($lblLogo2)

# --- PIE CHART ---
$pieBox = New-Object System.Windows.Forms.Panel
$pieBox.Dock = "Top"
$pieBox.Height = 180
$pieBox.BackColor = $theme.bgCard
$pieBox.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 10)
$sidebar.Controls.Add($pieBox)

$script:piePanel = New-Object DoubleBufferedPanel
$script:piePanel.Location = New-Object System.Drawing.Point(75, 10)
$script:piePanel.Size = New-Object System.Drawing.Size(140, 140)
$script:piePanel.BackColor = [System.Drawing.Color]::Transparent
$pieBox.Controls.Add($script:piePanel)

$script:piePanel.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
    $rect = New-Object System.Drawing.Rectangle(5, 5, 130, 130)
    $total = $script:pieCritical + $script:pieAttention + $script:pieTrouble
    
    # Background circle
    $bgB = New-Object System.Drawing.SolidBrush($theme.bgCardHover)
    $g.FillEllipse($bgB, $rect)
    $bgB.Dispose()
    
    if ($total -gt 0) {
        $start = -90
        if ($script:pieCritical -gt 0) {
            $sw = ($script:pieCritical / $total) * 360
            $br = New-Object System.Drawing.SolidBrush($theme.critical)
            $g.FillPie($br, $rect, $start, $sw); $br.Dispose(); $start += $sw
        }
        if ($script:pieAttention -gt 0) {
            $sw = ($script:pieAttention / $total) * 360
            $br = New-Object System.Drawing.SolidBrush($theme.attention)
            $g.FillPie($br, $rect, $start, $sw); $br.Dispose(); $start += $sw
        }
        if ($script:pieTrouble -gt 0) {
            $sw = ($script:pieTrouble / $total) * 360
            $br = New-Object System.Drawing.SolidBrush($theme.trouble)
            $g.FillPie($br, $rect, $start, $sw); $br.Dispose()
        }
    }
    
    # Inner circle (donut hole)
    $inner = New-Object System.Drawing.Rectangle(35, 35, 70, 70)
    $innerB = New-Object System.Drawing.SolidBrush($theme.bgCard)
    $g.FillEllipse($innerB, $inner); $innerB.Dispose()
    
    # Center text
    $txtB = New-Object System.Drawing.SolidBrush($theme.textPrimary)
    $fntBig = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $fntSm = New-Object System.Drawing.Font("Segoe UI", 7)
    $numStr = $total.ToString()
    $numSz = $g.MeasureString($numStr, $fntBig)
    $g.DrawString($numStr, $fntBig, $txtB, (140 - $numSz.Width)/2, 48)
    $mutedB = New-Object System.Drawing.SolidBrush($theme.textMuted)
    $g.DrawString("ACTIVE", $fntSm, $mutedB, 48, 78)
    $txtB.Dispose(); $mutedB.Dispose(); $fntBig.Dispose(); $fntSm.Dispose()
})

$lblPieTitle = New-Object System.Windows.Forms.Label
$lblPieTitle.Text = "ALERT OVERVIEW"
$lblPieTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblPieTitle.ForeColor = $theme.textMuted
$lblPieTitle.Location = New-Object System.Drawing.Point(95, 155)
$lblPieTitle.AutoSize = $true
$pieBox.Controls.Add($lblPieTitle)

# --- STAT CARDS ---
$statsBox = New-Object System.Windows.Forms.Panel
$statsBox.Dock = "Top"
$statsBox.Height = 155
$statsBox.BackColor = [System.Drawing.Color]::Transparent
$sidebar.Controls.Add($statsBox)

function New-StatRow {
    param($Parent, $Y, $Label, $Color, [ref]$OutLabel)
    
    $row = New-Object System.Windows.Forms.Panel
    $row.Location = New-Object System.Drawing.Point(0, $Y)
    $row.Size = New-Object System.Drawing.Size(288, 32)
    $row.BackColor = $theme.bgCard
    $Parent.Controls.Add($row)
    
    $bar = New-Object System.Windows.Forms.Panel
    $bar.Location = New-Object System.Drawing.Point(0, 0)
    $bar.Size = New-Object System.Drawing.Size(4, 32)
    $bar.BackColor = $Color
    $row.Controls.Add($bar)
    
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lbl.ForeColor = $theme.textSecond
    $lbl.Location = New-Object System.Drawing.Point(12, 7)
    $lbl.AutoSize = $true
    $row.Controls.Add($lbl)
    
    $val = New-Object System.Windows.Forms.Label
    $val.Text = "0"
    $val.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $val.ForeColor = $Color
    $val.TextAlign = "MiddleRight"
    $val.Location = New-Object System.Drawing.Point(230, 5)
    $val.Size = New-Object System.Drawing.Size(50, 22)
    $row.Controls.Add($val)
    
    $OutLabel.Value = $val
}

$script:lblValCritical = $null
$script:lblValAttention = $null
$script:lblValTrouble = $null
$script:lblValCleared = $null

New-StatRow -Parent $statsBox -Y 0 -Label "CRITICAL" -Color $theme.critical -OutLabel ([ref]$script:lblValCritical)
New-StatRow -Parent $statsBox -Y 38 -Label "ATTENTION" -Color $theme.attention -OutLabel ([ref]$script:lblValAttention)
New-StatRow -Parent $statsBox -Y 76 -Label "TROUBLE" -Color $theme.trouble -OutLabel ([ref]$script:lblValTrouble)
New-StatRow -Parent $statsBox -Y 114 -Label "CLEARED" -Color $theme.clear -OutLabel ([ref]$script:lblValCleared)

# --- SCAN SETTINGS ---
$settingsBox = New-Object System.Windows.Forms.Panel
$settingsBox.Dock = "Top"
$settingsBox.Height = 260
$settingsBox.BackColor = [System.Drawing.Color]::Transparent
$sidebar.Controls.Add($settingsBox)

$lblSettings = New-Object System.Windows.Forms.Label
$lblSettings.Text = "SCAN SETTINGS"
$lblSettings.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblSettings.ForeColor = $theme.textMuted
$lblSettings.Location = New-Object System.Drawing.Point(0, 10)
$lblSettings.AutoSize = $true
$settingsBox.Controls.Add($lblSettings)

# Profile
$lblProfile = New-Object System.Windows.Forms.Label
$lblProfile.Text = "Outlook Profile"
$lblProfile.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblProfile.ForeColor = $theme.textSecond
$lblProfile.Location = New-Object System.Drawing.Point(0, 35)
$lblProfile.AutoSize = $true
$settingsBox.Controls.Add($lblProfile)

$cmbProfile = New-Object System.Windows.Forms.ComboBox
$cmbProfile.Location = New-Object System.Drawing.Point(0, 52)
$cmbProfile.Size = New-Object System.Drawing.Size(288, 26)
$cmbProfile.DropDownStyle = "DropDownList"
$cmbProfile.BackColor = $theme.bgInput
$cmbProfile.ForeColor = $theme.textPrimary
$cmbProfile.FlatStyle = "Flat"
foreach ($store in $script:namespace.Stores) { [void]$cmbProfile.Items.Add($store.DisplayName) }
if ($cmbProfile.Items.Count -gt 0) { $cmbProfile.SelectedIndex = 0 }
$settingsBox.Controls.Add($cmbProfile)

# Date/Time
$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Text = "Scan From"
$lblFrom.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblFrom.ForeColor = $theme.textSecond
$lblFrom.Location = New-Object System.Drawing.Point(0, 85)
$lblFrom.AutoSize = $true
$settingsBox.Controls.Add($lblFrom)

$dtpDate = New-Object System.Windows.Forms.DateTimePicker
$dtpDate.Location = New-Object System.Drawing.Point(0, 102)
$dtpDate.Size = New-Object System.Drawing.Size(150, 26)
$dtpDate.Format = "Short"
$dtpDate.Value = (Get-Date).Date
$settingsBox.Controls.Add($dtpDate)

$nudHour = New-Object System.Windows.Forms.NumericUpDown
$nudHour.Location = New-Object System.Drawing.Point(160, 102)
$nudHour.Size = New-Object System.Drawing.Size(55, 26)
$nudHour.Minimum = 0; $nudHour.Maximum = 23; $nudHour.Value = 0
$nudHour.BackColor = $theme.bgInput; $nudHour.ForeColor = $theme.textPrimary
$settingsBox.Controls.Add($nudHour)

$lblColon = New-Object System.Windows.Forms.Label
$lblColon.Text = ":"
$lblColon.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$lblColon.ForeColor = $theme.textMuted
$lblColon.Location = New-Object System.Drawing.Point(218, 104)
$lblColon.AutoSize = $true
$settingsBox.Controls.Add($lblColon)

$nudMinute = New-Object System.Windows.Forms.NumericUpDown
$nudMinute.Location = New-Object System.Drawing.Point(232, 102)
$nudMinute.Size = New-Object System.Drawing.Size(55, 26)
$nudMinute.Minimum = 0; $nudMinute.Maximum = 59; $nudMinute.Value = 0
$nudMinute.BackColor = $theme.bgInput; $nudMinute.ForeColor = $theme.textPrimary
$settingsBox.Controls.Add($nudMinute)

# Interval
$lblInterval = New-Object System.Windows.Forms.Label
$lblInterval.Text = "Refresh (sec)"
$lblInterval.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblInterval.ForeColor = $theme.textSecond
$lblInterval.Location = New-Object System.Drawing.Point(0, 135)
$lblInterval.AutoSize = $true
$settingsBox.Controls.Add($lblInterval)

$nudInterval = New-Object System.Windows.Forms.NumericUpDown
$nudInterval.Location = New-Object System.Drawing.Point(0, 152)
$nudInterval.Size = New-Object System.Drawing.Size(70, 26)
$nudInterval.Minimum = 10; $nudInterval.Maximum = 300; $nudInterval.Value = 30
$nudInterval.BackColor = $theme.bgInput; $nudInterval.ForeColor = $theme.textPrimary
$settingsBox.Controls.Add($nudInterval)

# Checkboxes
$chkContinuous = New-Object System.Windows.Forms.CheckBox
$chkContinuous.Text = "Auto"
$chkContinuous.Location = New-Object System.Drawing.Point(85, 154)
$chkContinuous.ForeColor = $theme.textPrimary
$chkContinuous.AutoSize = $true
$chkContinuous.Checked = $true
$settingsBox.Controls.Add($chkContinuous)

$script:chkSound = New-Object System.Windows.Forms.CheckBox
$script:chkSound.Text = "Sound"
$script:chkSound.Location = New-Object System.Drawing.Point(145, 154)
$script:chkSound.ForeColor = $theme.attention
$script:chkSound.AutoSize = $true
$script:chkSound.Checked = $true
$settingsBox.Controls.Add($script:chkSound)

# Buttons
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "START SCAN"
$btnStart.Location = New-Object System.Drawing.Point(0, 190)
$btnStart.Size = New-Object System.Drawing.Size(140, 38)
$btnStart.BackColor = $theme.accent
$btnStart.ForeColor = [System.Drawing.Color]::FromArgb(10, 15, 25)
$btnStart.FlatStyle = "Flat"
$btnStart.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnStart.FlatAppearance.BorderSize = 0
$btnStart.Cursor = "Hand"
$settingsBox.Controls.Add($btnStart)

$btnClearAll = New-Object System.Windows.Forms.Button
$btnClearAll.Text = "CLEAR ALL"
$btnClearAll.Location = New-Object System.Drawing.Point(148, 190)
$btnClearAll.Size = New-Object System.Drawing.Size(140, 38)
$btnClearAll.BackColor = $theme.bgCard
$btnClearAll.ForeColor = $theme.critical
$btnClearAll.FlatStyle = "Flat"
$btnClearAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnClearAll.FlatAppearance.BorderColor = $theme.critical
$btnClearAll.FlatAppearance.BorderSize = 1
$btnClearAll.Cursor = "Hand"
$settingsBox.Controls.Add($btnClearAll)

# Status indicator
$statusRow = New-Object System.Windows.Forms.Panel
$statusRow.Location = New-Object System.Drawing.Point(0, 235)
$statusRow.Size = New-Object System.Drawing.Size(288, 24)
$statusRow.BackColor = $theme.bgCard
$settingsBox.Controls.Add($statusRow)

$script:lblStatusDot = New-Object System.Windows.Forms.Label
$script:lblStatusDot.Text = [char]0x25CF
$script:lblStatusDot.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$script:lblStatusDot.ForeColor = $theme.textMuted
$script:lblStatusDot.Location = New-Object System.Drawing.Point(8, 3)
$script:lblStatusDot.AutoSize = $true
$statusRow.Controls.Add($script:lblStatusDot)

$script:lblStatus = New-Object System.Windows.Forms.Label
$script:lblStatus.Text = "Ready"
$script:lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$script:lblStatus.ForeColor = $theme.textSecond
$script:lblStatus.Location = New-Object System.Drawing.Point(28, 4)
$script:lblStatus.AutoSize = $true
$statusRow.Controls.Add($script:lblStatus)

# --- THEME SELECTOR (bottom) ---
$themeBox = New-Object System.Windows.Forms.Panel
$themeBox.Dock = "Bottom"
$themeBox.Height = 55
$themeBox.BackColor = [System.Drawing.Color]::Transparent
$sidebar.Controls.Add($themeBox)

$lblTheme = New-Object System.Windows.Forms.Label
$lblTheme.Text = "Theme"
$lblTheme.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblTheme.ForeColor = $theme.textMuted
$lblTheme.Location = New-Object System.Drawing.Point(0, 5)
$lblTheme.AutoSize = $true
$themeBox.Controls.Add($lblTheme)

$cmbTheme = New-Object System.Windows.Forms.ComboBox
$cmbTheme.Location = New-Object System.Drawing.Point(0, 22)
$cmbTheme.Size = New-Object System.Drawing.Size(288, 26)
$cmbTheme.DropDownStyle = "DropDownList"
$cmbTheme.BackColor = $theme.bgInput
$cmbTheme.ForeColor = $theme.textPrimary
$cmbTheme.FlatStyle = "Flat"
foreach ($t in $script:themes.Keys | Sort-Object) { [void]$cmbTheme.Items.Add($t) }
$cmbTheme.SelectedItem = "Cyber Night"
$themeBox.Controls.Add($cmbTheme)


# =================== MAIN CONTENT AREA ===================
$mainArea = New-Object System.Windows.Forms.Panel
$mainArea.Dock = "Fill"
$mainArea.BackColor = $theme.bgMain
$mainArea.Padding = New-Object System.Windows.Forms.Padding(20, 15, 20, 15)
$form.Controls.Add($mainArea)

# --- TOP TOOLBAR (Filters & Rules) ---
$toolbar = New-Object System.Windows.Forms.Panel
$toolbar.Dock = "Top"
$toolbar.Height = 130
$toolbar.BackColor = $theme.bgCard
$toolbar.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)
$mainArea.Controls.Add($toolbar)

# Match Rules Row
$lblMatchRules = New-Object System.Windows.Forms.Label
$lblMatchRules.Text = "MATCH RULES"
$lblMatchRules.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblMatchRules.ForeColor = $theme.accent
$lblMatchRules.Location = New-Object System.Drawing.Point(15, 8)
$lblMatchRules.AutoSize = $true
$toolbar.Controls.Add($lblMatchRules)

$cmbMatchDomain = New-Object System.Windows.Forms.ComboBox
$cmbMatchDomain.Location = New-Object System.Drawing.Point(15, 28)
$cmbMatchDomain.Size = New-Object System.Drawing.Size(160, 24)
$cmbMatchDomain.DropDownStyle = "DropDownList"
$cmbMatchDomain.BackColor = $theme.bgInput
$cmbMatchDomain.ForeColor = $theme.textPrimary
$cmbMatchDomain.FlatStyle = "Flat"
[void]$cmbMatchDomain.Items.Add("-- ALL DOMAINS --")
$cmbMatchDomain.SelectedIndex = 0
$toolbar.Controls.Add($cmbMatchDomain)

$cmbMatchType = New-Object System.Windows.Forms.ComboBox
$cmbMatchType.Location = New-Object System.Drawing.Point(185, 28)
$cmbMatchType.Size = New-Object System.Drawing.Size(170, 24)
$cmbMatchType.DropDownStyle = "DropDownList"
$cmbMatchType.BackColor = $theme.bgInput
$cmbMatchType.ForeColor = $theme.textPrimary
$cmbMatchType.FlatStyle = "Flat"
foreach ($mt in $script:matchFieldTypes) { [void]$cmbMatchType.Items.Add($mt) }
$cmbMatchType.SelectedIndex = 0
$toolbar.Controls.Add($cmbMatchType)

$cmbMatchSev = New-Object System.Windows.Forms.ComboBox
$cmbMatchSev.Location = New-Object System.Drawing.Point(365, 28)
$cmbMatchSev.Size = New-Object System.Drawing.Size(90, 24)
$cmbMatchSev.DropDownStyle = "DropDownList"
$cmbMatchSev.BackColor = $theme.bgInput
$cmbMatchSev.ForeColor = $theme.textPrimary
$cmbMatchSev.FlatStyle = "Flat"
[void]$cmbMatchSev.Items.AddRange(@("ALL", "CRITICAL", "ATTENTION", "TROUBLE"))
$cmbMatchSev.SelectedIndex = 0
$toolbar.Controls.Add($cmbMatchSev)

$btnAddRule = New-Object System.Windows.Forms.Button
$btnAddRule.Text = "ADD RULE"
$btnAddRule.Location = New-Object System.Drawing.Point(465, 26)
$btnAddRule.Size = New-Object System.Drawing.Size(90, 28)
$btnAddRule.BackColor = $theme.accent
$btnAddRule.ForeColor = $theme.bgDeep
$btnAddRule.FlatStyle = "Flat"
$btnAddRule.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnAddRule.FlatAppearance.BorderSize = 0
$toolbar.Controls.Add($btnAddRule)

# Active Rules Display
$lblActiveRules = New-Object System.Windows.Forms.Label
$lblActiveRules.Text = "Active:"
$lblActiveRules.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblActiveRules.ForeColor = $theme.textMuted
$lblActiveRules.Location = New-Object System.Drawing.Point(570, 30)
$lblActiveRules.AutoSize = $true
$toolbar.Controls.Add($lblActiveRules)

$listRules = New-Object System.Windows.Forms.ListBox
$listRules.Location = New-Object System.Drawing.Point(620, 8)
$listRules.Size = New-Object System.Drawing.Size(400, 45)
$listRules.BackColor = $theme.bgInput
$listRules.ForeColor = $theme.textPrimary
$listRules.BorderStyle = "None"
$listRules.Font = New-Object System.Drawing.Font("Consolas", 8)
$toolbar.Controls.Add($listRules)

# Exclusion Row
$lblExclusion = New-Object System.Windows.Forms.Label
$lblExclusion.Text = "EXCLUSION FILTER"
$lblExclusion.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblExclusion.ForeColor = $theme.critical
$lblExclusion.Location = New-Object System.Drawing.Point(15, 62)
$lblExclusion.AutoSize = $true
$toolbar.Controls.Add($lblExclusion)

$cmbExcludeDomain = New-Object System.Windows.Forms.ComboBox
$cmbExcludeDomain.Location = New-Object System.Drawing.Point(15, 82)
$cmbExcludeDomain.Size = New-Object System.Drawing.Size(160, 24)
$cmbExcludeDomain.DropDownStyle = "DropDownList"
$cmbExcludeDomain.BackColor = $theme.bgInput
$cmbExcludeDomain.ForeColor = $theme.textPrimary
$cmbExcludeDomain.FlatStyle = "Flat"
[void]$cmbExcludeDomain.Items.Add("-- ALL --")
$cmbExcludeDomain.SelectedIndex = 0
$toolbar.Controls.Add($cmbExcludeDomain)

$txtExcludeWord = New-Object System.Windows.Forms.TextBox
$txtExcludeWord.Location = New-Object System.Drawing.Point(185, 82)
$txtExcludeWord.Size = New-Object System.Drawing.Size(200, 24)
$txtExcludeWord.BackColor = $theme.bgInput
$txtExcludeWord.ForeColor = $theme.textPrimary
$txtExcludeWord.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtExcludeWord.BorderStyle = "FixedSingle"
$toolbar.Controls.Add($txtExcludeWord)

$btnAddExclusion = New-Object System.Windows.Forms.Button
$btnAddExclusion.Text = "EXCLUDE"
$btnAddExclusion.Location = New-Object System.Drawing.Point(395, 80)
$btnAddExclusion.Size = New-Object System.Drawing.Size(90, 28)
$btnAddExclusion.BackColor = $theme.critical
$btnAddExclusion.ForeColor = [System.Drawing.Color]::White
$btnAddExclusion.FlatStyle = "Flat"
$btnAddExclusion.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnAddExclusion.FlatAppearance.BorderSize = 0
$toolbar.Controls.Add($btnAddExclusion)

$listExclusions = New-Object System.Windows.Forms.ListBox
$listExclusions.Location = New-Object System.Drawing.Point(500, 62)
$listExclusions.Size = New-Object System.Drawing.Size(300, 45)
$listExclusions.BackColor = $theme.bgInput
$listExclusions.ForeColor = $theme.textPrimary
$listExclusions.BorderStyle = "None"
$listExclusions.Font = New-Object System.Drawing.Font("Consolas", 8)
$toolbar.Controls.Add($listExclusions)

# Severity filters
$lblSevFilter = New-Object System.Windows.Forms.Label
$lblSevFilter.Text = "Show:"
$lblSevFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblSevFilter.ForeColor = $theme.textMuted
$lblSevFilter.Location = New-Object System.Drawing.Point(830, 62)
$lblSevFilter.AutoSize = $true
$toolbar.Controls.Add($lblSevFilter)

$chkCritical = New-Object System.Windows.Forms.CheckBox
$chkCritical.Text = "CRIT"
$chkCritical.Location = New-Object System.Drawing.Point(830, 82)
$chkCritical.ForeColor = $theme.critical
$chkCritical.AutoSize = $true
$chkCritical.Checked = $true
$chkCritical.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$toolbar.Controls.Add($chkCritical)

$chkAttention = New-Object System.Windows.Forms.CheckBox
$chkAttention.Text = "ATT"
$chkAttention.Location = New-Object System.Drawing.Point(890, 82)
$chkAttention.ForeColor = $theme.attention
$chkAttention.AutoSize = $true
$chkAttention.Checked = $true
$chkAttention.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$toolbar.Controls.Add($chkAttention)

$chkTrouble = New-Object System.Windows.Forms.CheckBox
$chkTrouble.Text = "TRB"
$chkTrouble.Location = New-Object System.Drawing.Point(945, 82)
$chkTrouble.ForeColor = $theme.trouble
$chkTrouble.AutoSize = $true
$chkTrouble.Checked = $true
$chkTrouble.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$toolbar.Controls.Add($chkTrouble)

$chkClear = New-Object System.Windows.Forms.CheckBox
$chkClear.Text = "CLR"
$chkClear.Location = New-Object System.Drawing.Point(1000, 82)
$chkClear.ForeColor = $theme.clear
$chkClear.AutoSize = $true
$chkClear.Checked = $true
$chkClear.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$toolbar.Controls.Add($chkClear)

# --- SEPARATOR ---
$sep1 = New-Object System.Windows.Forms.Panel
$sep1.Dock = "Top"
$sep1.Height = 10
$sep1.BackColor = $theme.bgMain
$mainArea.Controls.Add($sep1)

# --- DATA GRID ---
$grid = New-Object DoubleBufferedDGV
$grid.Dock = "Fill"
$grid.BackgroundColor = $theme.bgDeep
$grid.BorderStyle = "None"
$grid.CellBorderStyle = "SingleHorizontal"
$grid.ColumnHeadersBorderStyle = "None"
$grid.EnableHeadersVisualStyles = $false
$grid.GridColor = $theme.border
$grid.RowHeadersVisible = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "Fill"
$grid.RowTemplate.Height = 36

# Header style
$grid.ColumnHeadersDefaultCellStyle.BackColor = $theme.bgCard
$grid.ColumnHeadersDefaultCellStyle.ForeColor = $theme.textPrimary
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$grid.ColumnHeadersHeight = 42
$grid.ColumnHeadersHeightSizeMode = "DisableResizing"

# Cell style
$grid.DefaultCellStyle.BackColor = $theme.bgDeep
$grid.DefaultCellStyle.ForeColor = $theme.textPrimary
$grid.DefaultCellStyle.SelectionBackColor = $theme.bgCardHover
$grid.DefaultCellStyle.SelectionForeColor = $theme.textPrimary
$grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)

# Alternate row
$grid.AlternatingRowsDefaultCellStyle.BackColor = $theme.gridAlt

# Columns
$colSev = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSev.Name = "Severity"; $colSev.HeaderText = "SEVERITY"; $colSev.Width = 130; $colSev.MinimumWidth = 100
[void]$grid.Columns.Add($colSev)

$colDomain = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDomain.Name = "Domain"; $colDomain.HeaderText = "DOMAIN"; $colDomain.Width = 140; $colDomain.MinimumWidth = 100
[void]$grid.Columns.Add($colDomain)

$colSender = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSender.Name = "Sender"; $colSender.HeaderText = "SENDER"; $colSender.Width = 180; $colSender.MinimumWidth = 120
[void]$grid.Columns.Add($colSender)

$colSubject = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSubject.Name = "Subject"; $colSubject.HeaderText = "SUBJECT"; $colSubject.Width = 280; $colSubject.MinimumWidth = 150
[void]$grid.Columns.Add($colSubject)

$colTime = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colTime.Name = "Time"; $colTime.HeaderText = "TIME"; $colTime.Width = 145; $colTime.MinimumWidth = 120
[void]$grid.Columns.Add($colTime)

$colMatch = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colMatch.Name = "MatchKey"; $colMatch.HeaderText = "MATCH KEY"; $colMatch.Width = 160; $colMatch.MinimumWidth = 100
[void]$grid.Columns.Add($colMatch)

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.Name = "Status"; $colStatus.HeaderText = "STATUS"; $colStatus.Width = 90; $colStatus.MinimumWidth = 80
[void]$grid.Columns.Add($colStatus)

$mainArea.Controls.Add($grid)


# =================== ROW COLORING ===================
function Set-RowColor {
    param($Row, [string]$Severity, [bool]$IsCleared = $false)
    
    if ($IsCleared) {
        $Row.DefaultCellStyle.ForeColor = $theme.clear
        $Row.Cells["Status"].Style.ForeColor = $theme.clear
        $Row.Cells["Severity"].Style.ForeColor = $theme.clear
    } else {
        switch -Wildcard ($Severity) {
            "CRITICAL*" { 
                $Row.DefaultCellStyle.ForeColor = $theme.critical
                $Row.Cells["Severity"].Style.ForeColor = $theme.critical
                $Row.Cells["Severity"].Style.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            }
            "ATTENTION*" { 
                $Row.DefaultCellStyle.ForeColor = $theme.attention
                $Row.Cells["Severity"].Style.ForeColor = $theme.attention
            }
            "TROUBLE*" { 
                $Row.DefaultCellStyle.ForeColor = $theme.trouble
                $Row.Cells["Severity"].Style.ForeColor = $theme.trouble
            }
            "CLEAR*" { 
                $Row.DefaultCellStyle.ForeColor = $theme.clear
                $Row.Cells["Severity"].Style.ForeColor = $theme.clear
            }
        }
    }
}

# =================== UPDATE STATS ===================
function Update-Stats {
    $crit = 0; $att = 0; $trb = 0; $clr = 0
    
    foreach ($email in $script:allEmails) {
        if (Test-EmailExcluded -Email $email) { continue }
        if ($email.Sev -eq "CLEAR" -and $email.MatchedClearId) { continue }
        if ($email.IsEscalated) { continue }
        
        if ($email.IsCleared) { $clr++ }
        elseif ($email.Sev -eq "CRITICAL") { $crit++ }
        elseif ($email.Sev -eq "ATTENTION") { $att++ }
        elseif ($email.Sev -eq "TROUBLE") { $trb++ }
    }
    
    $script:lblValCritical.Text = $crit.ToString()
    $script:lblValAttention.Text = $att.ToString()
    $script:lblValTrouble.Text = $trb.ToString()
    $script:lblValCleared.Text = $clr.ToString()
    
    # Update pie chart
    $script:pieCritical = $crit
    $script:pieAttention = $att
    $script:pieTrouble = $trb
    $script:pieCleared = $clr
    $script:piePanel.Refresh()
}

# =================== UPDATE RULES LIST ===================
function Update-RulesList {
    $listRules.Items.Clear()
    if ($script:matchRules.Count -eq 0) {
        $listRules.Items.Add("(No rules - emails won't be matched)")
    } else {
        foreach ($r in $script:matchRules) {
            $listRules.Items.Add("[$($r.Domain)] [$($r.Severity)] -> $($r.MatchType)")
        }
    }
}

function Update-ExclusionsList {
    $listExclusions.Items.Clear()
    foreach ($f in $script:exclusionFilters) {
        $listExclusions.Items.Add("[$($f.Domain)] -> '$($f.Word)'")
    }
}

# =================== APPLY FILTERS ===================
function Apply-Filters {
    foreach ($row in $grid.Rows) {
        if ($row.IsNewRow) { continue }
        
        $sev = $row.Cells["Severity"].Value
        $show = $true
        
        $baseSev = $sev -replace " -> CLEAR", ""
        
        if ($baseSev -eq "CRITICAL" -and -not $chkCritical.Checked) { $show = $false }
        elseif ($baseSev -eq "ATTENTION" -and -not $chkAttention.Checked) { $show = $false }
        elseif ($baseSev -eq "TROUBLE" -and -not $chkTrouble.Checked) { $show = $false }
        elseif ($baseSev -eq "CLEAR" -and -not $chkClear.Checked) { $show = $false }
        
        $row.Visible = $show
    }
}

# =================== GET ALL MAIL FOLDERS ===================
function Get-AllMailFolders {
    param([object]$Folder, [int]$Depth = 0)
    
    $list = [System.Collections.ArrayList]::new()
    [void]$list.Add($Folder)
    
    foreach ($sub in $Folder.Folders) {
        try {
            $subList = Get-AllMailFolders -Folder $sub -Depth ($Depth + 1)
            foreach ($f in $subList) { [void]$list.Add($f) }
        } catch { continue }
    }
    return $list
}

# =================== SCAN MAILS FUNCTION ===================
function Invoke-ScanMails {
    param([bool]$IsInitial = $false)
    
    $newCount = 0
    $script:folderScanLog.Clear()
    
    try {
        $storeName = $cmbProfile.SelectedItem
        $store = $script:namespace.Stores | Where-Object { $_.DisplayName -eq $storeName }
        if (-not $store) { return 0 }
        
        $rootFolder = $store.GetRootFolder()
        $allFolders = Get-AllMailFolders -Folder $rootFolder
        
        $fromDate = $dtpDate.Value.Date.AddHours($nudHour.Value).AddMinutes($nudMinute.Value)
        
        [void]$script:folderScanLog.Add("Scanning from: $($fromDate.ToString('dd/MM/yyyy HH:mm'))")
        [void]$script:folderScanLog.Add("Folders found: $($allFolders.Count)")
        
        $newEmails = [System.Collections.ArrayList]::new()
        
        foreach ($folder in $allFolders) {
            if (-not $script:isRunning) { break }
            
            try {
                $folderPath = $folder.FolderPath
                $items = $folder.Items
                $items.Sort("[ReceivedTime]", $true)
                
                foreach ($mail in $items) {
                    try {
                        if ($mail.ReceivedTime -lt $fromDate) { continue }
                        
                        $entryId = $mail.EntryID
                        if ($script:processedIds.ContainsKey($entryId)) { continue }
                        
                        $subject = $mail.Subject
                        if (-not $subject) { continue }
                        
                        $sev = $null
                        if ($subject -match "CRITICAL") { $sev = "CRITICAL" }
                        elseif ($subject -match "ATTENTION") { $sev = "ATTENTION" }
                        elseif ($subject -match "TROUBLE") { $sev = "TROUBLE" }
                        elseif ($subject -match "CLEAR") { $sev = "CLEAR" }
                        
                        if (-not $sev) { continue }
                        
                        $script:processedIds[$entryId] = $true
                        
                        $sender = try { $mail.SenderEmailAddress } catch { "" }
                        $senderName = try { $mail.SenderName } catch { "" }
                        if (-not $sender) { $sender = $senderName }
                        
                        $body = try { $mail.Body } catch { "" }
                        $recTime = $mail.ReceivedTime
                        $domain = Get-EmailDomain -Email $sender
                        
                        $extractedIPs = Extract-IPAddresses -Text "$subject $body"
                        $ipDisplay = if ($extractedIPs.Count -gt 0) { $extractedIPs -join ", " } else { "-" }
                        
                        [void]$newEmails.Add(@{
                            Id = $entryId
                            Sev = $sev
                            IsCleared = $false
                            IsEscalated = $false
                            Domain = $domain
                            Sender = $sender
                            Subject = $subject
                            Time = $recTime
                            Mail = $mail
                            Body = $body
                            MatchKeys = @()
                            MatchKeyDisplay = $ipDisplay
                            MatchedClearId = $null
                        })
                        
                    } catch { continue }
                }
            } catch { continue }
        }
        
        if ($newEmails.Count -gt 0) {
            foreach ($email in $newEmails) {
                if (-not $script:allDomains.ContainsKey($email.Domain)) {
                    $script:allDomains[$email.Domain] = $true
                    [void]$cmbExcludeDomain.Items.Add($email.Domain)
                    [void]$cmbMatchDomain.Items.Add($email.Domain)
                }
                
                $rule = Get-MatchRuleForDomain -Domain $email.Domain -Severity $email.Sev
                if ($rule) {
                    $email.MatchKeys = Get-MatchKeyForEmail -Email $email -MatchType $rule.MatchType -CustomPattern $rule.CustomPattern
                    if ($email.MatchKeys.Count -gt 0) {
                        $email.MatchKeyDisplay = $email.MatchKeys -join ", "
                    }
                }
                
                [void]$script:allEmails.Add($email)
            }
            
            # Process matching and escalation
            Process-AlertMatching
            
            # Repopulate grid
            Refresh-Grid
            
            $newCount = $newEmails.Count
        }
        
        Update-Stats
        Apply-Filters
        
    } catch { }
    
    return $newCount
}

# =================== PROCESS ALERT MATCHING ===================
function Process-AlertMatching {
    # Reset states
    foreach ($email in $script:allEmails) {
        $email.IsCleared = $false
        $email.IsEscalated = $false
        $email.MatchedClearId = $null
    }
    
    $clearEmails = @{}
    $alertEmails = @{}
    $combinedClears = @()
    $combinedAlerts = @()
    
    foreach ($email in $script:allEmails) {
        $rule = Get-MatchRuleForDomain -Domain $email.Domain -Severity $email.Sev
        $isCombined = ($rule -and $rule.MatchType -eq "Combined (All Fields)")
        
        if ($isCombined) {
            if ($email.Sev -eq "CLEAR") { $combinedClears += $email }
            elseif ($email.Sev -in @("CRITICAL", "ATTENTION", "TROUBLE")) { $combinedAlerts += $email }
        } else {
            foreach ($key in $email.MatchKeys) {
                if ($email.Sev -eq "CLEAR") {
                    if (-not $clearEmails.ContainsKey($key)) { $clearEmails[$key] = @() }
                    $clearEmails[$key] += $email
                } elseif ($email.Sev -in @("CRITICAL", "ATTENTION", "TROUBLE")) {
                    if (-not $alertEmails.ContainsKey($key)) { $alertEmails[$key] = @() }
                    $alertEmails[$key] += $email
                }
            }
        }
    }
    
    # Escalation: ATTENTION -> TROUBLE -> CRITICAL
    $sevRank = @{ "ATTENTION" = 1; "TROUBLE" = 2; "CRITICAL" = 3 }
    $alertsByDevice = @{}
    
    foreach ($email in $script:allEmails) {
        if ($email.Sev -notin @("CRITICAL", "ATTENTION", "TROUBLE")) { continue }
        if ($email.MatchKeys.Count -eq 0) { continue }
        
        $devKey = ($email.MatchKeys | Sort-Object) -join "|"
        if (-not $alertsByDevice.ContainsKey($devKey)) { $alertsByDevice[$devKey] = @() }
        $alertsByDevice[$devKey] += $email
    }
    
    foreach ($devKey in $alertsByDevice.Keys) {
        $alerts = $alertsByDevice[$devKey]
        if ($alerts.Count -le 1) { continue }
        
        $maxRank = 0
        $maxEmail = $null
        foreach ($a in $alerts) {
            $r = $sevRank[$a.Sev]
            if ($r -gt $maxRank) { $maxRank = $r; $maxEmail = $a }
        }
        
        foreach ($a in $alerts) {
            if ($a.Id -ne $maxEmail.Id -and $sevRank[$a.Sev] -lt $maxRank) {
                $a.IsEscalated = $true
            }
        }
    }
    
    $matchedAlertIds = @{}
    $matchedClearIds = @{}
    
    # Combined matching (min 2 params)
    foreach ($alert in $combinedAlerts) {
        if ($matchedAlertIds.ContainsKey($alert.Id)) { continue }
        foreach ($clear in $combinedClears) {
            if ($matchedClearIds.ContainsKey($clear.Id)) { continue }
            if (Test-CombinedMatch -Keys1 $alert.MatchKeys -Keys2 $clear.MatchKeys -MinMatches 2) {
                $matchedAlertIds[$alert.Id] = $true
                $matchedClearIds[$clear.Id] = $true
                $alert.IsCleared = $true
                $alert.MatchedClearId = $clear.Id
                $clear.MatchedClearId = "used"
                break
            }
        }
    }
    
    # Standard key matching
    foreach ($key in $alertEmails.Keys) {
        if ($clearEmails.ContainsKey($key)) {
            $alerts = $alertEmails[$key] | Sort-Object { $_.Time }
            $clears = $clearEmails[$key] | Sort-Object { $_.Time }
            
            $bestClear = $null
            foreach ($c in $clears) {
                if (-not $matchedClearIds.ContainsKey($c.Id)) {
                    $bestClear = $c
                    $matchedClearIds[$bestClear.Id] = $true
                    break
                }
            }
            
            if ($bestClear) {
                foreach ($a in $alerts) {
                    if (-not $matchedAlertIds.ContainsKey($a.Id)) {
                        $matchedAlertIds[$a.Id] = $true
                        $a.IsCleared = $true
                        $a.MatchedClearId = $bestClear.Id
                    }
                }
                $bestClear.MatchedClearId = "used"
            }
        }
    }
}

# =================== REFRESH GRID ===================
function Refresh-Grid {
    $grid.SuspendLayout()
    $grid.Rows.Clear()
    $script:mailMap.Clear()
    
    $sorted = $script:allEmails | Sort-Object { $_.Time } -Descending
    
    foreach ($email in $sorted) {
        # Hide all standalone CLEARs
        if ($email.Sev -eq "CLEAR") { continue }
        # Hide escalated
        if ($email.IsEscalated) { continue }
        
        $dispSev = $email.Sev
        $dispStatus = "ACTIVE"
        
        if ($email.IsCleared) {
            $dispSev = "$($email.Sev) -> CLEAR"
            $dispStatus = "CLEARED"
        }
        
        $grid.Rows.Add(@(
            $dispSev,
            $email.Domain,
            $email.Sender,
            $email.Subject,
            $email.Time.ToString("dd/MM/yyyy HH:mm:ss"),
            $email.MatchKeyDisplay,
            $dispStatus
        ))
        
        $rowIdx = $grid.Rows.Count - 1
        $script:mailMap[$rowIdx] = $email.Mail
        Set-RowColor -Row $grid.Rows[$rowIdx] -Severity $email.Sev -IsCleared $email.IsCleared
    }
    
    $grid.ResumeLayout()
    $grid.Refresh()
}


# =================== THEME CHANGE ===================
function Apply-Theme {
    param([string]$ThemeName)
    
    $script:currentTheme = $ThemeName
    $script:theme = $script:themes[$ThemeName]
    $global:theme = $script:theme
    
    # Form
    $form.BackColor = $script:theme.bgMain
    
    # Sidebar
    $sidebar.BackColor = $script:theme.bgSidebar
    $sidebarLine.BackColor = $script:theme.border
    
    # Logo
    $lblLogo1.ForeColor = $script:theme.accent
    $lblLogo2.ForeColor = $script:theme.textSecond
    
    # Pie
    $pieBox.BackColor = $script:theme.bgCard
    $lblPieTitle.ForeColor = $script:theme.textMuted
    
    # Stats
    foreach ($ctrl in $statsBox.Controls) {
        if ($ctrl -is [System.Windows.Forms.Panel]) {
            $ctrl.BackColor = $script:theme.bgCard
        }
    }
    
    # Settings
    $lblSettings.ForeColor = $script:theme.textMuted
    $lblProfile.ForeColor = $script:theme.textSecond
    $lblFrom.ForeColor = $script:theme.textSecond
    $lblInterval.ForeColor = $script:theme.textSecond
    $cmbProfile.BackColor = $script:theme.bgInput
    $cmbProfile.ForeColor = $script:theme.textPrimary
    $nudInterval.BackColor = $script:theme.bgInput
    $nudInterval.ForeColor = $script:theme.textPrimary
    $chkContinuous.ForeColor = $script:theme.textPrimary
    $script:chkSound.ForeColor = $script:theme.attention
    $btnStart.BackColor = $script:theme.accent
    $btnClearAll.BackColor = $script:theme.bgCard
    $statusRow.BackColor = $script:theme.bgCard
    $script:lblStatusDot.ForeColor = $script:theme.textMuted
    $script:lblStatus.ForeColor = $script:theme.textSecond
    
    # Theme selector
    $lblTheme.ForeColor = $script:theme.textMuted
    $cmbTheme.BackColor = $script:theme.bgInput
    $cmbTheme.ForeColor = $script:theme.textPrimary
    
    # Main area
    $mainArea.BackColor = $script:theme.bgMain
    $toolbar.BackColor = $script:theme.bgCard
    
    # Toolbar controls
    $lblMatchRules.ForeColor = $script:theme.accent
    $lblExclusion.ForeColor = $script:theme.critical
    $cmbMatchDomain.BackColor = $script:theme.bgInput
    $cmbMatchDomain.ForeColor = $script:theme.textPrimary
    $cmbMatchType.BackColor = $script:theme.bgInput
    $cmbMatchType.ForeColor = $script:theme.textPrimary
    $cmbMatchSev.BackColor = $script:theme.bgInput
    $cmbMatchSev.ForeColor = $script:theme.textPrimary
    $btnAddRule.BackColor = $script:theme.accent
    $cmbExcludeDomain.BackColor = $script:theme.bgInput
    $cmbExcludeDomain.ForeColor = $script:theme.textPrimary
    $txtExcludeWord.BackColor = $script:theme.bgInput
    $txtExcludeWord.ForeColor = $script:theme.textPrimary
    $listRules.BackColor = $script:theme.bgInput
    $listRules.ForeColor = $script:theme.textPrimary
    $listExclusions.BackColor = $script:theme.bgInput
    $listExclusions.ForeColor = $script:theme.textPrimary
    
    # Grid
    $grid.BackgroundColor = $script:theme.bgDeep
    $grid.GridColor = $script:theme.border
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $script:theme.bgCard
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:theme.textPrimary
    $grid.DefaultCellStyle.BackColor = $script:theme.bgDeep
    $grid.DefaultCellStyle.ForeColor = $script:theme.textPrimary
    $grid.DefaultCellStyle.SelectionBackColor = $script:theme.bgCardHover
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $script:theme.gridAlt
    
    # Refresh
    $script:piePanel.Refresh()
    $form.Refresh()
}

$cmbTheme.Add_SelectedIndexChanged({
    Apply-Theme -ThemeName $cmbTheme.SelectedItem
})

# =================== TIMER ===================
$script:timer = New-Object System.Windows.Forms.Timer
$script:timer.Interval = 30000

$script:timer.Add_Tick({
    try {
        if (-not $script:isRunning -or -not $chkContinuous.Checked) { return }
        
        $script:lblStatus.Text = "Checking..."
        $script:lblStatusDot.ForeColor = $theme.attention
        [System.Windows.Forms.Application]::DoEvents()
        
        $prevCrit = $script:pieCritical
        $prevAtt = $script:pieAttention
        $prevTrb = $script:pieTrouble
        
        $new = Invoke-ScanMails -IsInitial $false
        
        if ($new -gt 0) {
            $script:lblStatus.Text = "+$new @ $(Get-Date -Format 'HH:mm:ss')"
            
            if ($script:chkSound.Checked) {
                if ($script:pieCritical -gt $prevCrit) { Play-SeveritySound -Severity "CRITICAL" }
                elseif ($script:pieAttention -gt $prevAtt) { Play-SeveritySound -Severity "ATTENTION" }
                elseif ($script:pieTrouble -gt $prevTrb) { Play-SeveritySound -Severity "TROUBLE" }
            }
        } else {
            $script:lblStatus.Text = "OK @ $(Get-Date -Format 'HH:mm:ss')"
        }
        $script:lblStatusDot.ForeColor = $theme.clear
    } catch {
        $script:lblStatus.Text = "Error"
    }
})

# =================== BUTTON EVENTS ===================
$btnStart.Add_Click({
    if (-not $cmbProfile.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Select a profile first!", "Warning", "OK", "Warning")
        return
    }
    
    if ($script:isRunning) {
        $script:isRunning = $false
        $script:timer.Stop()
        $btnStart.Text = "START SCAN"
        $btnStart.BackColor = $theme.accent
        $script:lblStatus.Text = "Stopped"
        $script:lblStatusDot.ForeColor = $theme.textMuted
    } else {
        $script:isRunning = $true
        $btnStart.Text = "STOP"
        $btnStart.BackColor = $theme.critical
        $script:lblStatus.Text = "Scanning..."
        $script:lblStatusDot.ForeColor = $theme.attention
        [System.Windows.Forms.Application]::DoEvents()
        
        Invoke-ScanMails -IsInitial $true
        
        $script:timer.Interval = [int]$nudInterval.Value * 1000
        if ($chkContinuous.Checked) { $script:timer.Start() }
        
        $script:lblStatus.Text = "Running"
        $script:lblStatusDot.ForeColor = $theme.clear
    }
})

$btnClearAll.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Clear all emails from view and free memory?`n`nThis does NOT delete emails from Outlook.",
        "Confirm Clear",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq "Yes") {
        $script:isRunning = $false
        $script:timer.Stop()
        
        $grid.Rows.Clear()
        $script:allEmails.Clear()
        $script:mailMap.Clear()
        $script:processedIds.Clear()
        
        $script:pieCritical = 0
        $script:pieAttention = 0
        $script:pieTrouble = 0
        $script:pieCleared = 0
        
        Update-Stats
        
        [System.GC]::Collect()
        
        $btnStart.Text = "START SCAN"
        $btnStart.BackColor = $theme.accent
        $script:lblStatus.Text = "Cleared"
        $script:lblStatusDot.ForeColor = $theme.textMuted
    }
})

$btnAddRule.Add_Click({
    $domain = $cmbMatchDomain.SelectedItem
    $matchType = $cmbMatchType.SelectedItem
    $severity = $cmbMatchSev.SelectedItem
    
    if (-not $domain -or -not $matchType -or -not $severity) { return }
    
    [void]$script:matchRules.Add(@{
        Domain = $domain
        MatchType = $matchType
        Severity = $severity
        CustomPattern = ""
    })
    
    Update-RulesList
    Process-AlertMatching
    Refresh-Grid
    Update-Stats
})

$btnAddExclusion.Add_Click({
    $domain = $cmbExcludeDomain.SelectedItem
    $word = $txtExcludeWord.Text.Trim()
    
    if (-not $domain -or -not $word) { return }
    
    [void]$script:exclusionFilters.Add(@{
        Domain = $domain
        Word = $word
    })
    
    $txtExcludeWord.Text = ""
    Update-ExclusionsList
    Apply-Filters
    Update-Stats
})

$listRules.Add_Click({
    $idx = $listRules.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $script:matchRules.Count) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Remove this rule?",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($result -eq "Yes") {
            $script:matchRules.RemoveAt($idx)
            Update-RulesList
            Process-AlertMatching
            Refresh-Grid
            Update-Stats
        }
    }
})

$listExclusions.Add_Click({
    $idx = $listExclusions.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $script:exclusionFilters.Count) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Remove this exclusion?",
            "Confirm",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($result -eq "Yes") {
            $script:exclusionFilters.RemoveAt($idx)
            Update-ExclusionsList
            Apply-Filters
            Update-Stats
        }
    }
})

# Severity filter checkboxes
$chkCritical.Add_CheckedChanged({ Apply-Filters })
$chkAttention.Add_CheckedChanged({ Apply-Filters })
$chkTrouble.Add_CheckedChanged({ Apply-Filters })
$chkClear.Add_CheckedChanged({ Apply-Filters })

# Double-click on grid row
$grid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    
    try {
        $mail = $script:mailMap[$e.RowIndex]
        if ($mail) { $mail.Display() }
    } catch {}
})

# Form closing
$form.Add_FormClosing({
    $script:isRunning = $false
    if ($script:timer) { $script:timer.Stop(); $script:timer.Dispose() }
})

# =================== DEFAULT RULE ===================
[void]$script:matchRules.Add(@{
    Domain = "-- ALL DOMAINS --"
    MatchType = "Combined (All Fields)"
    Severity = "ALL"
    CustomPattern = ""
})
Update-RulesList

# =================== RUN ===================
[void]$form.ShowDialog()
