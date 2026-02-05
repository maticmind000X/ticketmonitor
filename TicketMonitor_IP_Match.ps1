Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "SilentlyContinue"

# =================== DOUBLE BUFFERING ===================
Add-Type @"
using System;
using System.Windows.Forms;
using System.Reflection;

public class DoubleBufferedDataGridView : DataGridView
{
    public DoubleBufferedDataGridView()
    {
        this.DoubleBuffered = true;
        this.SetStyle(ControlStyles.OptimizedDoubleBuffer | 
                      ControlStyles.AllPaintingInWmPaint | 
                      ControlStyles.UserPaint, true);
        this.UpdateStyles();
    }
}
"@ -ReferencedAssemblies System.Windows.Forms, System.Drawing

# =================== STATO GLOBALE ===================
$script:isRunning = $false
$script:selectedStore = $null
$script:mailMap = @{}
$script:scannedIds = @{}
$script:allEmails = @()
$script:exclusionFilters = [System.Collections.ArrayList]::new()
$script:matchRules = [System.Collections.ArrayList]::new()
$script:alertClearMap = @{}
$script:allDomains = @{}
$script:timer = $null
$script:namespace = $null
$script:outlook = $null
$script:ipToRowMap = @{}
$script:currentFontSize = "Medium"  # Small, Medium, Large

# =================== FONT SIZES ===================
$script:fontSizes = @{
    Small = @{
        Normal = 8
        Grid = 8
        Header = 16
        Stats = 20
        Label = 7
        Button = 8
        RowHeight = 26
    }
    Medium = @{
        Normal = 9
        Grid = 9
        Header = 20
        Stats = 24
        Label = 8
        Button = 9
        RowHeight = 32
    }
    Large = @{
        Normal = 11
        Grid = 11
        Header = 24
        Stats = 30
        Label = 10
        Button = 11
        RowHeight = 40
    }
}

# =================== MATCH FIELD TYPES ===================
$script:matchFieldTypes = @(
    "Combined (All Fields)",
    "IP Address",
    "Device Name",
    "Serial Number",
    "Category",
    "Hostname",
    "Subject Pattern",
    "First Word After CLEAR/CRITICAL",
    "Custom Regex"
)

# =================== TEMI DISPONIBILI ===================
$script:themes = @{
    "Dark Blue" = @{
        bgDeep      = [System.Drawing.Color]::FromArgb(4, 6, 13)
        bgMain      = [System.Drawing.Color]::FromArgb(8, 12, 24)
        bgPanel     = [System.Drawing.Color]::FromArgb(13, 20, 36)
        bgCard      = [System.Drawing.Color]::FromArgb(17, 26, 46)
        bgElevated  = [System.Drawing.Color]::FromArgb(22, 32, 56)
        bgInput     = [System.Drawing.Color]::FromArgb(30, 42, 70)
        border      = [System.Drawing.Color]::FromArgb(50, 80, 120)
        textPrimary = [System.Drawing.Color]::FromArgb(232, 240, 255)
        textSecond  = [System.Drawing.Color]::FromArgb(139, 163, 199)
        textMuted   = [System.Drawing.Color]::FromArgb(90, 112, 148)
        accent      = [System.Drawing.Color]::FromArgb(0, 212, 255)
        critical    = [System.Drawing.Color]::FromArgb(255, 77, 106)
        attention   = [System.Drawing.Color]::FromArgb(255, 204, 0)
        trouble     = [System.Drawing.Color]::FromArgb(255, 140, 66)
        clear       = [System.Drawing.Color]::FromArgb(0, 232, 144)
        resolved    = [System.Drawing.Color]::FromArgb(100, 180, 100)
    }
    "Dark Purple" = @{
        bgDeep      = [System.Drawing.Color]::FromArgb(10, 5, 18)
        bgMain      = [System.Drawing.Color]::FromArgb(18, 10, 30)
        bgPanel     = [System.Drawing.Color]::FromArgb(28, 18, 45)
        bgCard      = [System.Drawing.Color]::FromArgb(38, 25, 60)
        bgElevated  = [System.Drawing.Color]::FromArgb(50, 35, 75)
        bgInput     = [System.Drawing.Color]::FromArgb(60, 45, 90)
        border      = [System.Drawing.Color]::FromArgb(100, 70, 140)
        textPrimary = [System.Drawing.Color]::FromArgb(240, 235, 255)
        textSecond  = [System.Drawing.Color]::FromArgb(180, 160, 210)
        textMuted   = [System.Drawing.Color]::FromArgb(120, 100, 160)
        accent      = [System.Drawing.Color]::FromArgb(180, 100, 255)
        critical    = [System.Drawing.Color]::FromArgb(255, 77, 106)
        attention   = [System.Drawing.Color]::FromArgb(255, 204, 0)
        trouble     = [System.Drawing.Color]::FromArgb(255, 140, 66)
        clear       = [System.Drawing.Color]::FromArgb(0, 232, 144)
        resolved    = [System.Drawing.Color]::FromArgb(100, 180, 100)
    }
    "Dark Green" = @{
        bgDeep      = [System.Drawing.Color]::FromArgb(5, 12, 8)
        bgMain      = [System.Drawing.Color]::FromArgb(8, 20, 14)
        bgPanel     = [System.Drawing.Color]::FromArgb(12, 30, 22)
        bgCard      = [System.Drawing.Color]::FromArgb(18, 42, 32)
        bgElevated  = [System.Drawing.Color]::FromArgb(25, 55, 42)
        bgInput     = [System.Drawing.Color]::FromArgb(35, 70, 55)
        border      = [System.Drawing.Color]::FromArgb(60, 110, 85)
        textPrimary = [System.Drawing.Color]::FromArgb(230, 255, 240)
        textSecond  = [System.Drawing.Color]::FromArgb(150, 200, 175)
        textMuted   = [System.Drawing.Color]::FromArgb(100, 150, 125)
        accent      = [System.Drawing.Color]::FromArgb(0, 255, 180)
        critical    = [System.Drawing.Color]::FromArgb(255, 77, 106)
        attention   = [System.Drawing.Color]::FromArgb(255, 204, 0)
        trouble     = [System.Drawing.Color]::FromArgb(255, 140, 66)
        clear       = [System.Drawing.Color]::FromArgb(0, 232, 144)
        resolved    = [System.Drawing.Color]::FromArgb(100, 180, 100)
    }
    "Light" = @{
        bgDeep      = [System.Drawing.Color]::FromArgb(245, 247, 250)
        bgMain      = [System.Drawing.Color]::FromArgb(238, 242, 248)
        bgPanel     = [System.Drawing.Color]::FromArgb(255, 255, 255)
        bgCard      = [System.Drawing.Color]::FromArgb(250, 252, 255)
        bgElevated  = [System.Drawing.Color]::FromArgb(235, 240, 248)
        bgInput     = [System.Drawing.Color]::FromArgb(255, 255, 255)
        border      = [System.Drawing.Color]::FromArgb(200, 210, 225)
        textPrimary = [System.Drawing.Color]::FromArgb(30, 40, 60)
        textSecond  = [System.Drawing.Color]::FromArgb(80, 95, 120)
        textMuted   = [System.Drawing.Color]::FromArgb(130, 145, 165)
        accent      = [System.Drawing.Color]::FromArgb(0, 120, 215)
        critical    = [System.Drawing.Color]::FromArgb(220, 50, 80)
        attention   = [System.Drawing.Color]::FromArgb(200, 160, 0)
        trouble     = [System.Drawing.Color]::FromArgb(220, 120, 50)
        clear       = [System.Drawing.Color]::FromArgb(0, 160, 100)
        resolved    = [System.Drawing.Color]::FromArgb(80, 150, 80)
    }
    "Midnight" = @{
        bgDeep      = [System.Drawing.Color]::FromArgb(0, 0, 0)
        bgMain      = [System.Drawing.Color]::FromArgb(10, 10, 12)
        bgPanel     = [System.Drawing.Color]::FromArgb(18, 18, 22)
        bgCard      = [System.Drawing.Color]::FromArgb(25, 25, 30)
        bgElevated  = [System.Drawing.Color]::FromArgb(35, 35, 42)
        bgInput     = [System.Drawing.Color]::FromArgb(45, 45, 55)
        border      = [System.Drawing.Color]::FromArgb(70, 70, 85)
        textPrimary = [System.Drawing.Color]::FromArgb(255, 255, 255)
        textSecond  = [System.Drawing.Color]::FromArgb(180, 180, 190)
        textMuted   = [System.Drawing.Color]::FromArgb(120, 120, 135)
        accent      = [System.Drawing.Color]::FromArgb(100, 200, 255)
        critical    = [System.Drawing.Color]::FromArgb(255, 77, 106)
        attention   = [System.Drawing.Color]::FromArgb(255, 204, 0)
        trouble     = [System.Drawing.Color]::FromArgb(255, 140, 66)
        clear       = [System.Drawing.Color]::FromArgb(0, 232, 144)
        resolved    = [System.Drawing.Color]::FromArgb(100, 180, 100)
    }
}

$script:currentTheme = "Dark Blue"
$theme = $script:themes[$script:currentTheme]

# =================== SUONI PER SEVERITY ===================
$script:soundEnabled = $true
$script:severitySounds = @{
    CRITICAL = @{ Frequency = 1000; Duration = 500; Repeat = 3 }    # Acuto, lungo, 3 volte
    ATTENTION = @{ Frequency = 700; Duration = 300; Repeat = 2 }    # Medio, 2 volte
    TROUBLE = @{ Frequency = 500; Duration = 200; Repeat = 1 }      # Basso, 1 volta
}

function Play-SeveritySound {
    param([string]$Severity)
    
    if (-not $script:soundEnabled) { return }
    if (-not $chkSound.Checked) { return }
    
    try {
        $sound = $script:severitySounds[$Severity]
        if ($sound) {
            for ($i = 0; $i -lt $sound.Repeat; $i++) {
                [Console]::Beep($sound.Frequency, $sound.Duration)
                if ($i -lt ($sound.Repeat - 1)) {
                    Start-Sleep -Milliseconds 100
                }
            }
        }
    } catch {}
}

# =================== FUNZIONI UTILITY ===================
function Get-EmailDomain {
    param([string]$Email)
    
    # Se l'email è vuota o null
    if ([string]::IsNullOrWhiteSpace($Email)) {
        return "no_sender"
    }
    
    # Caso 0: Formato "name:NomeMittente" (nostro fallback)
    if ($Email -match "^name:(.+)$") {
        $name = $matches[1] -replace '[^a-zA-Z0-9]', '_'
        if ($name.Length -gt 25) { $name = $name.Substring(0, 25) }
        return ("sender_" + $name).ToLower()
    }
    
    # Caso 1: Email standard con @
    if ($Email -match "@(.+)$") { 
        return $matches[1].ToLower().Trim() 
    }
    
    # Caso 2: Formato Exchange /O=ORG/OU=.../CN=...
    if ($Email -match "^/O=([^/]+)") { 
        return ("ex_" + $matches[1]).ToLower() 
    }
    
    # Caso 3: Formato X500 o simili
    if ($Email -match "^/") { 
        return "exchange_internal" 
    }
    
    # Caso 4: Potrebbe essere già solo il dominio
    if ($Email -match "^[a-zA-Z0-9\-]+\.[a-zA-Z]{2,}$") {
        return $Email.ToLower()
    }
    
    # Caso 5: Se contiene un punto, prova a estrarre qualcosa di utile
    if ($Email -match "\.([a-zA-Z0-9\-]+\.[a-zA-Z]{2,})$") {
        return $matches[1].ToLower()
    }
    
    # Ultimo tentativo: usa i primi 20 caratteri come identificatore
    $clean = $Email -replace '[^a-zA-Z0-9]', '_'
    if ($clean.Length -gt 20) { $clean = $clean.Substring(0, 20) }
    return ("other_" + $clean).ToLower()
}

function Extract-IPAddresses {
    param([string]$Text)
    
    $ips = @()
    $pattern = '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'
    
    $foundMatches = [regex]::Matches($Text, $pattern)
    foreach ($m in $foundMatches) {
        $ip = $m.Groups[1].Value
        if ($ip -and $ip -notin $ips) {
            $octets = $ip -split '\.'
            $valid = $true
            foreach ($octet in $octets) {
                $num = [int]$octet
                if ($num -lt 0 -or $num -gt 255) { $valid = $false; break }
            }
            if ($valid) { $ips += $ip }
        }
    }
    
    return $ips
}

function Extract-DeviceName {
    param([string]$Text)
    
    $devices = @()
    # Pattern per "Device: nome" o "Device Name: nome"
    $patterns = @(
        'Device\s*(?:Name)?\s*[:=]\s*([^\r\n,;]+)',
        'Nome\s*(?:Device|Dispositivo)?\s*[:=]\s*([^\r\n,;]+)'
    )
    
    foreach ($pattern in $patterns) {
        $foundMatches = [regex]::Matches($Text, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($m in $foundMatches) {
            $device = $m.Groups[1].Value.Trim()
            if ($device -and $device.Length -gt 1 -and $device -notin $devices) {
                $devices += $device
            }
        }
    }
    
    return $devices
}

function Extract-SerialNumber {
    param([string]$Text)
    
    $serials = @()
    # Pattern per "Serial Number: XXX" o "Serial: XXX" o "S/N: XXX"
    $patterns = @(
        'Serial\s*(?:Number|No\.?)?\s*[:=]\s*([A-Za-z0-9\-]+)',
        'S/?N\s*[:=]\s*([A-Za-z0-9\-]+)'
    )
    
    foreach ($pattern in $patterns) {
        $foundMatches = [regex]::Matches($Text, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($m in $foundMatches) {
            $serial = $m.Groups[1].Value.Trim()
            if ($serial -and $serial.Length -gt 3 -and $serial -notin $serials) {
                $serials += $serial
            }
        }
    }
    
    return $serials
}

function Extract-Category {
    param([string]$Text)
    
    $categories = @()
    # Pattern per "Category: XXX" o "Categoria: XXX" o "Type: XXX"
    $patterns = @(
        'Category\s*[:=]\s*([^\r\n,;]+)',
        'Categoria\s*[:=]\s*([^\r\n,;]+)',
        'Type\s*[:=]\s*([^\r\n,;]+)',
        'Tipo\s*[:=]\s*([^\r\n,;]+)'
    )
    
    foreach ($pattern in $patterns) {
        $foundMatches = [regex]::Matches($Text, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($m in $foundMatches) {
            $category = $m.Groups[1].Value.Trim()
            if ($category -and $category.Length -gt 1 -and $category -notin $categories) {
                $categories += $category
            }
        }
    }
    
    return $categories
}

function Extract-CombinedKey {
    param([string]$Text)
    
    # Estrai tutti i campi disponibili
    $ips = Extract-IPAddresses -Text $Text
    $devices = Extract-DeviceName -Text $Text
    $serials = Extract-SerialNumber -Text $Text
    $categories = Extract-Category -Text $Text
    
    # Costruisci una lista di chiavi individuali (ogni campo diventa una chiave)
    $keys = @()
    
    # Aggiungi ogni IP trovato
    foreach ($ip in $ips) {
        $keys += "IP:$($ip.ToLower())"
    }
    
    # Aggiungi ogni Device trovato
    foreach ($dev in $devices) {
        $keys += "DEV:$($dev.ToLower())"
    }
    
    # Aggiungi ogni Serial Number trovato
    foreach ($sn in $serials) {
        $keys += "SN:$($sn.ToLower())"
    }
    
    # Aggiungi ogni Category trovata
    foreach ($cat in $categories) {
        $keys += "CAT:$($cat.ToLower())"
    }
    
    return $keys
}

# Funzione per verificare se due set di chiavi matchano (almeno 2 parametri in comune)
function Test-CombinedMatch {
    param(
        [array]$Keys1,
        [array]$Keys2,
        [int]$MinMatches = 2
    )
    
    if ($Keys1.Count -eq 0 -or $Keys2.Count -eq 0) { return $false }
    
    $matchCount = 0
    $matchedKeys = @()
    
    foreach ($key1 in $Keys1) {
        if ($Keys2 -contains $key1) {
            $matchCount++
            $matchedKeys += $key1
        }
    }
    
    return ($matchCount -ge $MinMatches)
}

function Extract-Hostnames {
    param([string]$Text)
    
    $hostnames = @()
    $pattern = '\b([a-zA-Z][a-zA-Z0-9\-]{2,}(?:\.[a-zA-Z0-9\-]+)*)\b'
    
    $foundMatches = [regex]::Matches($Text, $pattern)
    foreach ($m in $foundMatches) {
        $hostname = $m.Groups[1].Value.ToLower()
        if ($hostname -notin @("critical", "attention", "trouble", "clear", "alert", "warning", "error", "info", "the", "and", "for", "from", "subject", "body", "email") -and 
            $hostname -notin $hostnames -and 
            $hostname.Length -gt 2) {
            $hostnames += $hostname
        }
    }
    
    return $hostnames
}

function Extract-FirstWordAfterSeverity {
    param([string]$Text)
    
    if ($Text -match '(?:CRITICAL|ATTENTION|TROUBLE|CLEAR)[:\s\-]+([A-Za-z0-9\-_.]+)') {
        return $matches[1].ToLower()
    }
    return $null
}

function Extract-SubjectPattern {
    param([string]$Subject)
    
    $pattern = $Subject -replace '(?i)(CRITICAL|ATTENTION|TROUBLE|CLEAR)', ''
    $pattern = $pattern -replace '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', '[IP]'
    $pattern = $pattern -replace '\d{2,}', '[NUM]'
    $pattern = $pattern.Trim() -replace '\s+', ' '
    return $pattern.ToLower()
}

function Extract-CustomRegex {
    param([string]$Text, [string]$Pattern)
    
    try {
        if ($Text -match $Pattern) {
            if ($matches[1]) { return $matches[1].ToLower() }
            return $matches[0].ToLower()
        }
    } catch {}
    return $null
}

function Get-MatchKeyForEmail {
    param($Email, [string]$MatchType, [string]$CustomPattern = "")
    
    $text = "$($Email.Subject) $($Email.Body)"
    
    switch ($MatchType) {
        "Combined (All Fields)" {
            return Extract-CombinedKey -Text $text
        }
        "IP Address" {
            return $Email.IPs
        }
        "Device Name" {
            return Extract-DeviceName -Text $text
        }
        "Serial Number" {
            return Extract-SerialNumber -Text $text
        }
        "Category" {
            return Extract-Category -Text $text
        }
        "Hostname" {
            return Extract-Hostnames -Text $text
        }
        "Subject Pattern" {
            $pattern = Extract-SubjectPattern -Subject $Email.Subject
            if ($pattern) { return @($pattern) }
            return @()
        }
        "First Word After CLEAR/CRITICAL" {
            $word = Extract-FirstWordAfterSeverity -Text $text
            if ($word) { return @($word) }
            return @()
        }
        "Custom Regex" {
            $result = Extract-CustomRegex -Text $text -Pattern $CustomPattern
            if ($result) { return @($result) }
            return @()
        }
        default {
            return $Email.IPs
        }
    }
}

function Get-MatchRuleForDomain {
    param([string]$Domain, [string]$Severity = "")
    
    # Cerca regola specifica per il dominio E severity
    foreach ($rule in $script:matchRules) {
        if ($rule.Domain -eq $Domain) {
            if ($rule.Severity -eq $Severity -or $rule.Severity -eq "ALL" -or [string]::IsNullOrEmpty($rule.Severity)) {
                return $rule
            }
        }
    }
    
    # Cerca regola "ALL DOMAINS" con severity specifica o ALL
    foreach ($rule in $script:matchRules) {
        if ($rule.Domain -eq "-- ALL DOMAINS --") {
            if ($rule.Severity -eq $Severity -or $rule.Severity -eq "ALL" -or [string]::IsNullOrEmpty($rule.Severity)) {
                return $rule
            }
        }
    }
    
    # NESSUNA REGOLA = NESSUN MATCH
    return $null
}

function Get-SenderInfo {
    param($Mail)
    $senderName = $Mail.SenderName
    $senderEmail = ""
    
    try {
        # Metodo 1: Per email Exchange (EX)
        if ($Mail.SenderEmailType -eq "EX") {
            # Prova GetExchangeUser
            try {
                $exUser = $Mail.Sender.GetExchangeUser()
                if ($exUser -and $exUser.PrimarySmtpAddress) { 
                    $senderEmail = $exUser.PrimarySmtpAddress 
                }
            } catch {}
            
            # Prova PR_SMTP_ADDRESS
            if ([string]::IsNullOrEmpty($senderEmail)) {
                try {
                    $PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
                    $senderEmail = $Mail.PropertyAccessor.GetProperty($PR_SMTP_ADDRESS)
                } catch {}
            }
            
            # Prova PR_SENDER_EMAIL_ADDRESS
            if ([string]::IsNullOrEmpty($senderEmail)) {
                try {
                    $PR_SENDER_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"
                    $senderEmail = $Mail.PropertyAccessor.GetProperty($PR_SENDER_EMAIL)
                } catch {}
            }
            
            # Prova SenderEmailAddress diretto
            if ([string]::IsNullOrEmpty($senderEmail)) {
                $senderEmail = $Mail.SenderEmailAddress
            }
        }
        # Metodo 2: Per email SMTP standard
        elseif ($Mail.SenderEmailType -eq "SMTP") {
            $senderEmail = $Mail.SenderEmailAddress
        }
        # Metodo 3: Altri tipi
        else {
            # Prova SenderEmailAddress
            $senderEmail = $Mail.SenderEmailAddress
            
            # Se vuoto, prova Reply-To
            if ([string]::IsNullOrEmpty($senderEmail)) {
                try {
                    $senderEmail = $Mail.ReplyRecipients.Item(1).Address
                } catch {}
            }
        }
        
        # Fallback: Se ancora vuoto, usa il SenderName come identificatore
        if ([string]::IsNullOrEmpty($senderEmail)) { 
            if (-not [string]::IsNullOrEmpty($senderName)) {
                # Usa il nome del mittente come pseudo-dominio
                $senderEmail = "name:" + $senderName
            } else {
                $senderEmail = "unknown_sender"
            }
        }
    } catch {
        if (-not [string]::IsNullOrEmpty($senderName)) {
            $senderEmail = "name:" + $senderName
        } else {
            $senderEmail = "error_reading_sender"
        }
    }
    
    return @{ Email = $senderEmail; Name = $senderName }
}

function Test-EmailExcluded {
    param($Email)
    $content = "$($Email.Subject) $($Email.Sender) $($Email.Body)".ToLower()
    
    foreach ($filter in $script:exclusionFilters) {
        if ($filter.Domain -eq "-- ALL --" -or $filter.Domain -eq $Email.Domain) {
            if ($content.Contains($filter.Word.ToLower())) {
                return $true
            }
        }
    }
    return $false
}

function Set-RowColor {
    param($Row, $Severity, $HasClear = $false)
    
    $clearedColor = [System.Drawing.Color]::FromArgb(255, 255, 0)
    
    switch ($Severity) {
        "CRITICAL" {
            if ($HasClear) {
                $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30, 255, 77, 106)
                $Row.DefaultCellStyle.ForeColor = $clearedColor
                $font = $Row.DataGridView.DefaultCellStyle.Font
                if ($font) {
                    $Row.DefaultCellStyle.Font = New-Object System.Drawing.Font($font.FontFamily, $font.Size, [System.Drawing.FontStyle]::Strikeout)
                }
            } else {
                $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40, 255, 77, 106)
                $Row.DefaultCellStyle.ForeColor = $theme.critical
            }
        }
        "ATTENTION" {
            if ($HasClear) {
                $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30, 255, 204, 0)
                $Row.DefaultCellStyle.ForeColor = $clearedColor
                $font = $Row.DataGridView.DefaultCellStyle.Font
                if ($font) {
                    $Row.DefaultCellStyle.Font = New-Object System.Drawing.Font($font.FontFamily, $font.Size, [System.Drawing.FontStyle]::Strikeout)
                }
            } else {
                $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40, 255, 204, 0)
                $Row.DefaultCellStyle.ForeColor = $theme.attention
            }
        }
        "TROUBLE" {
            if ($HasClear) {
                $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30, 255, 140, 66)
                $Row.DefaultCellStyle.ForeColor = $clearedColor
                $font = $Row.DataGridView.DefaultCellStyle.Font
                if ($font) {
                    $Row.DefaultCellStyle.Font = New-Object System.Drawing.Font($font.FontFamily, $font.Size, [System.Drawing.FontStyle]::Strikeout)
                }
            } else {
                $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(40, 255, 140, 66)
                $Row.DefaultCellStyle.ForeColor = $theme.trouble
            }
        }
        "CLEAR" {
            $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(25, 0, 232, 144)
            $Row.DefaultCellStyle.ForeColor = $theme.clear
        }
    }
}

# =================== OUTLOOK CONNECTION ===================
try {
    $script:outlook = New-Object -ComObject Outlook.Application
    $script:namespace = $script:outlook.GetNamespace("MAPI")
} catch {
    [System.Windows.Forms.MessageBox]::Show("Impossibile connettersi a Outlook. Assicurati che sia aperto.", "Errore", "OK", "Error")
    exit
}

# =================== RESOLUTION PRESETS ===================
$script:resolutionPresets = @{
    "Auto" = @{ Width = 0; Height = 0 }
    "1920x1080" = @{ Width = 1650; Height = 950 }
    "1680x1050" = @{ Width = 1500; Height = 900 }
    "1600x900" = @{ Width = 1400; Height = 800 }
    "1440x900" = @{ Width = 1300; Height = 800 }
    "1366x768" = @{ Width = 1200; Height = 700 }
    "1280x720" = @{ Width = 1100; Height = 650 }
    "1024x768" = @{ Width = 950; Height = 700 }
}

# =================== FORM PRINCIPALE ===================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Ticket Monitor Dashboard - Custom Match v2"
$form.Size = New-Object System.Drawing.Size(1650, 1000)
$form.MinimumSize = New-Object System.Drawing.Size(950, 600)
$form.StartPosition = "CenterScreen"
$form.BackColor = $theme.bgDeep
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# =================== DATA GRID ===================
$grid = New-Object DoubleBufferedDataGridView
$grid.Dock = "Fill"
$grid.BackgroundColor = $theme.bgDeep
$grid.ForeColor = $theme.textPrimary
$grid.GridColor = $theme.border
$grid.BorderStyle = "None"
$grid.CellBorderStyle = "SingleHorizontal"
$grid.ColumnHeadersBorderStyle = "Single"
$grid.EnableHeadersVisualStyles = $false
$grid.RowHeadersVisible = $false
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "Fill"
$grid.RowTemplate.Height = 32

$grid.DefaultCellStyle.BackColor = $theme.bgDeep
$grid.DefaultCellStyle.ForeColor = $theme.textPrimary
$grid.DefaultCellStyle.SelectionBackColor = $theme.bgElevated
$grid.DefaultCellStyle.SelectionForeColor = $theme.textPrimary
$grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 9)
$grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)

$grid.ColumnHeadersDefaultCellStyle.BackColor = $theme.bgPanel
$grid.ColumnHeadersDefaultCellStyle.ForeColor = $theme.textMuted
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 0, 8, 0)
$grid.ColumnHeadersHeight = 40

$colSev = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSev.Name = "Sev"; $colSev.HeaderText = "SEVERITY"; $colSev.Width = 140
[void]$grid.Columns.Add($colSev)

$colDomain = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDomain.Name = "Domain"; $colDomain.HeaderText = "DOMAIN"; $colDomain.Width = 150
[void]$grid.Columns.Add($colDomain)

$colSender = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSender.Name = "Sender"; $colSender.HeaderText = "SENDER"; $colSender.Width = 180
[void]$grid.Columns.Add($colSender)

$colSubject = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSubject.Name = "Subject"; $colSubject.HeaderText = "SUBJECT"; $colSubject.FillWeight = 100
[void]$grid.Columns.Add($colSubject)

$colTime = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colTime.Name = "Time"; $colTime.HeaderText = "TIME"; $colTime.Width = 150
[void]$grid.Columns.Add($colTime)

$colMatchKey = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colMatchKey.Name = "MatchKey"; $colMatchKey.HeaderText = "MATCH KEY"; $colMatchKey.Width = 150
[void]$grid.Columns.Add($colMatchKey)

$colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colStatus.Name = "Status"; $colStatus.HeaderText = "STATUS"; $colStatus.Width = 100
[void]$grid.Columns.Add($colStatus)

$form.Controls.Add($grid)

# =================== HEADER ===================
$panelHeader = New-Object System.Windows.Forms.Panel
$panelHeader.Dock = "Top"
$panelHeader.Height = 80
$panelHeader.BackColor = $theme.bgPanel
$form.Controls.Add($panelHeader)

$lblLogo = New-Object System.Windows.Forms.Label
$lblLogo.Text = "TICKET MONITOR"
$lblLogo.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
$lblLogo.ForeColor = $theme.accent
$lblLogo.Location = New-Object System.Drawing.Point(25, 22)
$lblLogo.AutoSize = $true
$panelHeader.Controls.Add($lblLogo)

$panelStatus = New-Object System.Windows.Forms.Panel
$panelStatus.Size = New-Object System.Drawing.Size(200, 40)
$panelStatus.Location = New-Object System.Drawing.Point(1420, 20)
$panelStatus.BackColor = $theme.bgCard
$panelHeader.Controls.Add($panelStatus)

$lblStatusDot = New-Object System.Windows.Forms.Label
$lblStatusDot.Text = "*"
$lblStatusDot.Font = New-Object System.Drawing.Font("Segoe UI", 14)
$lblStatusDot.ForeColor = $theme.textMuted
$lblStatusDot.Location = New-Object System.Drawing.Point(15, 8)
$lblStatusDot.AutoSize = $true
$panelStatus.Controls.Add($lblStatusDot)

$lblStatusText = New-Object System.Windows.Forms.Label
$lblStatusText.Text = "Ready"
$lblStatusText.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$lblStatusText.ForeColor = $theme.textSecond
$lblStatusText.Location = New-Object System.Drawing.Point(40, 10)
$lblStatusText.AutoSize = $true
$panelStatus.Controls.Add($lblStatusText)

# =================== FONT SIZE SELECTOR (in header) ===================
$lblFontSize = New-Object System.Windows.Forms.Label
$lblFontSize.Text = "TEXT SIZE:"
$lblFontSize.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblFontSize.ForeColor = $theme.accent
$lblFontSize.Location = New-Object System.Drawing.Point(350, 28)
$lblFontSize.AutoSize = $true
$panelHeader.Controls.Add($lblFontSize)

$cmbFontSize = New-Object System.Windows.Forms.ComboBox
$cmbFontSize.Location = New-Object System.Drawing.Point(450, 25)
$cmbFontSize.Size = New-Object System.Drawing.Size(100, 28)
$cmbFontSize.DropDownStyle = "DropDownList"
$cmbFontSize.BackColor = $theme.bgInput
$cmbFontSize.ForeColor = $theme.textPrimary
$cmbFontSize.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$cmbFontSize.FlatStyle = "Flat"
[void]$cmbFontSize.Items.Add("Small")
[void]$cmbFontSize.Items.Add("Medium")
[void]$cmbFontSize.Items.Add("Large")
$cmbFontSize.SelectedItem = "Medium"
$panelHeader.Controls.Add($cmbFontSize)

# =================== RESOLUTION SELECTOR (in header) ===================
$lblResolution = New-Object System.Windows.Forms.Label
$lblResolution.Text = "RESOLUTION:"
$lblResolution.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblResolution.ForeColor = $theme.attention
$lblResolution.Location = New-Object System.Drawing.Point(560, 28)
$lblResolution.AutoSize = $true
$panelHeader.Controls.Add($lblResolution)

$cmbResolution = New-Object System.Windows.Forms.ComboBox
$cmbResolution.Location = New-Object System.Drawing.Point(680, 25)
$cmbResolution.Size = New-Object System.Drawing.Size(110, 28)
$cmbResolution.DropDownStyle = "DropDownList"
$cmbResolution.BackColor = $theme.bgInput
$cmbResolution.ForeColor = $theme.textPrimary
$cmbResolution.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$cmbResolution.FlatStyle = "Flat"
[void]$cmbResolution.Items.Add("Auto")
[void]$cmbResolution.Items.Add("1920x1080")
[void]$cmbResolution.Items.Add("1680x1050")
[void]$cmbResolution.Items.Add("1600x900")
[void]$cmbResolution.Items.Add("1440x900")
[void]$cmbResolution.Items.Add("1366x768")
[void]$cmbResolution.Items.Add("1280x720")
[void]$cmbResolution.Items.Add("1024x768")
$cmbResolution.SelectedItem = "Auto"
$panelHeader.Controls.Add($cmbResolution)

# =================== THEME SELECTOR (nell'header, sotto TEXT SIZE) ===================
$lblTheme = New-Object System.Windows.Forms.Label
$lblTheme.Text = "THEME:"
$lblTheme.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblTheme.ForeColor = $theme.clear
$lblTheme.Location = New-Object System.Drawing.Point(350, 50)
$lblTheme.AutoSize = $true
$panelHeader.Controls.Add($lblTheme)

$cmbTheme = New-Object System.Windows.Forms.ComboBox
$cmbTheme.Location = New-Object System.Drawing.Point(410, 47)
$cmbTheme.Size = New-Object System.Drawing.Size(100, 24)
$cmbTheme.DropDownStyle = "DropDownList"
$cmbTheme.BackColor = $theme.bgInput
$cmbTheme.ForeColor = $theme.textPrimary
$cmbTheme.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$cmbTheme.FlatStyle = "Flat"
foreach ($themeName in $script:themes.Keys | Sort-Object) {
    [void]$cmbTheme.Items.Add($themeName)
}
$cmbTheme.SelectedItem = "Dark Blue"
$panelHeader.Controls.Add($cmbTheme)

# Funzione per applicare il tema a tutti i controlli
function Apply-Theme {
    param([string]$ThemeName)
    
    $script:currentTheme = $ThemeName
    $script:theme = $script:themes[$ThemeName]
    $global:theme = $script:theme
    
    # Form e panels principali
    $form.BackColor = $script:theme.bgMain
    $panelHeader.BackColor = $script:theme.bgPanel
    $panelStats.BackColor = $script:theme.bgDeep
    $panelControls.BackColor = $script:theme.bgPanel
    $panelFilter.BackColor = $script:theme.bgElevated
    $panelExcludeBar.BackColor = [System.Drawing.Color]::FromArgb(60, 30, 30)
    $panelMatch.BackColor = [System.Drawing.Color]::FromArgb(30, 45, 80)
    $panelLists.BackColor = $script:theme.bgMain
    
    # Header elements
    $lblLogo.ForeColor = $script:theme.accent
    $lblFontSize.ForeColor = $script:theme.accent
    $lblResolution.ForeColor = $script:theme.attention
    $lblTheme.ForeColor = $script:theme.clear
    
    # Combo boxes
    $cmbFontSize.BackColor = $script:theme.bgInput
    $cmbFontSize.ForeColor = $script:theme.textPrimary
    $cmbResolution.BackColor = $script:theme.bgInput
    $cmbResolution.ForeColor = $script:theme.textPrimary
    $cmbTheme.BackColor = $script:theme.bgInput
    $cmbTheme.ForeColor = $script:theme.textPrimary
    $cmbProfile.BackColor = $script:theme.bgInput
    $cmbProfile.ForeColor = $script:theme.textPrimary
    
    # Grid
    $grid.BackgroundColor = $script:theme.bgDeep
    $grid.DefaultCellStyle.BackColor = $script:theme.bgCard
    $grid.DefaultCellStyle.ForeColor = $script:theme.textPrimary
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $script:theme.bgElevated
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = $script:theme.textPrimary
    $grid.GridColor = $script:theme.border
    
    # Status
    $panelStatus.BackColor = $script:theme.bgElevated
    $lblStatusText.ForeColor = $script:theme.textSecond
    
    # GroupBoxes
    $grpExclusions.ForeColor = $script:theme.critical
    $grpMatchRules.ForeColor = $script:theme.accent
    $listExclusions.BackColor = $script:theme.bgInput
    $listExclusions.ForeColor = $script:theme.textPrimary
    $listMatchRules.BackColor = $script:theme.bgInput
    $listMatchRules.ForeColor = $script:theme.textPrimary
    
    # Pie chart panel
    if ($script:piePanel) { $script:piePanel.Refresh() }
    
    # Refresh
    $form.Refresh()
}

$cmbTheme.Add_SelectedIndexChanged({
    Apply-Theme -ThemeName $cmbTheme.SelectedItem
})

# Scala di riferimento (design originale)
$script:baseWidth = 1650

# Funzione per adattare TUTTO il layout alla risoluzione
function Adjust-LayoutToResolution {
    param([int]$newWidth, [int]$newHeight)
    
    $w = $newWidth
    
    # Determina modalità layout
    $isCompact = $w -lt 1300
    $isVeryCompact = $w -lt 1100
    
    if ($isVeryCompact) {
        # ===== LAYOUT MOLTO COMPATTO (< 1100px) =====
        
        $panelHeader.Height = 50
        $lblLogo.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        $lblLogo.Location = New-Object System.Drawing.Point(8, 14)
        
        $lblFontSize.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
        $lblFontSize.Location = New-Object System.Drawing.Point(160, 16)
        $cmbFontSize.Location = New-Object System.Drawing.Point(215, 12)
        $cmbFontSize.Size = New-Object System.Drawing.Size(60, 22)
        
        $lblResolution.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
        $lblResolution.Location = New-Object System.Drawing.Point(280, 16)
        $cmbResolution.Location = New-Object System.Drawing.Point(345, 12)
        $cmbResolution.Size = New-Object System.Drawing.Size(75, 22)
        
        $btnEmailTemplate.Text = "EMAIL"
        $btnEmailTemplate.Location = New-Object System.Drawing.Point(430, 8)
        $btnEmailTemplate.Size = New-Object System.Drawing.Size(55, 32)
        $btnEmailTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
        
        $btnManagement.Text = "MGMT"
        $btnManagement.Location = New-Object System.Drawing.Point(490, 8)
        $btnManagement.Size = New-Object System.Drawing.Size(50, 32)
        $btnManagement.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
        
        $btnGuida.Location = New-Object System.Drawing.Point(545, 8)
        $btnGuida.Size = New-Object System.Drawing.Size(45, 32)
        $btnGuida.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
        
        $btnAutoTicketApp.Text = "TICKET"
        $btnAutoTicketApp.Location = New-Object System.Drawing.Point(595, 8)
        $btnAutoTicketApp.Size = New-Object System.Drawing.Size(55, 32)
        $btnAutoTicketApp.Font = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Bold)
        
        $panelStatus.Location = New-Object System.Drawing.Point(($w - 115), 10)
        $panelStatus.Size = New-Object System.Drawing.Size(105, 28)
        $lblStatusDot.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $lblStatusText.Font = New-Object System.Drawing.Font("Segoe UI", 8)
        
        $panelStats.Height = 60
        $panelControls.Height = 45
        $panelFilter.Height = 35
        $panelExcludeBar.Height = 38
        $panelMatch.Height = 38
        $panelLists.Height = 80
        
        # Liste exclusions e match rules
        $halfWidth = [int](($w - 40) / 2)
        $grpExclusions.Location = New-Object System.Drawing.Point(10, 3)
        $grpExclusions.Size = New-Object System.Drawing.Size($halfWidth, 72)
        $listExclusions.Size = New-Object System.Drawing.Size(($halfWidth - 20), 52)
        
        $grpMatchRules.Location = New-Object System.Drawing.Point(($halfWidth + 20), 3)
        $grpMatchRules.Size = New-Object System.Drawing.Size($halfWidth, 72)
        $listMatchRules.Size = New-Object System.Drawing.Size(($halfWidth - 20), 52)
        
        $grid.RowTemplate.Height = 24
        $grid.ColumnHeadersHeight = 28
        $grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 7)
        
    } elseif ($isCompact) {
        # ===== LAYOUT COMPATTO (1100-1300px) =====
        
        $panelHeader.Height = 60
        $lblLogo.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
        $lblLogo.Location = New-Object System.Drawing.Point(12, 18)
        
        $lblFontSize.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        $lblFontSize.Location = New-Object System.Drawing.Point(195, 20)
        $cmbFontSize.Location = New-Object System.Drawing.Point(260, 16)
        $cmbFontSize.Size = New-Object System.Drawing.Size(70, 24)
        
        $lblResolution.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        $lblResolution.Location = New-Object System.Drawing.Point(340, 20)
        $cmbResolution.Location = New-Object System.Drawing.Point(420, 16)
        $cmbResolution.Size = New-Object System.Drawing.Size(85, 24)
        
        $btnEmailTemplate.Text = "EMAIL"
        $btnEmailTemplate.Location = New-Object System.Drawing.Point(515, 12)
        $btnEmailTemplate.Size = New-Object System.Drawing.Size(70, 34)
        $btnEmailTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        
        $btnManagement.Text = "MGMT"
        $btnManagement.Location = New-Object System.Drawing.Point(590, 12)
        $btnManagement.Size = New-Object System.Drawing.Size(65, 34)
        $btnManagement.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        
        $btnGuida.Location = New-Object System.Drawing.Point(660, 12)
        $btnGuida.Size = New-Object System.Drawing.Size(55, 34)
        $btnGuida.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        
        $btnAutoTicketApp.Text = "TICKET"
        $btnAutoTicketApp.Location = New-Object System.Drawing.Point(720, 12)
        $btnAutoTicketApp.Size = New-Object System.Drawing.Size(65, 34)
        $btnAutoTicketApp.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
        
        $panelStatus.Location = New-Object System.Drawing.Point(($w - 145), 14)
        $panelStatus.Size = New-Object System.Drawing.Size(130, 32)
        $lblStatusDot.Font = New-Object System.Drawing.Font("Segoe UI", 12)
        $lblStatusText.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        
        $panelStats.Height = 75
        $panelControls.Height = 50
        $panelFilter.Height = 40
        $panelExcludeBar.Height = 42
        $panelMatch.Height = 42
        $panelLists.Height = 95
        
        # Liste exclusions e match rules
        $halfWidth = [int](($w - 40) / 2)
        $grpExclusions.Location = New-Object System.Drawing.Point(12, 3)
        $grpExclusions.Size = New-Object System.Drawing.Size($halfWidth, 86)
        $listExclusions.Size = New-Object System.Drawing.Size(($halfWidth - 20), 62)
        
        $grpMatchRules.Location = New-Object System.Drawing.Point(($halfWidth + 22), 3)
        $grpMatchRules.Size = New-Object System.Drawing.Size($halfWidth, 86)
        $listMatchRules.Size = New-Object System.Drawing.Size(($halfWidth - 20), 62)
        
        $grid.RowTemplate.Height = 28
        $grid.ColumnHeadersHeight = 32
        $grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 8)
        
    } else {
        # ===== LAYOUT NORMALE (>= 1300px) =====
        
        $panelHeader.Height = 80
        $lblLogo.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
        $lblLogo.Location = New-Object System.Drawing.Point(25, 22)
        
        $lblFontSize.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $lblFontSize.Location = New-Object System.Drawing.Point(350, 28)
        $cmbFontSize.Location = New-Object System.Drawing.Point(450, 25)
        $cmbFontSize.Size = New-Object System.Drawing.Size(100, 28)
        
        $lblResolution.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $lblResolution.Location = New-Object System.Drawing.Point(560, 28)
        $cmbResolution.Location = New-Object System.Drawing.Point(680, 25)
        $cmbResolution.Size = New-Object System.Drawing.Size(110, 28)
        
        $btnEmailTemplate.Text = "SEND EMAIL"
        $btnEmailTemplate.Location = New-Object System.Drawing.Point(810, 20)
        $btnEmailTemplate.Size = New-Object System.Drawing.Size(130, 40)
        $btnEmailTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        
        $btnManagement.Text = "MANAGEMENT"
        $btnManagement.Location = New-Object System.Drawing.Point(950, 20)
        $btnManagement.Size = New-Object System.Drawing.Size(130, 40)
        $btnManagement.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        
        $btnGuida.Location = New-Object System.Drawing.Point(1090, 20)
        $btnGuida.Size = New-Object System.Drawing.Size(100, 40)
        $btnGuida.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        
        $btnAutoTicketApp.Text = "AUTO TICKETAPP"
        $btnAutoTicketApp.Location = New-Object System.Drawing.Point(1200, 20)
        $btnAutoTicketApp.Size = New-Object System.Drawing.Size(140, 40)
        $btnAutoTicketApp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        
        $panelStatus.Location = New-Object System.Drawing.Point(($w - 220), 20)
        $panelStatus.Size = New-Object System.Drawing.Size(200, 40)
        $lblStatusDot.Font = New-Object System.Drawing.Font("Segoe UI", 14)
        $lblStatusText.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        
        $panelStats.Height = 100
        $panelControls.Height = 60
        $panelFilter.Height = 50
        $panelExcludeBar.Height = 50
        $panelMatch.Height = 50
        $panelLists.Height = 120
        
        # Liste exclusions e match rules
        $halfWidth = [int](($w - 50) / 2)
        $grpExclusions.Location = New-Object System.Drawing.Point(15, 5)
        $grpExclusions.Size = New-Object System.Drawing.Size($halfWidth, 110)
        $listExclusions.Size = New-Object System.Drawing.Size(($halfWidth - 20), 85)
        
        $grpMatchRules.Location = New-Object System.Drawing.Point(($halfWidth + 25), 5)
        $grpMatchRules.Size = New-Object System.Drawing.Size($halfWidth, 110)
        $listMatchRules.Size = New-Object System.Drawing.Size(($halfWidth - 20), 85)
        
        $grid.RowTemplate.Height = 32
        $grid.ColumnHeadersHeight = 40
        $grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 9)
    }
    
    $form.Refresh()
}

$cmbResolution.Add_SelectedIndexChanged({
    $selected = $cmbResolution.SelectedItem
    $preset = $script:resolutionPresets[$selected]
    
    if ($selected -eq "Auto") {
        # Auto: usa 90% dello schermo
        $screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
        $screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height
        $newWidth = [Math]::Max(950, [Math]::Min(1650, [int]($screenWidth * 0.90)))
        $newHeight = [Math]::Max(600, [Math]::Min(1000, [int]($screenHeight * 0.90)))
    } else {
        $newWidth = $preset.Width
        $newHeight = $preset.Height
    }
    
    $form.Size = New-Object System.Drawing.Size($newWidth, $newHeight)
    Adjust-LayoutToResolution -newWidth $newWidth -newHeight $newHeight
    $form.CenterToScreen()
})

# =================== EMAIL TEMPLATE BUTTON ===================
$btnEmailTemplate = New-Object System.Windows.Forms.Button
$btnEmailTemplate.Text = "SEND EMAIL"
$btnEmailTemplate.Location = New-Object System.Drawing.Point(810, 20)
$btnEmailTemplate.Size = New-Object System.Drawing.Size(130, 40)
$btnEmailTemplate.BackColor = $theme.bgElevated
$btnEmailTemplate.ForeColor = $theme.accent
$btnEmailTemplate.FlatStyle = "Flat"
$btnEmailTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnEmailTemplate.FlatAppearance.BorderColor = $theme.accent
$btnEmailTemplate.FlatAppearance.BorderSize = 2
$panelHeader.Controls.Add($btnEmailTemplate)

# =================== MANAGEMENT BUTTON ===================
$btnManagement = New-Object System.Windows.Forms.Button
$btnManagement.Text = "MANAGEMENT"
$btnManagement.Location = New-Object System.Drawing.Point(950, 20)
$btnManagement.Size = New-Object System.Drawing.Size(130, 40)
$btnManagement.BackColor = $theme.bgElevated
$btnManagement.ForeColor = $theme.attention
$btnManagement.FlatStyle = "Flat"
$btnManagement.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnManagement.FlatAppearance.BorderColor = $theme.attention
$btnManagement.FlatAppearance.BorderSize = 2
$panelHeader.Controls.Add($btnManagement)

# Lista dei link di management
$script:managementLinks = @(
    @{ Name = "NOC Operator - NOC Vademecum"; Url = "" },
    @{ Name = "OpManager"; Url = "https://10.41.114.17:8061/" },
    @{ Name = "MSP OpManager"; Url = "https://service.maticmind.it:8061/" },
    @{ Name = "MATICMIND S.p.A. (SIM III - CUSTOMER)"; Url = "https://maticmindspa.simcard.com/" }
)

# Funzione per mostrare il popup Management
function Show-ManagementLinks {
    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = "Management Links"
    $popup.Size = New-Object System.Drawing.Size(500, 360)
    $popup.StartPosition = "CenterParent"
    $popup.BackColor = $theme.bgPanel
    $popup.FormBorderStyle = "FixedDialog"
    $popup.MaximizeBox = $false
    $popup.MinimizeBox = $false
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "MANAGEMENT LINKS"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $theme.attention
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.AutoSize = $true
    $popup.Controls.Add($lblTitle)
    
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Click a link to open in browser:"
    $lblInfo.Location = New-Object System.Drawing.Point(20, 50)
    $lblInfo.Size = New-Object System.Drawing.Size(400, 20)
    $lblInfo.ForeColor = $theme.textSecond
    $popup.Controls.Add($lblInfo)
    
    $yPos = 85
    foreach ($link in $script:managementLinks) {
        $btnLink = New-Object System.Windows.Forms.Button
        $btnLink.Text = $link.Name
        $btnLink.Location = New-Object System.Drawing.Point(20, $yPos)
        $btnLink.Size = New-Object System.Drawing.Size(440, 40)
        $btnLink.BackColor = $theme.bgInput
        $btnLink.ForeColor = $theme.textPrimary
        $btnLink.FlatStyle = "Flat"
        $btnLink.Font = New-Object System.Drawing.Font("Segoe UI", 11)
        $btnLink.FlatAppearance.BorderColor = $theme.border
        $btnLink.TextAlign = "MiddleLeft"
        $btnLink.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
        $btnLink.Tag = $link.Url
        
        $btnLink.Add_Click({
            $url = $this.Tag
            if ($url -and $url -ne "") {
                Start-Process $url
            } else {
                [System.Windows.Forms.MessageBox]::Show("URL not configured for this link.", "No URL", "OK", "Information")
            }
        })
        
        $popup.Controls.Add($btnLink)
        $yPos += 50
    }
    
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "CLOSE"
    $btnClose.Size = New-Object System.Drawing.Size(100, 35)
    $btnClose.Location = New-Object System.Drawing.Point(190, $yPos + 10)
    $btnClose.BackColor = $theme.bgElevated
    $btnClose.ForeColor = $theme.textSecond
    $btnClose.FlatStyle = "Flat"
    $btnClose.FlatAppearance.BorderColor = $theme.border
    $btnClose.Add_Click({ $popup.Close() })
    $popup.Controls.Add($btnClose)
    
    [void]$popup.ShowDialog()
}

$btnManagement.Add_Click({
    Show-ManagementLinks
})

# =================== GUIDA BUTTON ===================
$btnGuida = New-Object System.Windows.Forms.Button
$btnGuida.Text = "GUIDA"
$btnGuida.Location = New-Object System.Drawing.Point(1090, 20)
$btnGuida.Size = New-Object System.Drawing.Size(100, 40)
$btnGuida.BackColor = $theme.bgElevated
$btnGuida.ForeColor = $theme.clear
$btnGuida.FlatStyle = "Flat"
$btnGuida.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnGuida.FlatAppearance.BorderColor = $theme.clear
$btnGuida.FlatAppearance.BorderSize = 2
$panelHeader.Controls.Add($btnGuida)

$btnGuida.Add_Click({
    $guidaForm = New-Object System.Windows.Forms.Form
    $guidaForm.Text = "Guida - Ticket Monitor"
    $guidaForm.Size = New-Object System.Drawing.Size(700, 650)
    $guidaForm.StartPosition = "CenterScreen"
    $guidaForm.BackColor = $theme.bgDeep
    $guidaForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $txtGuida = New-Object System.Windows.Forms.RichTextBox
    $txtGuida.Dock = "Fill"
    $txtGuida.BackColor = $theme.bgInput
    $txtGuida.ForeColor = $theme.textPrimary
    $txtGuida.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtGuida.ReadOnly = $true
    $txtGuida.BorderStyle = "None"
    $txtGuida.Padding = New-Object System.Windows.Forms.Padding(15)
    
    $guidaText = @"
========================================
       TICKET MONITOR - GUIDA UTENTE
========================================

PANORAMICA
----------
Ticket Monitor e' un dashboard per monitorare gli alert email da Outlook.
Legge automaticamente le email con severity CRITICAL, ATTENTION, TROUBLE e CLEAR
e le associa automaticamente tramite regole di matching.


COME INIZIARE
-------------
1. Seleziona il PROFILE (account Outlook) dal menu a tendina
2. Imposta la DATA e ORA di partenza per la scansione
3. Clicca START per avviare il monitoraggio
4. Le email verranno caricate e visualizzate nella griglia


FILTRI SEVERITY
---------------
- CRITICAL (rosso): Alert critici che richiedono attenzione immediata
- ATTENTION (giallo): Alert di attenzione
- TROUBLE (arancione): Alert di problemi
- CLEAR (verde): Email che indicano la risoluzione di un alert

Usa i checkbox per mostrare/nascondere ogni tipo di severity.


MATCH RULES (Regole di Associazione)
------------------------------------
Le Match Rules permettono di associare automaticamente gli alert ai CLEAR.

Per aggiungere una regola:
1. Seleziona il DOMAIN (dominio email) o "ALL DOMAINS"
2. Scegli il MATCH TYPE:
   - IP Address: Associa per indirizzo IP trovato nell'email
   - Hostname: Associa per nome host
   - Subject Pattern: Associa per pattern nel subject
   - First Word After CLEAR/CRITICAL: Associa per la prima parola dopo la severity
   - Custom Regex: Usa un'espressione regolare personalizzata
3. Clicca ADD RULE

Quando un CRITICAL viene associato a un CLEAR, viene mostrato come "CRITICAL -> CLEAR"
con testo barrato e colore giallo.


EXCLUSION FILTERS (Filtri di Esclusione)
----------------------------------------
Puoi escludere email che contengono determinate parole:
1. Seleziona il dominio o "ALL"
2. Inserisci la parola da escludere
3. Clicca ADD

Le email che contengono quella parola verranno nascoste.


VISUALIZZAZIONE DETTAGLI
------------------------
- Doppio click su una riga: Apre i dettagli dell'email
- Per alert risolti (CLEARED): Mostra confronto tra alert originale e CLEAR
- Tutto il contenuto e' selezionabile e copiabile
- Pulsante COPY ALL: Copia tutto negli appunti


PULSANTI HEADER
---------------
- SEND EMAIL: Apre un template email precompilato per inviare notifiche
- MANAGEMENT: Link rapidi a OpManager e altri tool di gestione
- GUIDA: Questa guida
- FOLDERS: Mostra il log delle cartelle scansionate (debug)


IMPOSTAZIONI
------------
- TEXT SIZE: Cambia la dimensione del testo (Small/Medium/Large)
- INTERVAL: Intervallo di refresh automatico (in secondi)
- CONTINUOUS: Se attivo, continua a scansionare automaticamente


STATISTICHE
-----------
- CRITICAL: Numero di alert critici attivi (non risolti)
- CLEARED: Numero di alert risolti
- TOTAL: Totale alert attivi (CRITICAL + ATTENTION + TROUBLE)


SUGGERIMENTI
------------
- Imposta le Match Rules per ogni dominio per avere associazioni precise
- Usa i filtri di esclusione per nascondere alert non rilevanti
- Il pulsante FOLDERS aiuta a diagnosticare problemi di scansione
- Le email vengono lette da TUTTE le cartelle Outlook


========================================
         Buon lavoro!
========================================
"@
    
    $txtGuida.Text = $guidaText
    $guidaForm.Controls.Add($txtGuida)
    
    $panelBottom = New-Object System.Windows.Forms.Panel
    $panelBottom.Dock = "Bottom"
    $panelBottom.Height = 50
    $panelBottom.BackColor = $theme.bgPanel
    $guidaForm.Controls.Add($panelBottom)
    
    $btnChiudi = New-Object System.Windows.Forms.Button
    $btnChiudi.Text = "CHIUDI"
    $btnChiudi.Size = New-Object System.Drawing.Size(120, 35)
    $btnChiudi.Location = New-Object System.Drawing.Point(280, 8)
    $btnChiudi.BackColor = $theme.accent
    $btnChiudi.ForeColor = $theme.bgDeep
    $btnChiudi.FlatStyle = "Flat"
    $btnChiudi.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnChiudi.FlatAppearance.BorderSize = 0
    $btnChiudi.Add_Click({ $guidaForm.Close() })
    $panelBottom.Controls.Add($btnChiudi)
    
    [void]$guidaForm.ShowDialog()
})

# =================== AUTO TICKETAPP BUTTON ===================
$btnAutoTicketApp = New-Object System.Windows.Forms.Button
$btnAutoTicketApp.Text = "AUTO TICKETAPP"
$btnAutoTicketApp.Location = New-Object System.Drawing.Point(1200, 20)
$btnAutoTicketApp.Size = New-Object System.Drawing.Size(140, 40)
$btnAutoTicketApp.BackColor = $theme.bgElevated
$btnAutoTicketApp.ForeColor = $theme.trouble
$btnAutoTicketApp.FlatStyle = "Flat"
$btnAutoTicketApp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnAutoTicketApp.FlatAppearance.BorderColor = $theme.trouble
$btnAutoTicketApp.FlatAppearance.BorderSize = 2
$panelHeader.Controls.Add($btnAutoTicketApp)

# Variabile per memorizzare il percorso trovato
$script:autoTicketAppPath = $null

$btnAutoTicketApp.Add_Click({
    try {
        $lblStatusText.Text = "Searching AutoTicketApp..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Se abbiamo gia' trovato il percorso, usa quello
        if ($script:autoTicketAppPath -and (Test-Path $script:autoTicketAppPath)) {
            Start-Process $script:autoTicketAppPath
            $lblStatusText.Text = "AutoTicketApp launched"
            return
        }
        
        # Cerca il programma in vari percorsi comuni
        $searchPaths = @(
            "$env:USERPROFILE\Desktop",
            "$env:USERPROFILE\Documents",
            "$env:USERPROFILE\Downloads",
            "$env:APPDATA",
            "$env:LOCALAPPDATA",
            "$env:ProgramFiles",
            "${env:ProgramFiles(x86)}",
            "C:\Program Files",
            "C:\Program Files (x86)",
            "C:\Users\$env:USERNAME",
            "C:\"
        )
        
        $foundPath = $null
        
        # Cerca file .exe con nome AutoTicketApp
        foreach ($searchPath in $searchPaths) {
            if (-not (Test-Path $searchPath)) { continue }
            
            $lblStatusText.Text = "Searching in $searchPath..."
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                # Cerca ricorsivamente ma con limite di profondita'
                $files = Get-ChildItem -Path $searchPath -Filter "*AutoTicketApp*.exe" -Recurse -Depth 5 -ErrorAction SilentlyContinue
                if ($files) {
                    $foundPath = $files[0].FullName
                    break
                }
            } catch {
                continue
            }
        }
        
        if ($foundPath) {
            $script:autoTicketAppPath = $foundPath
            Start-Process $foundPath
            $lblStatusText.Text = "AutoTicketApp launched"
        } else {
            # Non trovato, chiedi all'utente di cercarlo manualmente
            $result = [System.Windows.Forms.MessageBox]::Show(
                "AutoTicketApp non trovato automaticamente.`n`nVuoi cercarlo manualmente?",
                "AutoTicketApp Non Trovato",
                "YesNo",
                "Question"
            )
            
            if ($result -eq "Yes") {
                $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                $openFileDialog.Title = "Seleziona AutoTicketApp"
                $openFileDialog.Filter = "Executable (*.exe)|*.exe|All Files (*.*)|*.*"
                $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
                
                if ($openFileDialog.ShowDialog() -eq "OK") {
                    $script:autoTicketAppPath = $openFileDialog.FileName
                    Start-Process $openFileDialog.FileName
                    $lblStatusText.Text = "AutoTicketApp launched"
                } else {
                    $lblStatusText.Text = "Ready"
                }
            } else {
                $lblStatusText.Text = "Ready"
            }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Errore durante l'avvio di AutoTicketApp:`n$($_.Exception.Message)",
            "Errore",
            "OK",
            "Error"
        )
        $lblStatusText.Text = "Error"
    }
})

# Lista dei template email disponibili
$script:emailTemplates = @(
    @{ Name = "Acque"; Path = "" },
    @{ Name = "Bacardi"; Path = "" },
    @{ Name = "BCC"; Path = "" },
    @{ Name = "BPM"; Path = "" },
    @{ Name = "Cassa Forense - Grafana"; Path = "" },
    @{ Name = "CGT"; Path = "" },
    @{ Name = "Chiesi"; Path = "" },
    @{ Name = "Coesia"; Path = "" },
    @{ Name = "Coesia (alt)"; Path = "" },
    @{ Name = "Crea"; Path = "" },
    @{ Name = "Fercam"; Path = "" },
    @{ Name = "Granarolo"; Path = "" },
    @{ Name = "Maccaferri"; Path = "" },
    @{ Name = "Noovle"; Path = "" },
    @{ Name = "Regione Marche"; Path = "" },
    @{ Name = "SecurItalia"; Path = "" },
    @{ Name = "SicurItalia"; Path = "" },
    @{ Name = "Snaitech"; Path = "" },
    @{ Name = "Tosano"; Path = "" },
    @{ Name = "Unicredit"; Path = "" },
    @{ Name = "Wind3"; Path = "" }
)

# Funzione per cercare i template nella cartella dell'utente
function Find-EmailTemplates {
    $searchPaths = @(
        [Environment]::GetFolderPath("MyDocuments"),
        [Environment]::GetFolderPath("Desktop"),
        "$env:USERPROFILE\Templates",
        "$env:APPDATA\Microsoft\Templates",
        "$env:USERPROFILE\Downloads"
    )
    
    foreach ($template in $script:emailTemplates) {
        $found = $false
        foreach ($searchPath in $searchPaths) {
            if (-not (Test-Path $searchPath)) { continue }
            
            $files = Get-ChildItem -Path $searchPath -Filter "*.oft" -Recurse -ErrorAction SilentlyContinue
            foreach ($file in $files) {
                if ($file.Name -like "*$($template.Name)*") {
                    $template.Path = $file.FullName
                    $found = $true
                    break
                }
            }
            if ($found) { break }
        }
    }
}

# Funzione per mostrare il popup di selezione template
function Show-EmailTemplateSelector {
    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = "Select Email Template"
    $popup.Size = New-Object System.Drawing.Size(450, 550)
    $popup.StartPosition = "CenterParent"
    $popup.BackColor = $theme.bgPanel
    $popup.FormBorderStyle = "FixedDialog"
    $popup.MaximizeBox = $false
    $popup.MinimizeBox = $false
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "SELECT TEMPLATE TO SEND"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $theme.accent
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.AutoSize = $true
    $popup.Controls.Add($lblTitle)
    
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Double-click to open template, or select and choose action:"
    $lblInfo.Location = New-Object System.Drawing.Point(20, 45)
    $lblInfo.Size = New-Object System.Drawing.Size(400, 20)
    $lblInfo.ForeColor = $theme.textSecond
    $popup.Controls.Add($lblInfo)
    
    $listTemplates = New-Object System.Windows.Forms.ListBox
    $listTemplates.Location = New-Object System.Drawing.Point(20, 75)
    $listTemplates.Size = New-Object System.Drawing.Size(395, 340)
    $listTemplates.BackColor = $theme.bgInput
    $listTemplates.ForeColor = $theme.textPrimary
    $listTemplates.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $listTemplates.BorderStyle = "None"
    $listTemplates.ItemHeight = 28
    
    foreach ($template in $script:emailTemplates) {
        $status = if ($template.Path -and (Test-Path $template.Path)) { "[OK]" } else { "[?]" }
        [void]$listTemplates.Items.Add("$status  $($template.Name)")
    }
    
    $popup.Controls.Add($listTemplates)
    
    # Panel per i pulsanti
    $panelButtons = New-Object System.Windows.Forms.Panel
    $panelButtons.Location = New-Object System.Drawing.Point(20, 425)
    $panelButtons.Size = New-Object System.Drawing.Size(395, 80)
    $popup.Controls.Add($panelButtons)
    
    $btnSendDirect = New-Object System.Windows.Forms.Button
    $btnSendDirect.Text = "OPEN IN OUTLOOK"
    $btnSendDirect.Size = New-Object System.Drawing.Size(180, 35)
    $btnSendDirect.Location = New-Object System.Drawing.Point(0, 5)
    $btnSendDirect.BackColor = $theme.clear
    $btnSendDirect.ForeColor = $theme.bgDeep
    $btnSendDirect.FlatStyle = "Flat"
    $btnSendDirect.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnSendDirect.FlatAppearance.BorderSize = 0
    $panelButtons.Controls.Add($btnSendDirect)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "CANCEL"
    $btnCancel.Size = New-Object System.Drawing.Size(100, 35)
    $btnCancel.Location = New-Object System.Drawing.Point(190, 5)
    $btnCancel.BackColor = $theme.bgElevated
    $btnCancel.ForeColor = $theme.textSecond
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.FlatAppearance.BorderColor = $theme.border
    $panelButtons.Controls.Add($btnCancel)
    
    $lblPathInfo = New-Object System.Windows.Forms.Label
    $lblPathInfo.Text = "Template path will appear here when selected"
    $lblPathInfo.Location = New-Object System.Drawing.Point(0, 50)
    $lblPathInfo.Size = New-Object System.Drawing.Size(395, 25)
    $lblPathInfo.ForeColor = $theme.textMuted
    $lblPathInfo.Font = New-Object System.Drawing.Font("Consolas", 8)
    $panelButtons.Controls.Add($lblPathInfo)
    
    # Funzione per aprire il template
    $openTemplate = {
        param([bool]$DisplayOnly = $true)
        
        $idx = $listTemplates.SelectedIndex
        if ($idx -lt 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select a template first.", "No Selection", "OK", "Warning")
            return
        }
        
        $template = $script:emailTemplates[$idx]
        
        if (-not $template.Path -or -not (Test-Path $template.Path)) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Template file not found.`n`nWould you like to browse for it?",
                "Template Not Found",
                "YesNo",
                "Question"
            )
            
            if ($result -eq "Yes") {
                $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                $openFileDialog.Title = "Select Template: $($template.Name)"
                $openFileDialog.Filter = "Outlook Template (*.oft)|*.oft|All Files (*.*)|*.*"
                $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
                
                if ($openFileDialog.ShowDialog() -eq "OK") {
                    $template.Path = $openFileDialog.FileName
                    $listTemplates.Items[$idx] = "[OK]  $($template.Name)"
                } else {
                    return
                }
            } else {
                return
            }
        }
        
        try {
            # Verifica Outlook
            if (-not $script:outlook) {
                $script:outlook = New-Object -ComObject Outlook.Application
            }
            
            # Apri il template in Outlook
            $mail = $script:outlook.CreateItemFromTemplate($template.Path)
            $mail.Display($DisplayOnly)
            $popup.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Error opening template:`n$($_.Exception.Message)`n`nAssicurati che Outlook sia aperto.",
                "Error",
                "OK",
                "Error"
            )
        }
    }
    
    $listTemplates.Add_SelectedIndexChanged({
        $idx = $listTemplates.SelectedIndex
        if ($idx -ge 0) {
            $template = $script:emailTemplates[$idx]
            if ($template.Path) {
                $lblPathInfo.Text = $template.Path
            } else {
                $lblPathInfo.Text = "Path not set - will prompt to browse"
            }
        }
    })
    
    $listTemplates.Add_DoubleClick({
        & $openTemplate $true
    })
    
    $btnSendDirect.Add_Click({
        & $openTemplate $true
    })
    
    $btnCancel.Add_Click({
        $popup.Close()
    })
    
    [void]$popup.ShowDialog()
}

$btnEmailTemplate.Add_Click({
    # Verifica che Outlook sia connesso
    if (-not $script:outlook) {
        try {
            $script:outlook = New-Object -ComObject Outlook.Application
            $script:namespace = $script:outlook.GetNamespace("MAPI")
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Impossibile connettersi a Outlook.`nAssicurati che Outlook sia aperto e riprova.",
                "Errore Outlook",
                "OK",
                "Error"
            )
            return
        }
    }
    
    # Cerca i template alla prima apertura
    try {
        Find-EmailTemplates
    } catch {
        # Ignora errori nella ricerca template
    }
    
    Show-EmailTemplateSelector
})

# =================== STATS CARDS ===================
$panelStats = New-Object System.Windows.Forms.Panel
$panelStats.Dock = "Top"
$panelStats.Height = 100
$panelStats.BackColor = $theme.bgDeep
$form.Controls.Add($panelStats)

function New-StatCard {
    param($Parent, $X, $Label, $Color)
    
    $card = New-Object System.Windows.Forms.Panel
    $card.Size = New-Object System.Drawing.Size(160, 70)
    $card.Location = New-Object System.Drawing.Point($X, 15)
    $card.BackColor = $theme.bgPanel
    $Parent.Controls.Add($card)
    
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Label
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $lbl.ForeColor = $theme.textMuted
    $lbl.Location = New-Object System.Drawing.Point(15, 10)
    $lbl.AutoSize = $true
    $card.Controls.Add($lbl)
    
    $val = New-Object System.Windows.Forms.Label
    $val.Text = "0"
    $val.Font = New-Object System.Drawing.Font("Consolas", 24, [System.Drawing.FontStyle]::Bold)
    $val.ForeColor = $Color
    $val.Location = New-Object System.Drawing.Point(15, 28)
    $val.AutoSize = $true
    $card.Controls.Add($val)
    
    return $val
}

$lblStatCritical = New-StatCard -Parent $panelStats -X 25 -Label "CRITICAL" -Color $theme.critical
$lblStatClear = New-StatCard -Parent $panelStats -X 200 -Label "CLEARED" -Color $theme.clear
$lblStatTotal = New-StatCard -Parent $panelStats -X 375 -Label "TOTAL" -Color $theme.accent

# =================== PIE CHART INTEGRATO (a destra delle stats) ===================
$script:piePanel = New-Object System.Windows.Forms.Panel
$script:piePanel.Location = New-Object System.Drawing.Point(560, 5)
$script:piePanel.Size = New-Object System.Drawing.Size(90, 90)
$script:piePanel.BackColor = $theme.bgPanel
$panelStats.Controls.Add($script:piePanel)

# Variabili per il pie chart
$script:pieCritical = 0
$script:pieAttention = 0
$script:pieTrouble = 0

$script:piePanel.Add_Paint({
    param($sender, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
    $rect = New-Object System.Drawing.Rectangle(5, 5, 80, 80)
    $total = $script:pieCritical + $script:pieAttention + $script:pieTrouble
    
    if ($total -eq 0) {
        # Cerchio vuoto grigio
        $brush = New-Object System.Drawing.SolidBrush($theme.textMuted)
        $g.FillEllipse($brush, $rect)
        $brush.Dispose()
    } else {
        $startAngle = -90
        
        # CRITICAL (rosso)
        if ($script:pieCritical -gt 0) {
            $sweepAngle = ($script:pieCritical / $total) * 360
            $brush = New-Object System.Drawing.SolidBrush($theme.critical)
            $g.FillPie($brush, $rect, $startAngle, $sweepAngle)
            $brush.Dispose()
            $startAngle += $sweepAngle
        }
        
        # ATTENTION (giallo)
        if ($script:pieAttention -gt 0) {
            $sweepAngle = ($script:pieAttention / $total) * 360
            $brush = New-Object System.Drawing.SolidBrush($theme.attention)
            $g.FillPie($brush, $rect, $startAngle, $sweepAngle)
            $brush.Dispose()
            $startAngle += $sweepAngle
        }
        
        # TROUBLE (arancione)
        if ($script:pieTrouble -gt 0) {
            $sweepAngle = ($script:pieTrouble / $total) * 360
            $brush = New-Object System.Drawing.SolidBrush($theme.trouble)
            $g.FillPie($brush, $rect, $startAngle, $sweepAngle)
            $brush.Dispose()
        }
        
        # Cerchio centrale (buco della ciambella)
        $innerRect = New-Object System.Drawing.Rectangle(25, 25, 40, 40)
        $innerBrush = New-Object System.Drawing.SolidBrush($theme.bgPanel)
        $g.FillEllipse($innerBrush, $innerRect)
        $innerBrush.Dispose()
        
        # Numero al centro
        $textBrush = New-Object System.Drawing.SolidBrush($theme.textPrimary)
        $font = New-Object System.Drawing.Font("Consolas", 11, [System.Drawing.FontStyle]::Bold)
        $text = $total.ToString()
        $textSize = $g.MeasureString($text, $font)
        $x = (90 - $textSize.Width) / 2
        $y = (90 - $textSize.Height) / 2
        $g.DrawString($text, $font, $textBrush, $x, $y)
        $textBrush.Dispose()
        $font.Dispose()
    }
})

# Legenda mini accanto al pie
$script:lblPieLegend = New-Object System.Windows.Forms.Label
$script:lblPieLegend.Location = New-Object System.Drawing.Point(660, 15)
$script:lblPieLegend.Size = New-Object System.Drawing.Size(150, 70)
$script:lblPieLegend.Font = New-Object System.Drawing.Font("Consolas", 8)
$script:lblPieLegend.ForeColor = $theme.textSecond
$script:lblPieLegend.Text = "ACTIVE ALERTS`n- CRITICAL: 0`n- ATTENTION: 0`n- TROUBLE: 0"
$panelStats.Controls.Add($script:lblPieLegend)

# Funzione per aggiornare il pie chart
function Update-PieChart {
    $script:pieCritical = ($script:allEmails | Where-Object { $_.Sev -eq "CRITICAL" -and -not $_.IsCleared -and -not $_.IsEscalated }).Count
    $script:pieAttention = ($script:allEmails | Where-Object { $_.Sev -eq "ATTENTION" -and -not $_.IsCleared -and -not $_.IsEscalated }).Count
    $script:pieTrouble = ($script:allEmails | Where-Object { $_.Sev -eq "TROUBLE" -and -not $_.IsCleared -and -not $_.IsEscalated }).Count
    
    $script:lblPieLegend.Text = "ACTIVE ALERTS`n- CRITICAL: $($script:pieCritical)`n- ATTENTION: $($script:pieAttention)`n- TROUBLE: $($script:pieTrouble)"
    
    $script:piePanel.Refresh()
}

# =================== CONTROLS BAR ===================
$panelControls = New-Object System.Windows.Forms.Panel
$panelControls.Dock = "Top"
$panelControls.Height = 60
$panelControls.BackColor = $theme.bgPanel
$form.Controls.Add($panelControls)

$lblProfile = New-Object System.Windows.Forms.Label
$lblProfile.Text = "PROFILE"
$lblProfile.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblProfile.ForeColor = $theme.textSecond
$lblProfile.Location = New-Object System.Drawing.Point(25, 8)
$lblProfile.AutoSize = $true
$panelControls.Controls.Add($lblProfile)

$cmbProfile = New-Object System.Windows.Forms.ComboBox
$cmbProfile.Location = New-Object System.Drawing.Point(25, 26)
$cmbProfile.Size = New-Object System.Drawing.Size(180, 25)
$cmbProfile.DropDownStyle = "DropDownList"
$cmbProfile.BackColor = $theme.bgInput
$cmbProfile.ForeColor = $theme.textPrimary
$cmbProfile.FlatStyle = "Flat"
foreach ($store in $script:namespace.Stores) { 
    [void]$cmbProfile.Items.Add($store.DisplayName) 
}
if ($cmbProfile.Items.Count -gt 0) { $cmbProfile.SelectedIndex = 0 }
$panelControls.Controls.Add($cmbProfile)

$lblDate = New-Object System.Windows.Forms.Label
$lblDate.Text = "FROM"
$lblDate.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblDate.ForeColor = $theme.textSecond
$lblDate.Location = New-Object System.Drawing.Point(220, 8)
$lblDate.AutoSize = $true
$panelControls.Controls.Add($lblDate)

$dtpDate = New-Object System.Windows.Forms.DateTimePicker
$dtpDate.Location = New-Object System.Drawing.Point(220, 26)
$dtpDate.Size = New-Object System.Drawing.Size(120, 25)
$dtpDate.Format = "Short"
$dtpDate.Value = (Get-Date).Date
$panelControls.Controls.Add($dtpDate)

$nudHour = New-Object System.Windows.Forms.NumericUpDown
$nudHour.Location = New-Object System.Drawing.Point(350, 26)
$nudHour.Size = New-Object System.Drawing.Size(45, 25)
$nudHour.Minimum = 0; $nudHour.Maximum = 23; $nudHour.Value = 0
$nudHour.BackColor = $theme.bgInput; $nudHour.ForeColor = $theme.textPrimary
$panelControls.Controls.Add($nudHour)

$lblColon = New-Object System.Windows.Forms.Label
$lblColon.Text = ":"
$lblColon.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$lblColon.ForeColor = $theme.accent
$lblColon.Location = New-Object System.Drawing.Point(397, 26)
$lblColon.AutoSize = $true
$panelControls.Controls.Add($lblColon)

$nudMinute = New-Object System.Windows.Forms.NumericUpDown
$nudMinute.Location = New-Object System.Drawing.Point(410, 26)
$nudMinute.Size = New-Object System.Drawing.Size(45, 25)
$nudMinute.Minimum = 0; $nudMinute.Maximum = 59; $nudMinute.Value = 0
$nudMinute.BackColor = $theme.bgInput; $nudMinute.ForeColor = $theme.textPrimary
$panelControls.Controls.Add($nudMinute)

$lblInterval = New-Object System.Windows.Forms.Label
$lblInterval.Text = "INTERVAL"
$lblInterval.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblInterval.ForeColor = $theme.textSecond
$lblInterval.Location = New-Object System.Drawing.Point(470, 8)
$lblInterval.AutoSize = $true
$panelControls.Controls.Add($lblInterval)

$nudInterval = New-Object System.Windows.Forms.NumericUpDown
$nudInterval.Location = New-Object System.Drawing.Point(470, 26)
$nudInterval.Size = New-Object System.Drawing.Size(50, 25)
$nudInterval.Minimum = 5; $nudInterval.Maximum = 300; $nudInterval.Value = 30
$nudInterval.BackColor = $theme.bgInput; $nudInterval.ForeColor = $theme.textPrimary
$panelControls.Controls.Add($nudInterval)

$chkContinuous = New-Object System.Windows.Forms.CheckBox
$chkContinuous.Text = "Continuous"
$chkContinuous.Location = New-Object System.Drawing.Point(535, 26)
$chkContinuous.ForeColor = $theme.textPrimary
$chkContinuous.AutoSize = $true
$chkContinuous.Checked = $true
$panelControls.Controls.Add($chkContinuous)

$chkSound = New-Object System.Windows.Forms.CheckBox
$chkSound.Text = "Sound"
$chkSound.Location = New-Object System.Drawing.Point(535, 8)
$chkSound.ForeColor = $theme.attention
$chkSound.AutoSize = $true
$chkSound.Checked = $true
$panelControls.Controls.Add($chkSound)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "START"
$btnStart.Location = New-Object System.Drawing.Point(650, 20)
$btnStart.Size = New-Object System.Drawing.Size(100, 32)
$btnStart.BackColor = $theme.accent
$btnStart.ForeColor = $theme.bgDeep
$btnStart.FlatStyle = "Flat"
$btnStart.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnStart.FlatAppearance.BorderSize = 0
$panelControls.Controls.Add($btnStart)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "CLEAR"
$btnClear.Location = New-Object System.Drawing.Point(760, 20)
$btnClear.Size = New-Object System.Drawing.Size(70, 32)
$btnClear.BackColor = $theme.bgElevated
$btnClear.ForeColor = $theme.textSecond
$btnClear.FlatStyle = "Flat"
$btnClear.FlatAppearance.BorderColor = $theme.border
$panelControls.Controls.Add($btnClear)

# =================== DEBUG FOLDERS BUTTON ===================
$btnDebugFolders = New-Object System.Windows.Forms.Button
$btnDebugFolders.Text = "FOLDERS"
$btnDebugFolders.Location = New-Object System.Drawing.Point(840, 20)
$btnDebugFolders.Size = New-Object System.Drawing.Size(80, 32)
$btnDebugFolders.BackColor = $theme.bgElevated
$btnDebugFolders.ForeColor = $theme.attention
$btnDebugFolders.FlatStyle = "Flat"
$btnDebugFolders.FlatAppearance.BorderColor = $theme.attention
$btnDebugFolders.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$panelControls.Controls.Add($btnDebugFolders)

# =================== DELETE BUTTON (Pulisce tutto e libera memoria) ===================
$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "DELETE"
$btnDelete.Location = New-Object System.Drawing.Point(930, 20)
$btnDelete.Size = New-Object System.Drawing.Size(80, 32)
$btnDelete.BackColor = [System.Drawing.Color]::FromArgb(80, 30, 30)
$btnDelete.ForeColor = $theme.critical
$btnDelete.FlatStyle = "Flat"
$btnDelete.FlatAppearance.BorderColor = $theme.critical
$btnDelete.FlatAppearance.BorderSize = 2
$btnDelete.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$panelControls.Controls.Add($btnDelete)

$btnDelete.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Vuoi cancellare TUTTE le email dalla vista e liberare la memoria?`n`nQuesto NON elimina le email da Outlook, solo dalla visualizzazione.",
        "Conferma DELETE",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq "Yes") {
        # Ferma il timer se attivo
        $wasRunning = $script:isRunning
        if ($script:isRunning) {
            $script:isRunning = $false
            $script:timer.Stop()
        }
        
        # Pulisci tutto
        $grid.Rows.Clear()
        $script:allEmails.Clear()
        $script:mailMap.Clear()
        $script:alertClearMap.Clear()
        $script:processedIds.Clear()
        $script:folderScanLog.Clear()
        
        # Reset contatori
        $script:criticalCount = 0
        $script:clearedCount = 0
        $script:totalCount = 0
        Update-Stats
        
        # Forza garbage collection per liberare memoria
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        
        # Aggiorna UI
        $btnStart.Text = "START"
        $btnStart.BackColor = $theme.accent
        $lblStatusText.Text = "Cleared"
        $lblStatusDot.ForeColor = $theme.textMuted
        
        [System.Windows.Forms.MessageBox]::Show(
            "Memoria liberata!`n`nPuoi cliccare START per ricominciare.",
            "DELETE Completato",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

$btnDebugFolders.Add_Click({
    try {
        if ($script:folderScanLog.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No scan has been performed yet.`nClick START first.", "Folder Scan Log", "OK", "Information") | Out-Null
            return
        }
        
        $debugForm = New-Object System.Windows.Forms.Form
        $debugForm.Text = "Folder Scan Log - Debug"
        $debugForm.Size = New-Object System.Drawing.Size(800, 600)
        $debugForm.StartPosition = "CenterScreen"
        $debugForm.BackColor = $theme.bgDeep
        
        $txtLog = New-Object System.Windows.Forms.TextBox
        $txtLog.Multiline = $true
        $txtLog.ScrollBars = "Both"
        $txtLog.Dock = "Fill"
        $txtLog.BackColor = $theme.bgInput
        $txtLog.ForeColor = $theme.textPrimary
        $txtLog.Font = New-Object System.Drawing.Font("Consolas", 10)
        $txtLog.ReadOnly = $true
        $txtLog.WordWrap = $false
        
        $logText = "=== FOLDER SCAN LOG ===`r`n`r`n"
        $logText += "Scanned: " + (Get-Date).ToString("dd/MM/yyyy HH:mm:ss") + "`r`n"
        $logText += "Profile: " + $cmbProfile.SelectedItem + "`r`n"
        $logText += "`r`n--- RESULTS ---`r`n`r`n"
        
        foreach ($line in $script:folderScanLog) {
            $logText += $line + "`r`n"
        }
        
        $logText += "`r`n--- SUMMARY ---`r`n"
        $logText += "Total CRITICAL in memory: " + @($script:allEmails | Where-Object { $_.Sev -eq "CRITICAL" }).Count + "`r`n"
        $logText += "Total CLEAR in memory: " + @($script:allEmails | Where-Object { $_.Sev -eq "CLEAR" }).Count + "`r`n"
        
        $txtLog.Text = $logText
        $debugForm.Controls.Add($txtLog)
        
        [void]$debugForm.ShowDialog()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Errore: $($_.Exception.Message)", "Errore", "OK", "Error")
    }
})

# =================== FILTER BAR ===================
$panelFilter = New-Object System.Windows.Forms.Panel
$panelFilter.Dock = "Top"
$panelFilter.Height = 50
$panelFilter.BackColor = $theme.bgCard
$form.Controls.Add($panelFilter)

$lblSevFilter = New-Object System.Windows.Forms.Label
$lblSevFilter.Text = "SEVERITY"
$lblSevFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblSevFilter.ForeColor = $theme.textMuted
$lblSevFilter.Location = New-Object System.Drawing.Point(25, 16)
$lblSevFilter.AutoSize = $true
$panelFilter.Controls.Add($lblSevFilter)

$chkCritical = New-Object System.Windows.Forms.CheckBox
$chkCritical.Text = "CRITICAL"
$chkCritical.ForeColor = $theme.critical
$chkCritical.Location = New-Object System.Drawing.Point(95, 14)
$chkCritical.AutoSize = $true
$chkCritical.Checked = $true
$panelFilter.Controls.Add($chkCritical)

$chkAttention = New-Object System.Windows.Forms.CheckBox
$chkAttention.Text = "ATTENTION"
$chkAttention.ForeColor = $theme.attention
$chkAttention.Location = New-Object System.Drawing.Point(195, 14)
$chkAttention.AutoSize = $true
$chkAttention.Checked = $true
$panelFilter.Controls.Add($chkAttention)

$chkTrouble = New-Object System.Windows.Forms.CheckBox
$chkTrouble.Text = "TROUBLE"
$chkTrouble.ForeColor = $theme.trouble
$chkTrouble.Location = New-Object System.Drawing.Point(305, 14)
$chkTrouble.AutoSize = $true
$chkTrouble.Checked = $true
$panelFilter.Controls.Add($chkTrouble)

$chkClearFilter = New-Object System.Windows.Forms.CheckBox
$chkClearFilter.Text = "CLEAR"
$chkClearFilter.ForeColor = $theme.clear
$chkClearFilter.Location = New-Object System.Drawing.Point(405, 14)
$chkClearFilter.AutoSize = $true
$chkClearFilter.Checked = $true
$panelFilter.Controls.Add($chkClearFilter)

$lblDomFilter = New-Object System.Windows.Forms.Label
$lblDomFilter.Text = "DOMAINS"
$lblDomFilter.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblDomFilter.ForeColor = $theme.textMuted
$lblDomFilter.Location = New-Object System.Drawing.Point(500, 16)
$lblDomFilter.AutoSize = $true
$panelFilter.Controls.Add($lblDomFilter)

$btnSelectDomains = New-Object System.Windows.Forms.Button
$btnSelectDomains.Text = "ALL DOMAINS"
$btnSelectDomains.Location = New-Object System.Drawing.Point(570, 10)
$btnSelectDomains.Size = New-Object System.Drawing.Size(180, 28)
$btnSelectDomains.BackColor = $theme.bgInput
$btnSelectDomains.ForeColor = $theme.textPrimary
$btnSelectDomains.FlatStyle = "Flat"
$btnSelectDomains.FlatAppearance.BorderColor = $theme.border
$btnSelectDomains.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$btnSelectDomains.TextAlign = "MiddleLeft"
$panelFilter.Controls.Add($btnSelectDomains)

# Lista domini selezionati (vuota = tutti)
$script:selectedDomains = [System.Collections.ArrayList]::new()

# Funzione per mostrare popup selezione domini
function Show-DomainSelector {
    $popup = New-Object System.Windows.Forms.Form
    $popup.Text = "Select Domains"
    $popup.Size = New-Object System.Drawing.Size(300, 400)
    $popup.StartPosition = "CenterParent"
    $popup.BackColor = $theme.bgPanel
    $popup.FormBorderStyle = "FixedDialog"
    $popup.MaximizeBox = $false
    $popup.MinimizeBox = $false
    
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Select domains to monitor (none = ALL):"
    $lblInfo.Location = New-Object System.Drawing.Point(10, 10)
    $lblInfo.Size = New-Object System.Drawing.Size(270, 20)
    $lblInfo.ForeColor = $theme.textPrimary
    $popup.Controls.Add($lblInfo)
    
    $checkedList = New-Object System.Windows.Forms.CheckedListBox
    $checkedList.Location = New-Object System.Drawing.Point(10, 35)
    $checkedList.Size = New-Object System.Drawing.Size(265, 260)
    $checkedList.BackColor = $theme.bgInput
    $checkedList.ForeColor = $theme.textPrimary
    $checkedList.BorderStyle = "FixedSingle"
    $checkedList.CheckOnClick = $true
    
    # Popola con tutti i domini trovati
    foreach ($domain in $script:allDomains.Keys | Sort-Object) {
        [void]$checkedList.Items.Add($domain)
        # Seleziona se era già selezionato
        if ($script:selectedDomains.Contains($domain)) {
            $idx = $checkedList.Items.IndexOf($domain)
            $checkedList.SetItemChecked($idx, $true)
        }
    }
    $popup.Controls.Add($checkedList)
    
    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = "Select All"
    $btnAll.Location = New-Object System.Drawing.Point(10, 305)
    $btnAll.Size = New-Object System.Drawing.Size(80, 28)
    $btnAll.BackColor = $theme.bgElevated
    $btnAll.ForeColor = $theme.textPrimary
    $btnAll.FlatStyle = "Flat"
    $btnAll.Add_Click({
        for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
            $checkedList.SetItemChecked($i, $true)
        }
    })
    $popup.Controls.Add($btnAll)
    
    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = "Clear All"
    $btnNone.Location = New-Object System.Drawing.Point(100, 305)
    $btnNone.Size = New-Object System.Drawing.Size(80, 28)
    $btnNone.BackColor = $theme.bgElevated
    $btnNone.ForeColor = $theme.textPrimary
    $btnNone.FlatStyle = "Flat"
    $btnNone.Add_Click({
        for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
            $checkedList.SetItemChecked($i, $false)
        }
    })
    $popup.Controls.Add($btnNone)
    
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(195, 305)
    $btnOK.Size = New-Object System.Drawing.Size(80, 28)
    $btnOK.BackColor = $theme.accent
    $btnOK.ForeColor = $theme.bgDeep
    $btnOK.FlatStyle = "Flat"
    $btnOK.Add_Click({
        $script:selectedDomains.Clear()
        foreach ($item in $checkedList.CheckedItems) {
            [void]$script:selectedDomains.Add($item)
        }
        # Aggiorna testo bottone
        if ($script:selectedDomains.Count -eq 0) {
            $btnSelectDomains.Text = "ALL DOMAINS"
        } elseif ($script:selectedDomains.Count -eq 1) {
            $btnSelectDomains.Text = $script:selectedDomains[0]
        } else {
            $btnSelectDomains.Text = "$($script:selectedDomains.Count) domains"
        }
        Apply-Filters
        $popup.Close()
    })
    $popup.Controls.Add($btnOK)
    
    [void]$popup.ShowDialog()
}

$btnSelectDomains.Add_Click({ Show-DomainSelector })

# =================== EXCLUSION FILTER BAR ===================
$panelExcludeBar = New-Object System.Windows.Forms.Panel
$panelExcludeBar.Dock = "Top"
$panelExcludeBar.Height = 50
$panelExcludeBar.BackColor = $theme.bgElevated
$form.Controls.Add($panelExcludeBar)

$lblExcludeTitle = New-Object System.Windows.Forms.Label
$lblExcludeTitle.Text = "EXCLUSION FILTER"
$lblExcludeTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblExcludeTitle.ForeColor = $theme.critical
$lblExcludeTitle.Location = New-Object System.Drawing.Point(25, 16)
$lblExcludeTitle.AutoSize = $true
$panelExcludeBar.Controls.Add($lblExcludeTitle)

$lblExcludeDom = New-Object System.Windows.Forms.Label
$lblExcludeDom.Text = "Domain:"
$lblExcludeDom.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblExcludeDom.ForeColor = $theme.textSecond
$lblExcludeDom.Location = New-Object System.Drawing.Point(150, 16)
$lblExcludeDom.AutoSize = $true
$panelExcludeBar.Controls.Add($lblExcludeDom)

$cmbExcludeDomain = New-Object System.Windows.Forms.ComboBox
$cmbExcludeDomain.Location = New-Object System.Drawing.Point(205, 12)
$cmbExcludeDomain.Size = New-Object System.Drawing.Size(130, 25)
$cmbExcludeDomain.DropDownStyle = "DropDownList"
$cmbExcludeDomain.BackColor = $theme.bgInput
$cmbExcludeDomain.ForeColor = $theme.textPrimary
[void]$cmbExcludeDomain.Items.Add("-- ALL --")
$cmbExcludeDomain.SelectedIndex = 0
$panelExcludeBar.Controls.Add($cmbExcludeDomain)

$lblExcludeWord = New-Object System.Windows.Forms.Label
$lblExcludeWord.Text = "Word:"
$lblExcludeWord.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblExcludeWord.ForeColor = $theme.textSecond
$lblExcludeWord.Location = New-Object System.Drawing.Point(345, 16)
$lblExcludeWord.AutoSize = $true
$panelExcludeBar.Controls.Add($lblExcludeWord)

$txtExcludeWord = New-Object System.Windows.Forms.TextBox
$txtExcludeWord.Location = New-Object System.Drawing.Point(385, 12)
$txtExcludeWord.Size = New-Object System.Drawing.Size(150, 25)
$txtExcludeWord.BackColor = $theme.bgInput
$txtExcludeWord.ForeColor = $theme.textPrimary
$txtExcludeWord.BorderStyle = "FixedSingle"
$panelExcludeBar.Controls.Add($txtExcludeWord)

$btnAddExclusion = New-Object System.Windows.Forms.Button
$btnAddExclusion.Text = "ADD"
$btnAddExclusion.Location = New-Object System.Drawing.Point(545, 10)
$btnAddExclusion.Size = New-Object System.Drawing.Size(60, 28)
$btnAddExclusion.BackColor = $theme.critical
$btnAddExclusion.ForeColor = [System.Drawing.Color]::White
$btnAddExclusion.FlatStyle = "Flat"
$btnAddExclusion.FlatAppearance.BorderSize = 0
$btnAddExclusion.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$panelExcludeBar.Controls.Add($btnAddExclusion)

# =================== MATCH RULES BAR ===================
$panelMatch = New-Object System.Windows.Forms.Panel
$panelMatch.Dock = "Top"
$panelMatch.Height = 50
$panelMatch.BackColor = $theme.bgCard
$form.Controls.Add($panelMatch)

$lblMatchTitle = New-Object System.Windows.Forms.Label
$lblMatchTitle.Text = "MATCH RULES"
$lblMatchTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblMatchTitle.ForeColor = $theme.accent
$lblMatchTitle.Location = New-Object System.Drawing.Point(25, 16)
$lblMatchTitle.AutoSize = $true
$panelMatch.Controls.Add($lblMatchTitle)

$lblMatchDomain = New-Object System.Windows.Forms.Label
$lblMatchDomain.Text = "Domain:"
$lblMatchDomain.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblMatchDomain.ForeColor = $theme.textSecond
$lblMatchDomain.Location = New-Object System.Drawing.Point(130, 16)
$lblMatchDomain.AutoSize = $true
$panelMatch.Controls.Add($lblMatchDomain)

$cmbMatchDomain = New-Object System.Windows.Forms.ComboBox
$cmbMatchDomain.Location = New-Object System.Drawing.Point(185, 12)
$cmbMatchDomain.Size = New-Object System.Drawing.Size(150, 25)
$cmbMatchDomain.DropDownStyle = "DropDownList"
$cmbMatchDomain.BackColor = $theme.bgInput
$cmbMatchDomain.ForeColor = $theme.textPrimary
[void]$cmbMatchDomain.Items.Add("-- ALL DOMAINS --")
$cmbMatchDomain.SelectedIndex = 0
$panelMatch.Controls.Add($cmbMatchDomain)

$lblMatchBy = New-Object System.Windows.Forms.Label
$lblMatchBy.Text = "Match by:"
$lblMatchBy.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblMatchBy.ForeColor = $theme.textSecond
$lblMatchBy.Location = New-Object System.Drawing.Point(350, 16)
$lblMatchBy.AutoSize = $true
$panelMatch.Controls.Add($lblMatchBy)

$cmbMatchType = New-Object System.Windows.Forms.ComboBox
$cmbMatchType.Location = New-Object System.Drawing.Point(415, 12)
$cmbMatchType.Size = New-Object System.Drawing.Size(180, 25)
$cmbMatchType.DropDownStyle = "DropDownList"
$cmbMatchType.BackColor = $theme.bgInput
$cmbMatchType.ForeColor = $theme.textPrimary
foreach ($mt in $script:matchFieldTypes) { [void]$cmbMatchType.Items.Add($mt) }
$cmbMatchType.SelectedIndex = 0
$panelMatch.Controls.Add($cmbMatchType)

$lblMatchSeverity = New-Object System.Windows.Forms.Label
$lblMatchSeverity.Text = "For:"
$lblMatchSeverity.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblMatchSeverity.ForeColor = $theme.textSecond
$lblMatchSeverity.Location = New-Object System.Drawing.Point(605, 16)
$lblMatchSeverity.AutoSize = $true
$panelMatch.Controls.Add($lblMatchSeverity)

$cmbMatchSeverity = New-Object System.Windows.Forms.ComboBox
$cmbMatchSeverity.Location = New-Object System.Drawing.Point(635, 12)
$cmbMatchSeverity.Size = New-Object System.Drawing.Size(120, 25)
$cmbMatchSeverity.DropDownStyle = "DropDownList"
$cmbMatchSeverity.BackColor = $theme.bgInput
$cmbMatchSeverity.ForeColor = $theme.textPrimary
[void]$cmbMatchSeverity.Items.Add("ALL")
[void]$cmbMatchSeverity.Items.Add("CRITICAL")
[void]$cmbMatchSeverity.Items.Add("ATTENTION")
[void]$cmbMatchSeverity.Items.Add("TROUBLE")
$cmbMatchSeverity.SelectedIndex = 0
$panelMatch.Controls.Add($cmbMatchSeverity)

$lblCustomPattern = New-Object System.Windows.Forms.Label
$lblCustomPattern.Text = "Pattern:"
$lblCustomPattern.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblCustomPattern.ForeColor = $theme.textSecond
$lblCustomPattern.Location = New-Object System.Drawing.Point(765, 16)
$lblCustomPattern.AutoSize = $true
$lblCustomPattern.Visible = $false
$panelMatch.Controls.Add($lblCustomPattern)

$txtCustomPattern = New-Object System.Windows.Forms.TextBox
$txtCustomPattern.Location = New-Object System.Drawing.Point(815, 12)
$txtCustomPattern.Size = New-Object System.Drawing.Size(120, 25)
$txtCustomPattern.BackColor = $theme.bgInput
$txtCustomPattern.ForeColor = $theme.textPrimary
$txtCustomPattern.BorderStyle = "FixedSingle"
$txtCustomPattern.Visible = $false
$panelMatch.Controls.Add($txtCustomPattern)

$btnAddMatchRule = New-Object System.Windows.Forms.Button
$btnAddMatchRule.Text = "ADD"
$btnAddMatchRule.Location = New-Object System.Drawing.Point(950, 10)
$btnAddMatchRule.Size = New-Object System.Drawing.Size(60, 28)
$btnAddMatchRule.BackColor = $theme.clear
$btnAddMatchRule.ForeColor = $theme.bgDeep
$btnAddMatchRule.FlatStyle = "Flat"
$btnAddMatchRule.FlatAppearance.BorderSize = 0
$btnAddMatchRule.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$panelMatch.Controls.Add($btnAddMatchRule)

$cmbMatchType.Add_SelectedIndexChanged({
    $isCustom = $cmbMatchType.SelectedItem -eq "Custom Regex"
    $lblCustomPattern.Visible = $isCustom
    $txtCustomPattern.Visible = $isCustom
})

# =================== LISTS PANEL (EXCLUSIONS + MATCH RULES) ===================
$panelLists = New-Object System.Windows.Forms.Panel
$panelLists.Dock = "Top"
$panelLists.Height = 120
$panelLists.BackColor = $theme.bgMain
$form.Controls.Add($panelLists)

# Lista Exclusioni
$grpExclusions = New-Object System.Windows.Forms.GroupBox
$grpExclusions.Text = "Active Exclusions (click to remove)"
$grpExclusions.ForeColor = $theme.critical
$grpExclusions.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$grpExclusions.Location = New-Object System.Drawing.Point(15, 5)
$grpExclusions.Size = New-Object System.Drawing.Size(780, 110)
$panelLists.Controls.Add($grpExclusions)

$listExclusions = New-Object System.Windows.Forms.ListBox
$listExclusions.Location = New-Object System.Drawing.Point(10, 18)
$listExclusions.Size = New-Object System.Drawing.Size(760, 85)
$listExclusions.BackColor = $theme.bgInput
$listExclusions.ForeColor = $theme.textPrimary
$listExclusions.Font = New-Object System.Drawing.Font("Consolas", 9)
$listExclusions.BorderStyle = "None"
$listExclusions.SelectionMode = "One"
$grpExclusions.Controls.Add($listExclusions)

# Lista Match Rules
$grpMatchRules = New-Object System.Windows.Forms.GroupBox
$grpMatchRules.Text = "Active Match Rules (click to remove)"
$grpMatchRules.ForeColor = $theme.accent
$grpMatchRules.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$grpMatchRules.Location = New-Object System.Drawing.Point(810, 5)
$grpMatchRules.Size = New-Object System.Drawing.Size(780, 110)
$panelLists.Controls.Add($grpMatchRules)

$listMatchRules = New-Object System.Windows.Forms.ListBox
$listMatchRules.Location = New-Object System.Drawing.Point(10, 18)
$listMatchRules.Size = New-Object System.Drawing.Size(760, 85)
$listMatchRules.BackColor = $theme.bgInput
$listMatchRules.ForeColor = $theme.textPrimary
$listMatchRules.Font = New-Object System.Drawing.Font("Consolas", 9)
$listMatchRules.BorderStyle = "None"
$listMatchRules.SelectionMode = "One"
$grpMatchRules.Controls.Add($listMatchRules)

# =================== FUNZIONI UPDATE LISTE ===================
function Update-ExclusionsList {
    $listExclusions.Items.Clear()
    foreach ($f in $script:exclusionFilters) {
        $listExclusions.Items.Add("[$($f.Domain)]  ->  '$($f.Word)'")
    }
}

function Update-MatchRulesList {
    $listMatchRules.Items.Clear()
    if ($script:matchRules.Count -eq 0) {
        $listMatchRules.Items.Add("(NO RULES - emails will NOT be matched!)")
    } else {
        foreach ($r in $script:matchRules) {
            $sevText = if ($r.Severity -and $r.Severity -ne "ALL") { $r.Severity } else { "ALL" }
            $ruleText = "[$($r.Domain)] [$sevText]  ->  $($r.MatchType)"
            if ($r.CustomPattern) { $ruleText += "  Pattern: '$($r.CustomPattern)'" }
            $listMatchRules.Items.Add($ruleText)
        }
    }
}

Update-MatchRulesList

# Click per rimuovere exclusion
$listExclusions.Add_Click({
    $idx = $listExclusions.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $script:exclusionFilters.Count) {
        $filter = $script:exclusionFilters[$idx]
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Rimuovere il filtro?`n`nDomain: $($filter.Domain)`nWord: $($filter.Word)", 
            "Conferma Rimozione", 
            "YesNo", 
            "Question"
        )
        if ($result -eq "Yes") {
            $script:exclusionFilters.RemoveAt($idx)
            Update-ExclusionsList
            Apply-Filters
        }
    }
})

# Click per rimuovere match rule
$listMatchRules.Add_Click({
    $idx = $listMatchRules.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $script:matchRules.Count) {
        $rule = $script:matchRules[$idx]
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Rimuovere la regola?`n`nDomain: $($rule.Domain)`nMatch Type: $($rule.MatchType)", 
            "Conferma Rimozione", 
            "YesNo", 
            "Question"
        )
        if ($result -eq "Yes") {
            $script:matchRules.RemoveAt($idx)
            Update-MatchRulesList
            Refresh-GridWithNewRules
        }
    }
})

# =================== UPDATE STATS ===================
function Update-Stats {
    $critical = 0; $attention = 0; $trouble = 0; $cleared = 0
    
    foreach ($email in $script:allEmails) {
        if (Test-EmailExcluded -Email $email) { continue }
        
        if ($email.Sev -eq "CLEAR" -and $email.MatchedClearId) { continue }
        
        # Non contare email escalate
        if ($email.IsEscalated) { continue }
        
        switch ($email.Sev) {
            "CRITICAL" { 
                if ($email.IsCleared) { $cleared++ } 
                else { $critical++ } 
            }
            "ATTENTION" { 
                if ($email.IsCleared) { $cleared++ } 
                else { $attention++ } 
            }
            "TROUBLE" { 
                if ($email.IsCleared) { $cleared++ } 
                else { $trouble++ } 
            }
            "CLEAR" { 
                $cleared++ 
            }
        }
    }
    
    $totalActive = $critical + $attention + $trouble
    $lblStatCritical.Text = $critical.ToString()
    $lblStatClear.Text = $cleared.ToString()
    $lblStatTotal.Text = $totalActive.ToString()
    
    # Aggiorna il grafico a torta
    Update-PieChart
}

# =================== APPLY FILTERS ===================
function Apply-Filters {
    foreach ($row in $grid.Rows) {
        if ($row.IsNewRow) { continue }
        
        $sev = $row.Cells["Sev"].Value
        $dom = $row.Cells["Domain"].Value
        $show = $true
        
        $baseSev = $sev -replace " -> CLEAR", ""
        
        if ($baseSev -eq "CRITICAL" -and -not $chkCritical.Checked) { $show = $false }
        elseif ($baseSev -eq "ATTENTION" -and -not $chkAttention.Checked) { $show = $false }
        elseif ($baseSev -eq "TROUBLE" -and -not $chkTrouble.Checked) { $show = $false }
        elseif ($sev -eq "CLEAR" -and -not $chkClearFilter.Checked) { $show = $false }
        
        # Filtro multi-dominio: se selectedDomains non è vuoto, mostra solo quelli
        if ($show -and $script:selectedDomains.Count -gt 0) {
            if ($dom -notin $script:selectedDomains) { $show = $false }
        }
        
        if ($show) {
            $subject = $row.Cells["Subject"].Value
            $sender = $row.Cells["Sender"].Value
            $content = "$subject $sender".ToLower()
            
            foreach ($filter in $script:exclusionFilters) {
                if ($filter.Domain -eq "-- ALL --" -or $filter.Domain -eq $dom) {
                    if ($content.Contains($filter.Word.ToLower())) {
                        $show = $false
                        break
                    }
                }
            }
        }
        
        $row.Visible = $show
    }
}

# =================== REFRESH GRID CON NUOVE REGOLE (senza riscan email) ===================
function Refresh-GridWithNewRules {
    if ($script:allEmails.Count -eq 0) { return }
    
    $lblStatusText.Text = "Applying rules..."
    [System.Windows.Forms.Application]::DoEvents()
    
    # Reset match state
    foreach ($email in $script:allEmails) {
        $email.IsCleared = $false
        $email.MatchedClearId = $null
        $email.ClearTime = $null
        $email.ClearEmailSubject = $null
        $email.ClearEmailBody = $null
        $email.ClearEmailSender = $null
        $email.ClearEmailIP = $null
        
        # Ricalcola match keys - solo se esiste una regola per questa severity
        $rule = Get-MatchRuleForDomain -Domain $email.Domain -Severity $email.Sev
        if ($rule) {
            $email.MatchKeys = Get-MatchKeyForEmail -Email $email -MatchType $rule.MatchType -CustomPattern $rule.CustomPattern
            if ($email.MatchKeys.Count -gt 0) {
                $email.MatchKeyDisplay = $email.MatchKeys -join ", "
            } else {
                $email.MatchKeyDisplay = "-"
            }
        } else {
            # Nessuna regola = nessun match key
            $email.MatchKeys = @()
            $email.MatchKeyDisplay = "(no rule)"
        }
    }
    
    # Rifai il matching
    $clearEmails = @{}
    $alertEmails = @{}
    
    # Per Combined, teniamo traccia delle email con i loro keys
    $combinedClears = @()
    $combinedAlerts = @()
    
    foreach ($email in $script:allEmails) {
        $rule = Get-MatchRuleForDomain -Domain $email.Domain -Severity $email.Sev
        $isCombined = ($rule -and $rule.MatchType -eq "Combined (All Fields)")
        
        if ($isCombined) {
            # Per Combined, salviamo l'email intera con i suoi keys
            if ($email.Sev -eq "CLEAR") {
                $combinedClears += $email
            }
            elseif ($email.Sev -in @("CRITICAL", "ATTENTION", "TROUBLE")) {
                $combinedAlerts += $email
            }
        } else {
            # Per altri tipi, usa il sistema standard basato su chiavi esatte
            foreach ($key in $email.MatchKeys) {
                if ($email.Sev -eq "CLEAR") {
                    if (-not $clearEmails.ContainsKey($key)) { $clearEmails[$key] = @() }
                    $clearEmails[$key] += $email
                }
                elseif ($email.Sev -in @("CRITICAL", "ATTENTION", "TROUBLE")) {
                    if (-not $alertEmails.ContainsKey($key)) { $alertEmails[$key] = @() }
                    $alertEmails[$key] += $email
                }
            }
        }
    }
    
    $matchedAlertIds = @{}
    $matchedClearIds = @{}
    
    # =================== ESCALATION SEVERITY ===================
    # ATTENTION -> TROUBLE -> CRITICAL (livello superiore nasconde inferiore)
    # Raggruppa alert per match keys
    $alertsByDevice = @{}
    
    foreach ($email in $script:allEmails) {
        if ($email.Sev -notin @("CRITICAL", "ATTENTION", "TROUBLE")) { continue }
        if ($email.MatchKeys.Count -eq 0) { continue }
        
        # Crea una chiave unica per questo dispositivo basata sui suoi match keys
        $deviceKey = ($email.MatchKeys | Sort-Object) -join "|"
        
        if (-not $alertsByDevice.ContainsKey($deviceKey)) {
            $alertsByDevice[$deviceKey] = @()
        }
        $alertsByDevice[$deviceKey] += $email
    }
    
    # Per ogni dispositivo, applica escalation
    $severityRank = @{ "ATTENTION" = 1; "TROUBLE" = 2; "CRITICAL" = 3 }
    $escalatedIds = @{}
    
    foreach ($deviceKey in $alertsByDevice.Keys) {
        $deviceAlerts = $alertsByDevice[$deviceKey] | Sort-Object { $_.Time }
        
        if ($deviceAlerts.Count -le 1) { continue }
        
        # Trova il livello più alto per questo dispositivo
        $maxSeverity = 0
        $maxSeverityEmail = $null
        
        foreach ($alert in $deviceAlerts) {
            $rank = $severityRank[$alert.Sev]
            if ($rank -gt $maxSeverity) {
                $maxSeverity = $rank
                $maxSeverityEmail = $alert
            }
        }
        
        # Nascondi tutti gli alert di livello inferiore
        foreach ($alert in $deviceAlerts) {
            if ($alert.Id -ne $maxSeverityEmail.Id) {
                $rank = $severityRank[$alert.Sev]
                if ($rank -lt $maxSeverity) {
                    $escalatedIds[$alert.Id] = $maxSeverityEmail.Sev
                }
            }
        }
    }
    
    # Segna le email escalate
    foreach ($email in $script:allEmails) {
        if ($escalatedIds.ContainsKey($email.Id)) {
            $email.IsEscalated = $true
            $email.EscalatedTo = $escalatedIds[$email.Id]
        } else {
            $email.IsEscalated = $false
            $email.EscalatedTo = $null
        }
    }
    
    # MATCHING PER COMBINED (almeno 2 parametri uguali)
    foreach ($alert in $combinedAlerts) {
        if ($matchedAlertIds.ContainsKey($alert.Id)) { continue }
        
        foreach ($clear in $combinedClears) {
            if ($matchedClearIds.ContainsKey($clear.Id)) { continue }
            
            # Verifica se almeno 2 parametri sono uguali
            if (Test-CombinedMatch -Keys1 $alert.MatchKeys -Keys2 $clear.MatchKeys -MinMatches 2) {
                $matchedAlertIds[$alert.Id] = $true
                $matchedClearIds[$clear.Id] = $true
                
                $alert.IsCleared = $true
                $alert.ClearTime = $clear.Time
                $alert.ClearEmailSubject = $clear.Subject
                $alert.ClearEmailBody = $clear.Body
                $alert.ClearEmailSender = $clear.Sender
                $alert.ClearEmailIP = $clear.MatchKeyDisplay
                $alert.MatchedClearId = $clear.Id
                
                $clear.MatchedClearId = "used"
                break
            }
        }
    }
    
    # UN CLEAR RISOLVE TUTTI GLI ALERT CON LO STESSO MATCH KEY (per matching standard)
    foreach ($key in $alertEmails.Keys) {
        if ($clearEmails.ContainsKey($key)) {
            $alerts = $alertEmails[$key] | Sort-Object { $_.Time }
            $clears = $clearEmails[$key] | Sort-Object { $_.Time }
            
            # Prendi il primo CLEAR disponibile per questo key
            $bestClear = $null
            foreach ($clear in $clears) {
                if (-not $matchedClearIds.ContainsKey($clear.Id)) {
                    $bestClear = $clear
                    $matchedClearIds[$bestClear.Id] = $true
                    break
                }
            }
            
            if ($bestClear) {
                # Questo CLEAR risolve TUTTI gli alert con lo stesso key
                foreach ($alert in $alerts) {
                    if ($matchedAlertIds.ContainsKey($alert.Id)) { continue }
                    
                    $matchedAlertIds[$alert.Id] = $true
                    
                    $alert.IsCleared = $true
                    $alert.ClearTime = $bestClear.Time
                    $alert.ClearEmailSubject = $bestClear.Subject
                    $alert.ClearEmailBody = $bestClear.Body
                    $alert.ClearEmailSender = $bestClear.Sender
                    $alert.ClearEmailIP = $bestClear.MatchKeyDisplay
                    $alert.MatchedClearId = $bestClear.Id
                }
                
                $bestClear.MatchedClearId = "multiple"
            }
        }
    }
    
    # Ripopola grid
    $grid.SuspendLayout()
    $grid.Rows.Clear()
    $script:mailMap.Clear()
    
    # Ordina per data decrescente (più recenti in alto)
    $sortedEmails = $script:allEmails | Sort-Object { $_.Time } -Descending
    
    foreach ($email in $sortedEmails) {
        # Nascondi TUTTI i CLEAR standalone (si vedono solo come risoluzione di alert CRITICAL/ATTENTION/TROUBLE)
        if ($email.Sev -eq "CLEAR") { continue }
        
        # Nascondi email escalate (es. ATTENTION nascosto da TROUBLE/CRITICAL)
        if ($email.IsEscalated) { continue }
        
        $displaySev = $email.Sev
        $displayStatus = "ACTIVE"
        
        if ($email.IsCleared) {
            $displaySev = "$($email.Sev) -> CLEAR"
            $displayStatus = "CLEARED"
        }
        elseif ($email.Sev -eq "CLEAR") {
            $displayStatus = "CLEAR"
        }
        
        $grid.Rows.Add(@(
            $displaySev,
            $email.Domain,
            $email.Sender,
            $email.Subject,
            $email.Time.ToString("dd/MM/yyyy HH:mm:ss"),
            $email.MatchKeyDisplay,
            $displayStatus
        ))
        
        $rowIdx = $grid.Rows.Count - 1
        $script:mailMap[$rowIdx] = $email.Mail
        
        if ($email.IsCleared) {
            Set-RowColor -Row $grid.Rows[$rowIdx] -Severity $email.Sev -HasClear $true
        } else {
            Set-RowColor -Row $grid.Rows[$rowIdx] -Severity $email.Sev
        }
    }
    
    $grid.ResumeLayout()
    $grid.Refresh()
    
    Update-Stats
    Apply-Filters
    
    $lblStatusText.Text = "Rules applied"
    $lblStatusDot.ForeColor = $theme.clear
}

function Get-AllMailFolders {
    param([object]$Folder, [int]$Depth = 0)
    
    $list = New-Object System.Collections.ArrayList
    [void]$list.Add($Folder)
    
    foreach ($sub in $Folder.Folders) {
        try {
            $subList = Get-AllMailFolders -Folder $sub -Depth ($Depth + 1)
            foreach ($f in $subList) { [void]$list.Add($f) }
        } catch {}
    }
    
    return $list
}

# Variabile per log delle cartelle
$script:folderScanLog = [System.Collections.ArrayList]::new()

function Invoke-ScanMails {
    param([bool]$IsInitial = $false)
    
    $newCount = 0
    
    try {
        if (-not $script:selectedStore) { return 0 }
        
        $startDT = $dtpDate.Value.Date.AddHours($nudHour.Value).AddMinutes($nudMinute.Value)
        # Usa formato dd/MM/yyyy per sistemi italiani
        $filterDate = $startDT.ToString("dd/MM/yyyy HH:mm")
        $filter = "[ReceivedTime] >= '$filterDate'"
        
        $root = $script:selectedStore.GetRootFolder()
        $allFolders = Get-AllMailFolders -Folder $root
        
        # Reset log cartelle
        $script:folderScanLog.Clear()
        
        if ($IsInitial) {
            $lblStatusText.Text = "Scanning " + $allFolders.Count + " folders..."
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        $processed = 0
        $newEmails = [System.Collections.ArrayList]::new()
        $totalFoundInFolders = 0
        
        foreach ($folder in $allFolders) {
            
            $folderName = $folder.Name
            $folderPath = try { $folder.FolderPath } catch { $folderName }
            
            # Nessuna esclusione - legge TUTTE le cartelle
            
            $folderEmailCount = 0
            $folderCriticalCount = 0
            $folderClearCount = 0
            
            try {
                $filteredItems = $folder.Items.Restrict($filter)
                $filteredItems.Sort("[ReceivedTime]", $true)
            } catch {
                [void]$script:folderScanLog.Add("ERROR: $folderPath - Cannot access")
                continue
            }
            
            foreach ($mail in $filteredItems) {
                
                $processed++
                if ($processed % 20 -eq 0) {
                    [System.Windows.Forms.Application]::DoEvents()
                    if (-not $script:isRunning) { break }
                }
                
                try {
                    if ($mail.Class -ne 43) { continue }
                    
                    $entryId = $mail.EntryID
                    if ($script:scannedIds.ContainsKey($entryId)) { continue }
                    $script:scannedIds[$entryId] = $true
                    
                    $subject = $mail.Subject
                    $body = $mail.Body
                    $text = ("$subject $body").ToUpperInvariant()
                    
                    $sev = $null
                    $isClear = $false
                    
                    # Riconosci tutte le severity
                    if ($text -match "CLEAR") { $sev = "CLEAR"; $isClear = $true }
                    elseif ($text -match "CRITICAL") { $sev = "CRITICAL" }
                    elseif ($text -match "ATTENTION") { $sev = "ATTENTION" }
                    elseif ($text -match "TROUBLE") { $sev = "TROUBLE" }
                    else { continue }
                    
                    $senderInfo = Get-SenderInfo -Mail $mail
                    $sender = $senderInfo.Email
                    $senderName = $senderInfo.Name
                    $domain = Get-EmailDomain -Email $sender
                    $recTime = $mail.ReceivedTime
                    
                    $displaySender = if ($sender -match "^/O=") { $senderName } else { $sender }
                    
                    $extractedIPs = Extract-IPAddresses -Text "$subject $body"
                    $ipDisplay = "-"
                    if ($extractedIPs.Count -gt 0) { 
                        $ipDisplay = $extractedIPs -join ", " 
                    }
                    
                    [void]$newEmails.Add(@{
                        Id = $entryId
                        Sev = $sev
                        IsClear = $isClear
                        IsCleared = $false
                        Domain = $domain
                        Sender = $displaySender
                        SenderName = $senderName
                        Subject = $subject
                        Time = $recTime
                        Mail = $mail
                        Body = $body
                        IPs = $extractedIPs
                        IPDisplay = $ipDisplay
                        MatchKeys = @()
                        MatchKeyDisplay = $ipDisplay
                        ClearTime = $null
                        ClearEmailSubject = $null
                        ClearEmailBody = $null
                        ClearEmailSender = $null
                        ClearEmailIP = $null
                        MatchedClearId = $null
                        FolderPath = $folderPath
                    })
                    
                    $folderEmailCount++
                    if ($sev -eq "CRITICAL") { $folderCriticalCount++ }
                    elseif ($sev -eq "CLEAR") { $folderClearCount++ }
                    
                } catch { continue }
            }
            
            # Log del risultato per questa cartella
            if ($folderEmailCount -gt 0) {
                [void]$script:folderScanLog.Add("FOUND: $folderPath -> $folderCriticalCount CRITICAL, $folderClearCount CLEAR")
                $totalFoundInFolders += $folderEmailCount
            }
            
            if (-not $script:isRunning) { break }
        }
        
        # Log finale
        [void]$script:folderScanLog.Add("---")
        [void]$script:folderScanLog.Add("TOTAL: $totalFoundInFolders alerts in " + $allFolders.Count + " folders")
        
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
                    } else {
                        $email.MatchKeyDisplay = "-"
                    }
                } else {
                    # Nessuna regola = nessun match key
                    $email.MatchKeys = @()
                    $email.MatchKeyDisplay = "(no rule)"
                }
                
                $script:allEmails = @($email) + $script:allEmails
            }
            
            # Matching
            $clearEmails = @{}
            $alertEmails = @{}
            
            # Per Combined, teniamo traccia delle email con i loro keys
            $combinedClears = @()
            $combinedAlerts = @()
            
            foreach ($email in $script:allEmails) {
                if ($email.MatchedClearId) { continue }
                
                $rule = Get-MatchRuleForDomain -Domain $email.Domain -Severity $email.Sev
                $isCombined = ($rule -and $rule.MatchType -eq "Combined (All Fields)")
                
                if ($isCombined) {
                    if ($email.Sev -eq "CLEAR" -and -not $email.MatchedClearId) {
                        $combinedClears += $email
                    }
                    elseif ($email.Sev -in @("CRITICAL", "ATTENTION", "TROUBLE") -and -not $email.IsCleared) {
                        $combinedAlerts += $email
                    }
                } else {
                    foreach ($key in $email.MatchKeys) {
                        if ($email.Sev -eq "CLEAR" -and -not $email.MatchedClearId) {
                            if (-not $clearEmails.ContainsKey($key)) { $clearEmails[$key] = @() }
                            $clearEmails[$key] += $email
                        }
                        elseif ($email.Sev -in @("CRITICAL", "ATTENTION", "TROUBLE") -and -not $email.IsCleared) {
                            if (-not $alertEmails.ContainsKey($key)) { $alertEmails[$key] = @() }
                            $alertEmails[$key] += $email
                        }
                    }
                }
            }
            
            $matchedAlertIds = @{}
            $matchedClearIds = @{}
            
            # =================== ESCALATION SEVERITY ===================
            # ATTENTION -> TROUBLE -> CRITICAL (livello superiore nasconde inferiore)
            $alertsByDevice = @{}
            
            foreach ($email in $script:allEmails) {
                if ($email.Sev -notin @("CRITICAL", "ATTENTION", "TROUBLE")) { continue }
                if ($email.MatchKeys.Count -eq 0) { continue }
                
                $deviceKey = ($email.MatchKeys | Sort-Object) -join "|"
                
                if (-not $alertsByDevice.ContainsKey($deviceKey)) {
                    $alertsByDevice[$deviceKey] = @()
                }
                $alertsByDevice[$deviceKey] += $email
            }
            
            $severityRank = @{ "ATTENTION" = 1; "TROUBLE" = 2; "CRITICAL" = 3 }
            $escalatedIds = @{}
            
            foreach ($deviceKey in $alertsByDevice.Keys) {
                $deviceAlerts = $alertsByDevice[$deviceKey] | Sort-Object { $_.Time }
                
                if ($deviceAlerts.Count -le 1) { continue }
                
                $maxSeverity = 0
                $maxSeverityEmail = $null
                
                foreach ($alert in $deviceAlerts) {
                    $rank = $severityRank[$alert.Sev]
                    if ($rank -gt $maxSeverity) {
                        $maxSeverity = $rank
                        $maxSeverityEmail = $alert
                    }
                }
                
                foreach ($alert in $deviceAlerts) {
                    if ($alert.Id -ne $maxSeverityEmail.Id) {
                        $rank = $severityRank[$alert.Sev]
                        if ($rank -lt $maxSeverity) {
                            $escalatedIds[$alert.Id] = $maxSeverityEmail.Sev
                        }
                    }
                }
            }
            
            foreach ($email in $script:allEmails) {
                if ($escalatedIds.ContainsKey($email.Id)) {
                    $email.IsEscalated = $true
                    $email.EscalatedTo = $escalatedIds[$email.Id]
                } else {
                    $email.IsEscalated = $false
                    $email.EscalatedTo = $null
                }
            }
            
            # MATCHING PER COMBINED (almeno 2 parametri uguali)
            foreach ($alert in $combinedAlerts) {
                if ($matchedAlertIds.ContainsKey($alert.Id)) { continue }
                if ($alert.IsCleared) { continue }
                
                foreach ($clear in $combinedClears) {
                    if ($matchedClearIds.ContainsKey($clear.Id)) { continue }
                    if ($clear.MatchedClearId) { continue }
                    
                    # Verifica se almeno 2 parametri sono uguali
                    if (Test-CombinedMatch -Keys1 $alert.MatchKeys -Keys2 $clear.MatchKeys -MinMatches 2) {
                        $matchedAlertIds[$alert.Id] = $true
                        $matchedClearIds[$clear.Id] = $true
                        
                        $alert.IsCleared = $true
                        $alert.ClearTime = $clear.Time
                        $alert.ClearEmailSubject = $clear.Subject
                        $alert.ClearEmailBody = $clear.Body
                        $alert.ClearEmailSender = $clear.Sender
                        $alert.ClearEmailIP = $clear.MatchKeyDisplay
                        $alert.MatchedClearId = $clear.Id
                        
                        $clear.MatchedClearId = "used"
                        break
                    }
                }
            }
            
            # UN CLEAR RISOLVE TUTTI GLI ALERT CON LO STESSO MATCH KEY (per matching standard)
            foreach ($key in $alertEmails.Keys) {
                if ($clearEmails.ContainsKey($key)) {
                    $alerts = $alertEmails[$key] | Sort-Object { $_.Time }
                    $clears = $clearEmails[$key] | Sort-Object { $_.Time }
                    
                    # Prendi il primo CLEAR disponibile per questo key
                    $bestClear = $null
                    foreach ($clear in $clears) {
                        if (-not $matchedClearIds.ContainsKey($clear.Id)) {
                            if (-not $clear.MatchedClearId) {
                                $bestClear = $clear
                                $matchedClearIds[$bestClear.Id] = $true
                                break
                            }
                        }
                    }
                    
                    if ($bestClear) {
                        # Questo CLEAR risolve TUTTI gli alert con lo stesso key
                        foreach ($alert in $alerts) {
                            if ($matchedAlertIds.ContainsKey($alert.Id)) { continue }
                            if ($alert.IsCleared) { continue }
                            
                            $matchedAlertIds[$alert.Id] = $true
                            
                            $alert.IsCleared = $true
                            $alert.ClearTime = $bestClear.Time
                            $alert.ClearEmailSubject = $bestClear.Subject
                            $alert.ClearEmailBody = $bestClear.Body
                            $alert.ClearEmailSender = $bestClear.Sender
                            $alert.ClearEmailIP = $bestClear.MatchKeyDisplay
                            $alert.MatchedClearId = $bestClear.Id
                        }
                        
                        $bestClear.MatchedClearId = "multiple"
                    }
                }
            }
            
            $grid.SuspendLayout()
            $grid.Rows.Clear()
            $script:mailMap.Clear()
            
            # Ordina per data decrescente (più recenti in alto)
            $sortedEmails = $script:allEmails | Sort-Object { $_.Time } -Descending
            
            foreach ($email in $sortedEmails) {
                # Nascondi TUTTI i CLEAR standalone (si vedono solo come risoluzione di alert CRITICAL/ATTENTION/TROUBLE)
                if ($email.Sev -eq "CLEAR") { continue }
                
                # Nascondi email escalate (es. ATTENTION nascosto da TROUBLE/CRITICAL)
                if ($email.IsEscalated) { continue }
                
                $displaySev = $email.Sev
                $displayStatus = "ACTIVE"
                
                if ($email.IsCleared) {
                    $displaySev = "$($email.Sev) -> CLEAR"
                    $displayStatus = "CLEARED"
                }
                
                $grid.Rows.Add(@(
                    $displaySev,
                    $email.Domain,
                    $email.Sender,
                    $email.Subject,
                    $email.Time.ToString("dd/MM/yyyy HH:mm:ss"),
                    $email.MatchKeyDisplay,
                    $displayStatus
                ))
                
                $rowIdx = $grid.Rows.Count - 1
                $script:mailMap[$rowIdx] = $email.Mail
                
                if ($email.IsCleared) {
                    Set-RowColor -Row $grid.Rows[$rowIdx] -Severity $email.Sev -HasClear $true
                } else {
                    Set-RowColor -Row $grid.Rows[$rowIdx] -Severity $email.Sev
                }
            }
            
            $grid.ResumeLayout()
            $grid.Refresh()
            $newCount = $newEmails.Count
        }
        
        Update-Stats
        Apply-Filters
        
    } catch {}
    
    [System.Windows.Forms.Application]::DoEvents()
    return $newCount
}

# =================== TIMER ===================
$script:timer = New-Object System.Windows.Forms.Timer
$script:timer.Add_Tick({
    try {
        if (-not $script:isRunning -or -not $chkContinuous.Checked) { return }
        $lblStatusText.Text = "Checking..."
        $lblStatusDot.ForeColor = $theme.attention
        [System.Windows.Forms.Application]::DoEvents()
        
        # Memorizza conteggi prima dello scan
        $prevCritical = ($script:allEmails | Where-Object { $_.Sev -eq "CRITICAL" -and -not $_.IsCleared }).Count
        $prevAttention = ($script:allEmails | Where-Object { $_.Sev -eq "ATTENTION" -and -not $_.IsCleared }).Count
        $prevTrouble = ($script:allEmails | Where-Object { $_.Sev -eq "TROUBLE" -and -not $_.IsCleared }).Count
        
        $new = Invoke-ScanMails -IsInitial $false
        
        if ($new -gt 0) { 
            $lblStatusText.Text = "+$new new @ $(Get-Date -Format 'HH:mm:ss')" 
            
            # Calcola nuovi conteggi
            $newCritical = ($script:allEmails | Where-Object { $_.Sev -eq "CRITICAL" -and -not $_.IsCleared }).Count
            $newAttention = ($script:allEmails | Where-Object { $_.Sev -eq "ATTENTION" -and -not $_.IsCleared }).Count
            $newTrouble = ($script:allEmails | Where-Object { $_.Sev -eq "TROUBLE" -and -not $_.IsCleared }).Count
            
            # Suono basato sulla severity più alta trovata (se abilitato)
            if ($chkSound.Checked) {
                if ($newCritical -gt $prevCritical) {
                    Play-SeveritySound -Severity "CRITICAL"
                } elseif ($newAttention -gt $prevAttention) {
                    Play-SeveritySound -Severity "ATTENTION"
                } elseif ($newTrouble -gt $prevTrouble) {
                    Play-SeveritySound -Severity "TROUBLE"
                } else {
                    # Suono generico per CLEAR o altro
                    try { [Console]::Beep(600, 150) } catch {}
                }
            }
        }
        else { $lblStatusText.Text = "OK @ $(Get-Date -Format 'HH:mm:ss')" }
        $lblStatusDot.ForeColor = $theme.clear
    } catch { $lblStatusText.Text = "Error" }
})

# =================== EVENTS ===================
$btnStart.Add_Click({
    try {
        if (-not $cmbProfile.SelectedItem) { 
            [System.Windows.Forms.MessageBox]::Show("Select profile", "Warning", "OK", "Warning") | Out-Null
            return 
        }
        
        if ($script:isRunning) {
            $script:isRunning = $false
            $script:timer.Stop()
            $btnStart.Text = "START"
            $btnStart.BackColor = $theme.accent
            $lblStatusText.Text = "Stopped"
            $lblStatusDot.ForeColor = $theme.textMuted
            [System.Windows.Forms.Application]::DoEvents()
            return
        }
        
        $script:isRunning = $true
        $btnStart.Text = "STOP"
        $btnStart.BackColor = $theme.critical
        $lblStatusText.Text = "Starting..."
        $lblStatusDot.ForeColor = $theme.attention
        [System.Windows.Forms.Application]::DoEvents()
        
        $grid.SuspendLayout()
        $grid.Rows.Clear()
        $script:mailMap.Clear()
        $script:scannedIds.Clear()
        $script:allEmails = @()
        $script:allDomains.Clear()
        $script:alertClearMap.Clear()
        $script:ipToRowMap.Clear()
        $script:selectedDomains.Clear()
        $btnSelectDomains.Text = "ALL DOMAINS"
        $cmbExcludeDomain.Items.Clear()
        [void]$cmbExcludeDomain.Items.Add("-- ALL --")
        $cmbExcludeDomain.SelectedIndex = 0
        $cmbMatchDomain.Items.Clear()
        [void]$cmbMatchDomain.Items.Add("-- ALL DOMAINS --")
        $cmbMatchDomain.SelectedIndex = 0
        $grid.ResumeLayout()
        [System.Windows.Forms.Application]::DoEvents()
        
        $script:selectedStore = $null
        foreach ($store in $script:namespace.Stores) { 
            if ($store.DisplayName -eq $cmbProfile.SelectedItem) { 
                $script:selectedStore = $store
                break 
            } 
        }
        
        if (-not $script:selectedStore) { 
            [System.Windows.Forms.MessageBox]::Show("Store not found", "Error", "OK", "Error") | Out-Null
            $script:isRunning = $false
            $btnStart.Text = "START"
            $btnStart.BackColor = $theme.accent
            return 
        }
        
        $cnt = Invoke-ScanMails -IsInitial $true
        $lblStatusText.Text = "Found $cnt alerts"
        $lblStatusDot.ForeColor = $theme.clear
        [System.Windows.Forms.Application]::DoEvents()
        
        if ($chkContinuous.Checked) { 
            $script:timer.Interval = [int]($nudInterval.Value * 1000)
            $script:timer.Start()
            $lblStatusText.Text = "Active - $cnt alerts"
        }
    } catch { 
        $script:isRunning = $false
        $btnStart.Text = "START"
        $btnStart.BackColor = $theme.accent
        $lblStatusText.Text = "Error"
        $lblStatusDot.ForeColor = $theme.critical
        [System.Windows.Forms.MessageBox]::Show(
            "Errore durante la scansione:`n$($_.Exception.Message)",
            "Errore",
            "OK",
            "Error"
        )
        [System.Windows.Forms.Application]::DoEvents() 
    }
})

$btnClear.Add_Click({ 
    $grid.Rows.Clear()
    $script:mailMap.Clear()
    $script:allEmails = @()
    $script:alertClearMap.Clear()
    $script:ipToRowMap.Clear()
    Update-Stats
    $lblStatusText.Text = "Cleared"
})

$chkCritical.Add_CheckedChanged({ Apply-Filters })
$chkAttention.Add_CheckedChanged({ Apply-Filters })
$chkTrouble.Add_CheckedChanged({ Apply-Filters })
$chkClearFilter.Add_CheckedChanged({ Apply-Filters })

# =================== FONT SIZE FUNCTION ===================
function Apply-FontSize {
    param([string]$Size)
    
    $script:currentFontSize = $Size
    $fs = $script:fontSizes[$Size]
    
    # Grid
    $grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", $fs.Grid)
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Label, [System.Drawing.FontStyle]::Bold)
    $grid.RowTemplate.Height = $fs.RowHeight
    
    # Aggiorna altezza righe esistenti
    foreach ($row in $grid.Rows) {
        $row.Height = $fs.RowHeight
    }
    
    # Header logo
    $lblLogo.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Header, [System.Drawing.FontStyle]::Bold)
    
    # Stats
    $lblStatCritical.Font = New-Object System.Drawing.Font("Consolas", $fs.Stats, [System.Drawing.FontStyle]::Bold)
    $lblStatClear.Font = New-Object System.Drawing.Font("Consolas", $fs.Stats, [System.Drawing.FontStyle]::Bold)
    $lblStatTotal.Font = New-Object System.Drawing.Font("Consolas", $fs.Stats, [System.Drawing.FontStyle]::Bold)
    
    # Labels
    $lblProfile.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Label, [System.Drawing.FontStyle]::Bold)
    $lblDate.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Label, [System.Drawing.FontStyle]::Bold)
    $lblInterval.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Label, [System.Drawing.FontStyle]::Bold)
    $lblFontSize.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Label, [System.Drawing.FontStyle]::Bold)
    $lblSevFilter.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Label, [System.Drawing.FontStyle]::Bold)
    $lblDomFilter.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Label, [System.Drawing.FontStyle]::Bold)
    $lblExcludeTitle.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Label, [System.Drawing.FontStyle]::Bold)
    $lblMatchTitle.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Label, [System.Drawing.FontStyle]::Bold)
    
    # Buttons
    $btnStart.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Button, [System.Drawing.FontStyle]::Bold)
    $btnAddExclusion.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Button, [System.Drawing.FontStyle]::Bold)
    $btnAddMatchRule.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Button, [System.Drawing.FontStyle]::Bold)
    $btnSelectDomains.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Button)
    
    # Checkboxes
    $chkCritical.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    $chkAttention.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    $chkTrouble.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    $chkClearFilter.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    $chkContinuous.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    
    # ComboBoxes
    $cmbProfile.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    $cmbFontSize.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    $cmbExcludeDomain.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    $cmbMatchDomain.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    $cmbMatchType.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    
    # TextBoxes
    $txtExcludeWord.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    $txtCustomPattern.Font = New-Object System.Drawing.Font("Segoe UI", $fs.Normal)
    
    # Lists
    $listExclusions.Font = New-Object System.Drawing.Font("Consolas", $fs.Normal)
    $listMatchRules.Font = New-Object System.Drawing.Font("Consolas", $fs.Normal)
    
    # Status
    $lblStatusText.Font = New-Object System.Drawing.Font("Segoe UI", ($fs.Normal + 1))
    
    $grid.Refresh()
}

$cmbFontSize.Add_SelectedIndexChanged({
    Apply-FontSize -Size $cmbFontSize.SelectedItem.ToString()
})

$btnAddExclusion.Add_Click({
    $word = $txtExcludeWord.Text.Trim()
    if ([string]::IsNullOrEmpty($word)) { return }
    
    $domain = $cmbExcludeDomain.SelectedItem.ToString()
    
    $exists = $script:exclusionFilters | Where-Object { $_.Domain -eq $domain -and $_.Word -eq $word }
    if ($exists) { return }
    
    [void]$script:exclusionFilters.Add(@{ Domain = $domain; Word = $word })
    $txtExcludeWord.Text = ""
    Update-ExclusionsList
    Apply-Filters
})

$btnAddMatchRule.Add_Click({
    $domain = $cmbMatchDomain.SelectedItem.ToString()
    $matchType = $cmbMatchType.SelectedItem.ToString()
    $severity = $cmbMatchSeverity.SelectedItem.ToString()
    $customPattern = $txtCustomPattern.Text.Trim()
    
    if ($matchType -eq "Custom Regex" -and [string]::IsNullOrEmpty($customPattern)) {
        [System.Windows.Forms.MessageBox]::Show("Inserisci un pattern regex per Custom Regex", "Warning", "OK", "Warning") | Out-Null
        return
    }
    
    # Rimuove regole esistenti per lo stesso dominio E severity
    $toRemove = $script:matchRules | Where-Object { $_.Domain -eq $domain -and $_.Severity -eq $severity }
    if ($toRemove) {
        $script:matchRules.Remove($toRemove)
    }
    
    [void]$script:matchRules.Add(@{ 
        Domain = $domain
        MatchType = $matchType
        Severity = $severity
        CustomPattern = $customPattern
    })
    
    $txtCustomPattern.Text = ""
    Update-MatchRulesList
    Refresh-GridWithNewRules
})

$nudInterval.Add_ValueChanged({ 
    if ($script:timer.Enabled) { 
        $script:timer.Stop()
        $script:timer.Interval = [int]($nudInterval.Value * 1000)
        $script:timer.Start() 
    } 
})

# =================== DUAL EMAIL DETAILS ===================
function Show-DualEmailDetails {
    param(
        $AlertEmail,
        $Domain,
        $MatchKey,
        $Time
    )
    
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $formWidth = [Math]::Min(1400, [int]($screen.Width * 0.85))
    $formHeight = [Math]::Min(800, [int]($screen.Height * 0.85))
    
    $detailForm = New-Object System.Windows.Forms.Form
    $detailForm.Text = "Alert Risolto - Dettagli Completi"
    $detailForm.Size = New-Object System.Drawing.Size($formWidth, $formHeight)
    $detailForm.StartPosition = "CenterScreen"
    $detailForm.BackColor = $theme.bgDeep
    $detailForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $detailForm.MinimumSize = New-Object System.Drawing.Size(500, 400)
    
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = "Top"
    $headerPanel.Height = 50
    $headerPanel.BackColor = $theme.bgPanel
    $detailForm.Controls.Add($headerPanel)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "ALERT RISOLTO - CONFRONTO EMAIL"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $theme.accent
    $lblTitle.Location = New-Object System.Drawing.Point(20, 12)
    $lblTitle.AutoSize = $true
    $headerPanel.Controls.Add($lblTitle)
    
    $splitContainer = New-Object System.Windows.Forms.SplitContainer
    $splitContainer.Dock = "Fill"
    $splitContainer.Orientation = "Vertical"
    $splitContainer.BackColor = $theme.bgDeep
    $splitContainer.SplitterWidth = 8
    $splitContainer.Panel1MinSize = 100
    $splitContainer.Panel2MinSize = 100
    $detailForm.Controls.Add($splitContainer)
    
    # Imposta SplitterDistance nell'evento Load per evitare errori
    $detailForm.Add_Load({
        try {
            $splitContainer.SplitterDistance = [int]($splitContainer.Width / 2)
        } catch {
            # Ignora errori di SplitterDistance
        }
    })
    
    $leftPanel = New-Object System.Windows.Forms.Panel
    $leftPanel.Dock = "Fill"
    $leftPanel.BackColor = $theme.bgCard
    $leftPanel.Padding = New-Object System.Windows.Forms.Padding(15)
    $splitContainer.Panel1.Controls.Add($leftPanel)
    
    $leftHeader = New-Object System.Windows.Forms.Label
    $leftHeader.Text = "ALERT ORIGINALE"
    $leftHeader.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $leftHeader.ForeColor = $theme.critical
    $leftHeader.Dock = "Top"
    $leftHeader.Height = 35
    $leftHeader.TextAlign = "MiddleCenter"
    $leftHeader.BackColor = [System.Drawing.Color]::FromArgb(40, 255, 77, 106)
    $leftPanel.Controls.Add($leftHeader)
    
    $leftInfo = New-Object System.Windows.Forms.Panel
    $leftInfo.Dock = "Top"
    $leftInfo.Height = 130
    $leftInfo.BackColor = $theme.bgPanel
    $leftInfo.Padding = New-Object System.Windows.Forms.Padding(10)
    $leftPanel.Controls.Add($leftInfo)
    
    $leftInfoText = New-Object System.Windows.Forms.Label
    $leftInfoText.Text = "SEVERITY: $($AlertEmail.Sev)`nMATCH KEY: $MatchKey`nDOMAIN: $Domain`nFROM: $($AlertEmail.Sender)`nTIME: $Time`nSUBJECT: $($AlertEmail.Subject)"
    $leftInfoText.Font = New-Object System.Drawing.Font("Consolas", 9)
    $leftInfoText.ForeColor = $theme.textPrimary
    $leftInfoText.Dock = "Fill"
    $leftInfoText.AutoSize = $false
    $leftInfo.Controls.Add($leftInfoText)
    
    $leftBody = New-Object System.Windows.Forms.TextBox
    $leftBody.Multiline = $true
    $leftBody.ScrollBars = "Vertical"
    $leftBody.Dock = "Fill"
    $leftBody.BackColor = $theme.bgInput
    $leftBody.ForeColor = $theme.textPrimary
    $leftBody.Font = New-Object System.Drawing.Font("Consolas", 9)
    $leftBody.ReadOnly = $true
    $leftBody.Text = $AlertEmail.Body
    $leftBody.BorderStyle = "None"
    $leftPanel.Controls.Add($leftBody)
    
    $leftBody.BringToFront()
    
    $rightPanel = New-Object System.Windows.Forms.Panel
    $rightPanel.Dock = "Fill"
    $rightPanel.BackColor = $theme.bgCard
    $rightPanel.Padding = New-Object System.Windows.Forms.Padding(15)
    $splitContainer.Panel2.Controls.Add($rightPanel)
    
    $rightHeader = New-Object System.Windows.Forms.Label
    $rightHeader.Text = "EMAIL CLEAR (RISOLUZIONE)"
    $rightHeader.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $rightHeader.ForeColor = $theme.clear
    $rightHeader.Dock = "Top"
    $rightHeader.Height = 35
    $rightHeader.TextAlign = "MiddleCenter"
    $rightHeader.BackColor = [System.Drawing.Color]::FromArgb(40, 0, 232, 144)
    $rightPanel.Controls.Add($rightHeader)
    
    $rightInfo = New-Object System.Windows.Forms.Panel
    $rightInfo.Dock = "Top"
    $rightInfo.Height = 130
    $rightInfo.BackColor = $theme.bgPanel
    $rightInfo.Padding = New-Object System.Windows.Forms.Padding(10)
    $rightPanel.Controls.Add($rightInfo)
    
    $clearTimeStr = ""
    if ($AlertEmail.ClearTime) {
        $clearTimeStr = $AlertEmail.ClearTime.ToString('dd/MM/yyyy HH:mm:ss')
    }
    
    $rightInfoText = New-Object System.Windows.Forms.Label
    $rightInfoText.Text = "SEVERITY: CLEAR`nMATCH KEY: $($AlertEmail.ClearEmailIP)`nDOMAIN: $Domain`nFROM: $($AlertEmail.ClearEmailSender)`nTIME: $clearTimeStr`nSUBJECT: $($AlertEmail.ClearEmailSubject)"
    $rightInfoText.Font = New-Object System.Drawing.Font("Consolas", 9)
    $rightInfoText.ForeColor = $theme.textPrimary
    $rightInfoText.Dock = "Fill"
    $rightInfoText.AutoSize = $false
    $rightInfo.Controls.Add($rightInfoText)
    
    $rightBody = New-Object System.Windows.Forms.TextBox
    $rightBody.Multiline = $true
    $rightBody.ScrollBars = "Vertical"
    $rightBody.Dock = "Fill"
    $rightBody.BackColor = $theme.bgInput
    $rightBody.ForeColor = $theme.textPrimary
    $rightBody.Font = New-Object System.Drawing.Font("Consolas", 9)
    $rightBody.ReadOnly = $true
    $rightBody.Text = $AlertEmail.ClearEmailBody
    $rightBody.BorderStyle = "None"
    $rightPanel.Controls.Add($rightBody)
    
    $rightBody.BringToFront()
    
    $footerPanel = New-Object System.Windows.Forms.Panel
    $footerPanel.Dock = "Bottom"
    $footerPanel.Height = 50
    $footerPanel.BackColor = $theme.bgPanel
    $detailForm.Controls.Add($footerPanel)
    
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "CHIUDI"
    $btnClose.Size = New-Object System.Drawing.Size(120, 35)
    $btnClose.Location = New-Object System.Drawing.Point(([int](($formWidth - 120) / 2)), 8)
    $btnClose.BackColor = $theme.accent
    $btnClose.ForeColor = $theme.bgDeep
    $btnClose.FlatStyle = "Flat"
    $btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Add_Click({ $detailForm.Close() })
    $footerPanel.Controls.Add($btnClose)
    
    $detailForm.Add_Resize({
        $btnClose.Location = New-Object System.Drawing.Point(([int](($detailForm.ClientSize.Width - 120) / 2)), 8)
    })
    
    $splitContainer.BringToFront()
    
    [void]$detailForm.ShowDialog()
}

$grid.Add_CellDoubleClick({
    param($s, $e)
    
    try {
        if ($e.RowIndex -lt 0) { return }
        
        $row = $grid.Rows[$e.RowIndex]
        if (-not $row) { return }
        
        $subject = $row.Cells["Subject"].Value
        $sender = $row.Cells["Sender"].Value
        $domain = $row.Cells["Domain"].Value
        $time = $row.Cells["Time"].Value
        $sev = $row.Cells["Sev"].Value
        $matchKey = $row.Cells["MatchKey"].Value
        $status = $row.Cells["Status"].Value
        
        $foundEmail = $null
        foreach ($email in $script:allEmails) {
            if ($email.Subject -eq $subject -and $email.Sender -eq $sender) {
                $foundEmail = $email
                break
            }
        }
        
        if (-not $foundEmail) {
            [System.Windows.Forms.MessageBox]::Show("Email non trovata nei dati.", "Errore", "OK", "Warning")
            return
        }
        
        if (($status -eq "CLEARED" -or $sev -like "*-> CLEAR*") -and $foundEmail.IsCleared -and $foundEmail.ClearEmailSubject) {
            Show-DualEmailDetails -AlertEmail $foundEmail -Domain $domain -MatchKey $matchKey -Time $time
        }
        else {
            # Form con contenuto copiabile
            $detailForm = New-Object System.Windows.Forms.Form
            $detailForm.Text = "Email Details - $sev"
            $detailForm.Size = New-Object System.Drawing.Size(700, 550)
            $detailForm.StartPosition = "CenterScreen"
            $detailForm.BackColor = $theme.bgDeep
            $detailForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
            
            # Header
            $headerPanel = New-Object System.Windows.Forms.Panel
            $headerPanel.Dock = "Top"
            $headerPanel.Height = 45
            $headerPanel.BackColor = $theme.bgPanel
            $detailForm.Controls.Add($headerPanel)
            
            $lblTitle = New-Object System.Windows.Forms.Label
            $lblTitle.Text = "$sev - DETAILS"
            $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
            $lblTitle.ForeColor = $(if ($sev -eq "CRITICAL") { $theme.critical } elseif ($sev -eq "CLEAR") { $theme.clear } else { $theme.accent })
            $lblTitle.Location = New-Object System.Drawing.Point(20, 10)
            $lblTitle.AutoSize = $true
            $headerPanel.Controls.Add($lblTitle)
            
            # Info panel
            $infoPanel = New-Object System.Windows.Forms.Panel
            $infoPanel.Dock = "Top"
            $infoPanel.Height = 100
            $infoPanel.BackColor = $theme.bgCard
            $infoPanel.Padding = New-Object System.Windows.Forms.Padding(10)
            $detailForm.Controls.Add($infoPanel)
            
            $lblInfo = New-Object System.Windows.Forms.TextBox
            $lblInfo.Multiline = $true
            $lblInfo.ReadOnly = $true
            $lblInfo.BorderStyle = "None"
            $lblInfo.BackColor = $theme.bgCard
            $lblInfo.ForeColor = $theme.textPrimary
            $lblInfo.Font = New-Object System.Drawing.Font("Consolas", 9)
            $lblInfo.Dock = "Fill"
            $lblInfo.Text = "SEVERITY: $sev`r`nSTATUS: $status`r`nMATCH KEY: $matchKey`r`nDOMAIN: $domain`r`nFROM: $sender`r`nTIME: $time"
            $infoPanel.Controls.Add($lblInfo)
            
            # Subject label
            $lblSubjectHeader = New-Object System.Windows.Forms.Label
            $lblSubjectHeader.Text = "SUBJECT:"
            $lblSubjectHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $lblSubjectHeader.ForeColor = $theme.textMuted
            $lblSubjectHeader.Dock = "Top"
            $lblSubjectHeader.Height = 25
            $lblSubjectHeader.Padding = New-Object System.Windows.Forms.Padding(10, 8, 0, 0)
            $detailForm.Controls.Add($lblSubjectHeader)
            
            # Subject textbox
            $txtSubject = New-Object System.Windows.Forms.TextBox
            $txtSubject.Dock = "Top"
            $txtSubject.Height = 25
            $txtSubject.BackColor = $theme.bgInput
            $txtSubject.ForeColor = $theme.textPrimary
            $txtSubject.Font = New-Object System.Drawing.Font("Consolas", 9)
            $txtSubject.ReadOnly = $true
            $txtSubject.Text = $subject
            $txtSubject.BorderStyle = "None"
            $detailForm.Controls.Add($txtSubject)
            
            # Body label
            $lblBodyHeader = New-Object System.Windows.Forms.Label
            $lblBodyHeader.Text = "BODY:"
            $lblBodyHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
            $lblBodyHeader.ForeColor = $theme.textMuted
            $lblBodyHeader.Dock = "Top"
            $lblBodyHeader.Height = 25
            $lblBodyHeader.Padding = New-Object System.Windows.Forms.Padding(10, 8, 0, 0)
            $detailForm.Controls.Add($lblBodyHeader)
            
            # Body textbox (copiabile)
            $txtBody = New-Object System.Windows.Forms.TextBox
            $txtBody.Multiline = $true
            $txtBody.ScrollBars = "Vertical"
            $txtBody.Dock = "Fill"
            $txtBody.BackColor = $theme.bgInput
            $txtBody.ForeColor = $theme.textPrimary
            $txtBody.Font = New-Object System.Drawing.Font("Consolas", 9)
            $txtBody.ReadOnly = $true
            $txtBody.Text = $foundEmail.Body
            $txtBody.BorderStyle = "None"
            $detailForm.Controls.Add($txtBody)
            
            # Footer con pulsanti
            $footerPanel = New-Object System.Windows.Forms.Panel
            $footerPanel.Dock = "Bottom"
            $footerPanel.Height = 50
            $footerPanel.BackColor = $theme.bgPanel
            $detailForm.Controls.Add($footerPanel)
            
            $btnCopyAll = New-Object System.Windows.Forms.Button
            $btnCopyAll.Text = "COPY ALL"
            $btnCopyAll.Size = New-Object System.Drawing.Size(100, 35)
            $btnCopyAll.Location = New-Object System.Drawing.Point(200, 8)
            $btnCopyAll.BackColor = $theme.bgElevated
            $btnCopyAll.ForeColor = $theme.textPrimary
            $btnCopyAll.FlatStyle = "Flat"
            $btnCopyAll.FlatAppearance.BorderColor = $theme.border
            $btnCopyAll.Add_Click({
                $fullText = "SEVERITY: $sev`r`nSTATUS: $status`r`nMATCH KEY: $matchKey`r`nDOMAIN: $domain`r`nFROM: $sender`r`nTIME: $time`r`n`r`nSUBJECT: $subject`r`n`r`nBODY:`r`n$($foundEmail.Body)"
                [System.Windows.Forms.Clipboard]::SetText($fullText)
                $btnCopyAll.Text = "COPIED!"
                $timer = New-Object System.Windows.Forms.Timer
                $timer.Interval = 1500
                $timer.Add_Tick({ $btnCopyAll.Text = "COPY ALL"; $timer.Stop(); $timer.Dispose() })
                $timer.Start()
            })
            $footerPanel.Controls.Add($btnCopyAll)
            
            $btnClose = New-Object System.Windows.Forms.Button
            $btnClose.Text = "CLOSE"
            $btnClose.Size = New-Object System.Drawing.Size(100, 35)
            $btnClose.Location = New-Object System.Drawing.Point(380, 8)
            $btnClose.BackColor = $theme.accent
            $btnClose.ForeColor = $theme.bgDeep
            $btnClose.FlatStyle = "Flat"
            $btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $btnClose.FlatAppearance.BorderSize = 0
            $btnClose.Add_Click({ $detailForm.Close() })
            $footerPanel.Controls.Add($btnClose)
            
            # Porta in primo piano gli elementi nell'ordine giusto
            $txtBody.BringToFront()
            
            [void]$detailForm.ShowDialog()
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Errore durante l'apertura dei dettagli:`n$($_.Exception.Message)",
            "Errore",
            "OK",
            "Error"
        )
    }
})

$form.Add_FormClosing({ 
    $script:isRunning = $false
    if ($script:timer -ne $null) { $script:timer.Stop(); $script:timer.Dispose() }
})

# =================== DEFAULT MATCH RULE ===================
# Aggiungi regola di default: Combined (Device+IP+Serial+Category) per ALL DOMAINS
[void]$script:matchRules.Add(@{ 
    Domain = "-- ALL DOMAINS --"
    MatchType = "Combined (All Fields)"
    Severity = "ALL"
    CustomPattern = ""
})
Update-MatchRulesList

# =================== RUN ===================
[void]$form.ShowDialog()
