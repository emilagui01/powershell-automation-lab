<#
.SYNOPSIS
    Account health report: locked out, disabled, inactive, and password problems.

.DESCRIPTION
    Read-only. Checks every user account in the domain and flags:
      - Locked out accounts
      - Disabled accounts
      - Enabled accounts that have never signed in, or have been idle for InactiveDays+
      - Passwords that must be changed, have expired, or expire within ExpiringDays
      - Accounts whose password never expires

    Writes a CSV (every account, for filtering in Excel) and an HTML report
    (one section per problem, for a manager or a ticket attachment).

.PARAMETER InactiveDays
    Days without a sign-in before an enabled account counts as inactive. Default 30.

.PARAMETER ExpiringDays
    Flag passwords that expire within this many days. Default 14.

.PARAMETER Open
    Open the HTML report when finished.

.EXAMPLE
    .\Get-LabUserReport.ps1 -Open

.EXAMPLE
    .\Get-LabUserReport.ps1 -InactiveDays 60 -ExpiringDays 45

.NOTES
    Author: Emil Aguirre Herrera
    Lab:    Home Lab 3 - PowerShell Automation (emillab.local)
    LastLogonDate comes from lastLogonTimestamp, which AD only updates every
    9-14 days by design. Good for "inactive 30+ days", not for "signed in today".
#>
[CmdletBinding()]
param(
    [int]$InactiveDays  = 30,
    [int]$ExpiringDays  = 14,
    [string]$ReportDir  = "C:\Scripts\Lab3\Reports",
    [string]$SearchBase,
    [switch]$Open
)

Import-Module ActiveDirectory -ErrorAction Stop

if (-not $SearchBase) { $SearchBase = (Get-ADDomain).DistinguishedName }
if (-not (Test-Path $ReportDir)) { New-Item -Path $ReportDir -ItemType Directory | Out-Null }

$Now      = Get-Date
$Stamp    = $Now.ToString("yyyyMMdd-HHmmss")
$CsvFile  = Join-Path $ReportDir "UserReport-$Stamp.csv"
$HtmlFile = Join-Path $ReportDir "UserReport-$Stamp.html"

# Built-in accounts Windows keeps disabled on purpose - not worth reporting
$Skip  = @("krbtgt", "Guest", "DefaultAccount", "WDAGUtilityAccount")
$Props = @("Enabled", "LockedOut", "LastLogonDate", "PasswordLastSet", "PasswordExpired",
           "PasswordNeverExpires", "msDS-UserPasswordExpiryTimeComputed", "Department", "Title")

$users = Get-ADUser -Filter * -SearchBase $SearchBase -Properties $Props |
         Where-Object { $_.SamAccountName -notin $Skip }

# ---------------------------------------------------------------------------
# Check each account
# ---------------------------------------------------------------------------
$rows = foreach ($u in $users) {
    # AD stores the expiry as a FileTime; 0 = must change now, MaxValue = never expires
    $raw     = $u.'msDS-UserPasswordExpiryTimeComputed'
    $expires = if ($raw -and $raw -gt 0 -and $raw -lt [long]::MaxValue) { [datetime]::FromFileTime($raw) } else { $null }
    $idle    = if ($u.LastLogonDate) { [int][math]::Floor(($Now - $u.LastLogonDate).TotalDays) } else { $null }

    $issues = @()
    if (-not $u.Enabled) { $issues += "Disabled" }
    if ($u.LockedOut)    { $issues += "Locked out" }

    if ($u.Enabled) {
        if (-not $u.LastLogonDate)   { $issues += "Never signed in" }
        elseif ($idle -ge $InactiveDays) { $issues += "Inactive $idle days" }

        if (-not $u.PasswordLastSet) { $issues += "Must change password" }
        elseif ($u.PasswordExpired)  { $issues += "Password expired" }
        elseif ($expires -and $expires -le $Now.AddDays($ExpiringDays)) {
            $issues += "Password expires in $([int][math]::Ceiling(($expires - $Now).TotalDays)) days"
        }

        if ($u.PasswordNeverExpires) { $issues += "Password never expires" }
    }

    [pscustomobject]@{
        Username        = $u.SamAccountName
        Name            = $u.Name
        Department      = $u.Department
        Enabled         = $u.Enabled
        LockedOut       = $u.LockedOut
        LastSignIn      = if ($u.LastLogonDate) { $u.LastLogonDate.ToString("yyyy-MM-dd HH:mm") } else { "Never" }
        PasswordLastSet = if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString("yyyy-MM-dd") } else { "Must change" }
        PasswordExpires = if ($u.PasswordNeverExpires) { "Never" } elseif ($expires) { $expires.ToString("yyyy-MM-dd") } else { "" }
        OU              = ($u.DistinguishedName -split "(?<!\\),", 2)[1]
        Issues          = if ($issues) { $issues -join "; " } else { "OK" }
    }
}
$rows = @($rows | Sort-Object Username)

# ---------------------------------------------------------------------------
# Group into report sections
# ---------------------------------------------------------------------------
$sections = [ordered]@{
    "Locked out"                              = @($rows | Where-Object { $_.Issues -match "Locked out" })
    "Disabled"                                = @($rows | Where-Object { $_.Issues -match "Disabled" })
    "Never signed in"                         = @($rows | Where-Object { $_.Issues -match "Never signed in" })
    "Inactive $InactiveDays+ days"            = @($rows | Where-Object { $_.Issues -match "Inactive \d+ days" })
    "Must change password at next sign-in"    = @($rows | Where-Object { $_.Issues -match "Must change password" })
    "Password expired or expiring within $ExpiringDays days" = @($rows | Where-Object { $_.Issues -match "Password expire" })
    "Password never expires"                  = @($rows | Where-Object { $_.Issues -match "never expires" })
}

# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------
$rows | Export-Csv -Path $CsvFile -NoTypeInformation

# ---------------------------------------------------------------------------
# HTML
# ---------------------------------------------------------------------------
$css = @"
<style>
 body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1b1b1b; }
 h1 { margin-bottom: 4px; } .sub { color: #555; margin-top: 0; }
 h2 { margin-top: 28px; border-bottom: 2px solid #012456; padding-bottom: 4px; color: #012456; }
 table { border-collapse: collapse; font-size: 13px; margin-top: 6px; }
 th { background: #012456; color: #fff; text-align: left; padding: 5px 10px; }
 td { border-bottom: 1px solid #ddd; padding: 4px 10px; }
 tr:nth-child(even) td { background: #f4f6fa; }
 .summary td:last-child { text-align: right; font-weight: bold; }
 .none { color: #2e7d32; }
</style>
"@

$cols = "Username", "Name", "Department", "LastSignIn", "PasswordLastSet", "PasswordExpires", "OU", "Issues"

$summaryRows = foreach ($k in $sections.Keys) { [pscustomobject]@{ Check = $k; Accounts = $sections[$k].Count } }
$body  = "<h1>User Account Report - $((Get-ADDomain).DNSRoot)</h1>"
$body += "<p class='sub'>Generated $($Now.ToString('yyyy-MM-dd HH:mm')) by $env:USERNAME on $env:COMPUTERNAME | $($rows.Count) accounts checked</p>"
$body += ($summaryRows | ConvertTo-Html -Fragment -PreContent "<h2>Summary</h2>") -replace "<table>", "<table class='summary'>"

foreach ($k in $sections.Keys) {
    if ($sections[$k].Count) {
        $body += $sections[$k] | Select-Object $cols | ConvertTo-Html -Fragment -PreContent "<h2>$k ($($sections[$k].Count))</h2>"
    }
    else {
        $body += "<h2>$k (0)</h2><p class='none'>None</p>"
    }
}
$body += $rows | Select-Object $cols | ConvertTo-Html -Fragment -PreContent "<h2>All accounts ($($rows.Count))</h2>"

ConvertTo-Html -Title "User Account Report" -Head $css -Body $body | Out-File -FilePath $HtmlFile -Encoding utf8

# ---------------------------------------------------------------------------
# Console summary
# ---------------------------------------------------------------------------
Write-Host "`nUser Account Report - $($rows.Count) accounts checked" -ForegroundColor Cyan
$summaryRows | Format-Table -AutoSize | Out-Host
foreach ($k in "Locked out", "Disabled") {
    if ($sections[$k].Count) {
        Write-Host "$k`:" -ForegroundColor Yellow
        $sections[$k] | Select-Object Username, Name, LastSignIn, OU | Format-Table -AutoSize | Out-Host
    }
}
Write-Host "CSV:  $CsvFile"
Write-Host "HTML: $HtmlFile"

if ($Open) { Invoke-Item $HtmlFile }
