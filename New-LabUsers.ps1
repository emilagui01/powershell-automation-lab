<#
.SYNOPSIS
    Bulk-creates Active Directory user accounts from a CSV of new hires.

.DESCRIPTION
    For each row in the CSV, the script:
      1. Validates the row (name present, department exists)
      2. Skips anyone who already has an account (safe to re-run)
      3. Generates a unique username (first initial + last name, with a number if taken)
      4. Creates the account in the department OU with a random temporary password
         and "must change password at next logon"
      5. Adds the user to their department group and to Lab Staff
      6. Creates a personal home folder on \\DC01\Home$ that only they (and admins) can open
      7. Logs every action and exports a results CSV

    Supports -WhatIf to preview every change without making it.

.PARAMETER CsvPath
    CSV with columns: FirstName, LastName, Department, Title

.EXAMPLE
    .\New-LabUsers.ps1 -CsvPath .\NewHires.csv -WhatIf
    Dry run: shows what would be created.

.EXAMPLE
    .\New-LabUsers.ps1 -CsvPath .\NewHires.csv
    Creates the accounts.

.EXAMPLE
    .\New-LabUsers.ps1 -CsvPath .\NewHires-520879.csv -Ticket 520879
    Creates the accounts and stamps the ticket number on each one's Description.

.NOTES
    Author: Emil Aguirre Herrera
    Lab:    Home Lab 3 - PowerShell Automation (emillab.local)
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [string]$Ticket,                                   # help desk ticket number, stamped on each account

    [string]$LogDir   = "C:\Scripts\Lab3\Logs",
    [string]$Domain    = "emillab.local",
    [string]$BaseOU    = "OU=Lab Users,DC=emillab,DC=local",
    [string]$HomeRoot  = "\\DC01\Home$",
    [string]$HomeDrive = "H:"
)

Import-Module ActiveDirectory -ErrorAction Stop

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
$Departments = @("IT", "HR", "Finance", "Marketing")   # must match the OUs and GRP- groups
$StaffGroup  = "Lab Staff"                             # gives S: drive and Remote Desktop access (Lab 1 GPOs)
$NetBios     = (Get-ADDomain).NetBIOSName              # EMILLAB
$Stamp       = Get-Date -Format "yyyyMMdd-HHmmss"
$LogFile     = Join-Path $LogDir "New-LabUsers-$Stamp.log"
$ResultFile  = Join-Path $LogDir "New-LabUsers-$Stamp-results.csv"
$HandoffFile = Join-Path $LogDir "New-LabUsers-$Stamp-HANDOFF.csv"

$Results  = New-Object System.Collections.Generic.List[object]
$Handoff  = New-Object System.Collections.Generic.List[object]
$Reserved = New-Object System.Collections.Generic.HashSet[string]   # usernames claimed during this run

# Logging must happen even during -WhatIf, so these calls use -WhatIf:$false
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -WhatIf:$false | Out-Null }

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO", "OK", "WARN", "ERROR")][string]$Level = "INFO"
    )
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $line -WhatIf:$false
    $color = @{ INFO = "Gray"; OK = "Green"; WARN = "Yellow"; ERROR = "Red" }[$Level]
    Write-Host $line -ForegroundColor $color
}

function Add-Result {
    param($Name, $Username, $Department, $Status, $Detail)
    $Results.Add([pscustomobject]@{
        Name       = $Name
        Username   = $Username
        Department = $Department
        Status     = $Status
        Detail     = $Detail
    })
}

function Get-UniqueSamAccountName {
    # First initial + last name, lowercase, letters and numbers only, max 20 characters.
    # If taken (in AD or earlier in this run), append 2, 3, 4...
    param([string]$First, [string]$Last)
    $base = (($First.Substring(0, 1) + $Last).ToLower()) -replace '[^a-z0-9]', ''
    if ($base.Length -gt 18) { $base = $base.Substring(0, 18) }

    $sam = $base
    $n   = 1
    while ($Reserved.Contains($sam) -or (Get-ADUser -Filter "SamAccountName -eq '$sam'")) {
        $n++
        $sam = "$base$n"
    }
    [void]$Reserved.Add($sam)
    return $sam
}

function New-TempPassword {
    # 12 characters with at least one upper, lower, digit and symbol, so it always meets
    # domain complexity. (If New-ADUser's password is rejected, the account is still
    # created but left disabled - see Lab 1, troubleshooting.)
    # Ambiguous characters (0/O, 1/l/I) are left out so it can be read aloud.
    $upper = "ABCDEFGHJKLMNPQRSTUVWXYZ"
    $lower = "abcdefghijkmnopqrstuvwxyz"
    $digit = "23456789"
    $sym   = "!@#$%*?"
    $all   = $upper + $lower + $digit + $sym

    $chars = @(
        $upper[(Get-Random -Maximum $upper.Length)]
        $lower[(Get-Random -Maximum $lower.Length)]
        $digit[(Get-Random -Maximum $digit.Length)]
        $sym[(Get-Random -Maximum $sym.Length)]
    )
    $chars += 1..8 | ForEach-Object { $all[(Get-Random -Maximum $all.Length)] }
    return -join ($chars | Sort-Object { Get-Random })
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$rows = @(Import-Csv -Path $CsvPath)
Write-Log "Run started by $env:USERNAME on $env:COMPUTERNAME | $($rows.Count) rows from $CsvPath | Ticket: $(if ($Ticket) { "#$Ticket" } else { 'none' }) | WhatIf: $WhatIfPreference"

foreach ($row in $rows) {
    $first   = "$($row.FirstName)".Trim()
    $last    = "$($row.LastName)".Trim()
    $dept    = "$($row.Department)".Trim()
    $title   = "$($row.Title)".Trim()
    $display = "$first $last"

    # --- Validate -----------------------------------------------------------
    if (-not $first -or -not $last) {
        Write-Log "SKIP '$display': missing first or last name" WARN
        Add-Result $display "" $dept "Skipped" "Missing first or last name"
        continue
    }
    if ($dept -notin $Departments) {
        Write-Log "SKIP $display`: unknown department '$dept' (valid: $($Departments -join ', '))" WARN
        Add-Result $display "" $dept "Skipped" "Unknown department '$dept'"
        continue
    }

    # --- Already has an account? (makes the script safe to run twice) -------
    # Same name = probably the same person. Skip and flag for a human
    # to review instead of creating a duplicate account.
    $f = $first -replace "'", "''"
    $l = $last  -replace "'", "''"
    $n = $display -replace "'", "''"
    # Checks Name too: accounts made by hand in Lab 1 have a Name but no GivenName/Surname
    $existing = @(Get-ADUser -Filter "(GivenName -eq '$f' -and Surname -eq '$l') -or Name -eq '$n'")
    if ($existing.Count -gt 0) {
        $who = ($existing | ForEach-Object { $_.SamAccountName }) -join ", "
        Write-Log "SKIP $display`: account already exists ($who) - review manually" WARN
        Add-Result $display $who $dept "Skipped" "Account already exists - review manually"
        continue
    }

    # --- Build the account details ------------------------------------------
    $sam     = Get-UniqueSamAccountName -First $first -Last $last
    $ou      = "OU=$dept,$BaseOU"
    $group   = "GRP-$dept"
    $homeDir = "$HomeRoot\$sam"
    $note    = if ($sam -ne ((($first.Substring(0, 1) + $last).ToLower()) -replace '[^a-z0-9]', '')) { " (base name taken)" } else { "" }

    if (-not $PSCmdlet.ShouldProcess("$display -> $sam in $ou", "Create user, add to $group and $StaffGroup, create $homeDir")) {
        Write-Log "WHATIF $display -> $sam$note | OU=$dept | groups: $group, $StaffGroup | home: $homeDir"
        Add-Result $display $sam $dept "WhatIf" "Would be created$note"
        continue
    }

    # --- Create ---------------------------------------------------------------
    try {
        $password = New-TempPassword
        $userParams = @{
            Name                  = $display
            GivenName             = $first
            Surname               = $last
            DisplayName           = $display
            SamAccountName        = $sam
            UserPrincipalName     = "$sam@$Domain"
            EmailAddress          = "$sam@$Domain"
            Title                 = $title
            Department            = $dept
            Path                  = $ou
            AccountPassword       = (ConvertTo-SecureString $password -AsPlainText -Force)
            ChangePasswordAtLogon = $true
            Enabled               = $true
            HomeDirectory         = $homeDir
            HomeDrive             = $HomeDrive
            ErrorAction           = "Stop"
        }
        if ($Ticket) { $userParams.Description = "Created $(Get-Date -Format 'yyyy-MM-dd') by $env:USERNAME - Ticket #$Ticket" }
        New-ADUser @userParams
        Write-Log "OK   created $display -> $sam$note in OU=$dept" OK

        Add-ADGroupMember -Identity $group      -Members $sam -ErrorAction Stop
        Add-ADGroupMember -Identity $StaffGroup -Members $sam -ErrorAction Stop
        Write-Log "OK   $sam added to $group, $StaffGroup" OK

        New-Item -Path $homeDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        $null = icacls $homeDir /grant "$NetBios\${sam}:(OI)(CI)M"
        if ($LASTEXITCODE -ne 0) { throw "icacls failed to grant $sam Modify on $homeDir" }
        Write-Log "OK   home folder $homeDir (Modify for $sam only)" OK

        Add-Result $display $sam $dept "Created" "OU=$dept; $group; $StaffGroup; $homeDir$note"
        $Handoff.Add([pscustomobject]@{ Name = $display; Username = $sam; TempPassword = $password })
    }
    catch {
        Write-Log "FAIL $display ($sam): $($_.Exception.Message)" ERROR
        Add-Result $display $sam $dept "Failed" $_.Exception.Message
    }
}

# ---------------------------------------------------------------------------
# Wrap up
# ---------------------------------------------------------------------------
$Results | Export-Csv -Path $ResultFile -NoTypeInformation -WhatIf:$false
if ($Handoff.Count -gt 0) {
    $Handoff | Export-Csv -Path $HandoffFile -NoTypeInformation -WhatIf:$false
    Write-Log "Temporary passwords saved to $HandoffFile - hand them over securely, then DELETE this file" WARN
}

$summary = ($Results | Group-Object Status | ForEach-Object { "$($_.Name): $($_.Count)" }) -join " | "
Write-Log "Run finished | $summary"
Write-Log "Log: $LogFile"
Write-Log "Results: $ResultFile"

$Results | Format-Table -AutoSize
