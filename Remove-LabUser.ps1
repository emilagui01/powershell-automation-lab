<#
.SYNOPSIS
    Offboards departing users: disables, strips access, archives, and documents.

.DESCRIPTION
    For each username, the script:
      1. Refuses to touch protected accounts (Administrator, Domain Admins, service accounts)
      2. Records the user's group memberships in the log (so access could be restored)
      3. Disables the account and resets the password to a random value
      4. Removes the user from every group except Domain Users
      5. Stamps the Description with the date, who did it, and the ticket number
      6. Moves the account to the Disabled Users OU (outside Lab Users, so no GPOs apply)
      7. Moves the home folder to an admin-only archive and clears the H: drive mapping

    Nothing is deleted. Supports -WhatIf to preview every change.

.PARAMETER Identity
    One or more usernames (SamAccountName).

.PARAMETER Ticket
    Help desk ticket number for the audit trail, e.g. 482915.

.EXAMPLE
    .\Remove-LabUser.ps1 -Identity slang -Ticket 482915 -WhatIf

.EXAMPLE
    "slang","wmaximoff" | .\Remove-LabUser.ps1 -Ticket 482915

.NOTES
    Author: Emil Aguirre Herrera
    Lab:    Home Lab 3 - PowerShell Automation (emillab.local)
    Run on DC01 (home folders are moved using local paths).
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, ValueFromPipeline)]
    [string[]]$Identity,

    [Parameter(Mandatory)]
    [string]$Ticket,

    [string]$LogDir      = "C:\Scripts\Lab3\Logs",
    [string]$DisabledOU  = "OU=Disabled Users,DC=emillab,DC=local",
    [string]$HomeLocal   = "C:\Shares\Home",
    [string]$ArchiveRoot = "C:\Shares\Archive"
)

begin {
    Import-Module ActiveDirectory -ErrorAction Stop

    $Protected  = @("Administrator", "krbtgt", "Guest")
    $Stamp      = Get-Date -Format "yyyyMMdd-HHmmss"
    $Today      = Get-Date -Format "yyyy-MM-dd"
    $LogFile    = Join-Path $LogDir "Remove-LabUser-$Stamp.log"
    $ResultFile = Join-Path $LogDir "Remove-LabUser-$Stamp-results.csv"
    $Results    = New-Object System.Collections.Generic.List[object]

    if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -WhatIf:$false | Out-Null }

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
        param($Username, $Name, $Status, $Detail)
        $Results.Add([pscustomobject]@{ Username = $Username; Name = $Name; Status = $Status; Detail = $Detail })
    }

    function New-RandomPassword {
        # Long random password nobody knows, so the account can't be used even if re-enabled by mistake
        $chars = [char[]]"ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%*?"
        return "Xx9!" + (-join (1..28 | ForEach-Object { $chars | Get-Random }))
    }

    # Pre-flight: the Disabled Users OU and archive folder must exist
    try { Get-ADOrganizationalUnit -Identity $DisabledOU -ErrorAction Stop | Out-Null }
    catch { throw "Disabled Users OU not found: $DisabledOU. Create it first." }
    if (-not (Test-Path $ArchiveRoot)) { throw "Archive folder not found: $ArchiveRoot. Create it first." }

    Write-Log "Run started by $env:USERNAME | Ticket #$Ticket | WhatIf: $WhatIfPreference"
}

process {
    foreach ($sam in $Identity) {
        $sam = $sam.Trim()

        # --- Find the user ---------------------------------------------------
        $user = Get-ADUser -Filter "SamAccountName -eq '$($sam -replace "'", "''")'" -Properties MemberOf, HomeDirectory, Description, Enabled
        if (-not $user) {
            Write-Log "SKIP $sam`: no such user in AD" WARN
            Add-Result $sam "" "Skipped" "User not found"
            continue
        }
        $name = $user.Name

        # --- Safety checks -----------------------------------------------------
        $groups = @($user.MemberOf | ForEach-Object { (Get-ADGroup $_).Name })
        if ($sam -in $Protected -or $groups -contains "Domain Admins" -or $user.DistinguishedName -like "*OU=Service Accounts*") {
            Write-Log "REFUSED $sam ($name): protected account (admin, built-in or service account)" ERROR
            Add-Result $sam $name "Refused" "Protected account"
            continue
        }
        if ($user.DistinguishedName -like "*$DisabledOU") {
            Write-Log "SKIP $sam ($name): already offboarded" WARN
            Add-Result $sam $name "Skipped" "Already in Disabled Users"
            continue
        }

        $groupList  = if ($groups.Count) { $groups -join ", " } else { "(none)" }
        $archiveDir = Join-Path $ArchiveRoot "$sam-$Today"
        $homeDir    = Join-Path $HomeLocal $sam

        Write-Log "$sam ($name) | groups before: $groupList"

        if (-not $PSCmdlet.ShouldProcess("$name ($sam)", "Disable, reset password, remove from groups, move to Disabled Users, archive home folder")) {
            Write-Log "WHATIF $sam would be disabled, removed from: $groupList, moved to Disabled Users, home folder -> $archiveDir"
            Add-Result $sam $name "WhatIf" "Would be offboarded (groups: $groupList)"
            continue
        }

        # --- Offboard ----------------------------------------------------------
        try {
            Disable-ADAccount -Identity $user -ErrorAction Stop
            Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString (New-RandomPassword) -AsPlainText -Force) -ErrorAction Stop
            Write-Log "OK   $sam disabled and password reset" OK

            foreach ($g in $groups) {
                Remove-ADGroupMember -Identity $g -Members $user -Confirm:$false -ErrorAction Stop
            }
            Write-Log "OK   $sam removed from $($groups.Count) group(s)" OK

            Set-ADUser -Identity $user -Description "Disabled $Today by $env:USERNAME - Ticket #$Ticket" -HomeDirectory $null -HomeDrive $null -ErrorAction Stop
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $DisabledOU -ErrorAction Stop
            Write-Log "OK   $sam moved to Disabled Users, description stamped" OK

            if (Test-Path $homeDir) {
                Move-Item -Path $homeDir -Destination $archiveDir -ErrorAction Stop
                Write-Log "OK   home folder archived to $archiveDir" OK
                $homeNote = "home -> $archiveDir"
            }
            else {
                Write-Log "     no home folder at $homeDir" INFO
                $homeNote = "no home folder"
            }

            Add-Result $sam $name "Offboarded" "Removed from: $groupList; $homeNote"
        }
        catch {
            Write-Log "FAIL $sam ($name): $($_.Exception.Message)" ERROR
            Add-Result $sam $name "Failed" $_.Exception.Message
        }
    }
}

end {
    $Results | Export-Csv -Path $ResultFile -NoTypeInformation -WhatIf:$false
    $summary = ($Results | Group-Object Status | ForEach-Object { "$($_.Name): $($_.Count)" }) -join " | "
    Write-Log "Run finished | $summary"
    Write-Log "Log: $LogFile"
    $Results | Format-Table -AutoSize
}
