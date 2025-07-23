<#
    Designed by: Dharmik Pandya
    Script: GroupOwnersManager.ps1

    Summary:
    - Connects to Microsoft Graph
    - Exports current owners of each Azure AD group listed in a CSV
    - Adds 3 specified owners to each group
    - Removes all other owners
    - Logs all actions and errors, then shows a summary

    CSV Format (groups.csv):
    GroupName
    1One1
    1Two1
    1Three1

    Notes:
    - Update the `$ownersToKeep` list with the 3 UPNs to keep as owners
    - Requires Microsoft Graph PowerShell SDK and appropriate permissions
#>

# === CONFIGURATION ===
$inputCsv   = "groups.csv"                             # Input CSV with GroupName column
$outputDir  = "owners"                                # Directory to export owners
$logFile    = "Group_Read_Log.txt"                    # Main log file
$errorFile  = "Group_Error_Log.txt"                   # Error log file

# === CLEANUP FROM LAST RUN ===
Remove-Item $logFile, $errorFile -ErrorAction SilentlyContinue           # Clean old logs
Start-Transcript -Path $logFile -Append                                 # Begin logging console output
if (-not (Test-Path $outputDir)) { New-Item $outputDir -ItemType Directory | Out-Null }  # Create output folder if missing

# === PART 1: Connect & Export Owners ===
Write-Progress -Activity "Script Execution" -Status "Step 1 of 3: Exporting Group Owners" -PercentComplete 33
Write-Host "`n== PART 1: Exporting Group Owners =="

# Connect to Microsoft Graph with necessary scopes
Write-Host "Connecting to Microsoft Graph..."
try {
    Connect-MgGraph -Scopes "Group.ReadWrite.All", "Directory.ReadWrite.All" -ErrorAction Stop
    Write-Host "Connected to Microsoft Graph."
} catch {
    "Connection Error: $($_.Exception.Message)" | Out-File $errorFile -Append
    Stop-Transcript
    exit
}

# Load group list from CSV
$groups = Import-Csv $inputCsv | Where-Object { $_.GroupName -ne "" }
$total = $groups.Count
if ($total -eq 0) {
    Write-Host "No group names found in CSV."
    Stop-Transcript
    exit
}

# Load all Azure AD groups once to avoid multiple API calls
$allGroups = Get-MgGroup -All

# Export current owners of each group to a CSV
$count = 1
foreach ($row in $groups) {
    $name = $row.GroupName
    $percent = [math]::Round(($count / $total) * 100)
    Write-Progress -Activity "Exporting Group Owners" -Status "$percent% Complete" -PercentComplete $percent
    Write-Host "`n[$count/$total] $name"

    try {
        $group = $allGroups | Where-Object { $_.DisplayName -eq $name }

        if (-not $group) {
            "Group not found: $name" | Out-File $errorFile -Append
        } else {
            $rawOwners = Get-MgGroupOwner -GroupId $group.Id
            $owners = @()

            foreach ($owner in $rawOwners) {
                try {
                    $user = Get-MgUser -UserId $owner.Id
                    $owners += [PSCustomObject]@{
                        Id                = $user.Id
                        DisplayName       = $user.DisplayName
                        UserPrincipalName = $user.UserPrincipalName
                    }
                } catch {
                    "Unable to fetch user details for owner ID: $($owner.Id)" | Out-File $errorFile -Append
                }
            }

            if ($owners.Count -gt 0) {
                $owners | Export-Csv "$outputDir\$name.csv" -NoTypeInformation
                Write-Host "Exported owners for: $name"
            } else {
                Write-Host "No owners for: $name"
            }
        }
    } catch {
        "Error with group ${name}: $($_.Exception.Message)" | Out-File $errorFile -Append
    }

    $count++
}

Write-Host "`n== PART 1 Complete =="


# === PART 2: Enforce Only 3 Owners per Group ===
Write-Progress -Activity "Script Execution" -Status "Step 2 of 3: Replacing Group Owners" -PercentComplete 66
Write-Host "`n== PART 2: Replacing Owners With Specified 3 =="

# UPNs of the 3 users you want to keep as owners in all groups
$ownersToKeep = @(
    "dharmik0417@szlk.onmicrosoft.com",
    "TShark@szlk.onmicrosoft.com",
    "PattiF@szlk.onmicrosoft.com"
)

$count = 1
foreach ($row in $groups) {
    $name = $row.GroupName
    $percent = [math]::Round(($count / $total) * 100)
    Write-Progress -Activity "Updating Owners" -Status "$percent% Complete" -PercentComplete $percent
    Write-Host "`n[$count/$total] $name"

    try {
        $group = $allGroups | Where-Object { $_.DisplayName -eq $name }

        if (-not $group) {
            "Group not found: $name" | Out-File $errorFile -Append
            continue
        }

        # Step 1: Add all 3 owners (safe to re-add if already assigned)
        foreach ($upn in $ownersToKeep) {
            try {
                $user = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction Stop | Select-Object -First 1
                New-MgGroupOwnerByRef -GroupId $group.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($user.Id)" } -ErrorAction Stop
                Write-Host " → Ensured owner: $($user.DisplayName) ($upn)"
            } catch {
                "Failed to add owner ${upn} to group ${name}: $($_.Exception.Message)" | Out-File -Append $errorFile
            }
        }

        # Step 2: Remove all current owners not in the keeper list
        $owners = Get-MgGroupOwner -GroupId $group.Id
        foreach ($owner in $owners) {
            try {
                $user = Get-MgUser -UserId $owner.Id -ErrorAction SilentlyContinue
                $upn = $user.UserPrincipalName

                if ($ownersToKeep -notcontains $upn) {
                    Remove-MgGroupOwnerByRef -GroupId $group.Id -DirectoryObjectId $owner.Id -ErrorAction Stop
                    Write-Host " → Removed owner: $($user.DisplayName) ($upn)"
                }
            } catch {
                "Failed to remove owner ID $($owner.Id) from group ${name}: $($_.Exception.Message)" | Out-File -Append $errorFile
            }
        }
    } catch {
        "Error processing group ${name}: $($_.Exception.Message)" | Out-File -Append $errorFile
    }

    $count++
}

Write-Host "`n== PART 2 Complete =="
Write-Progress -Activity "Script Execution" -Status "Step 2 Complete" -PercentComplete 66


# === PART 3: SUMMARY ===
Write-Progress -Activity "Script Execution" -Status "Step 3 of 3: Summary Report" -PercentComplete 100
Write-Host "`n== PART 3: Summary =="

$exportedCount = (Get-ChildItem "$outputDir\*.csv").Count
$errorCount = if (Test-Path $errorFile) { (Get-Content $errorFile | Where-Object { $_ -match '\S' }).Count } else { 0 }

Write-Host "`nSummary of operations:"
Write-Host " → Total groups processed  : $total"
Write-Host " → Owners exported         : $exportedCount"
Write-Host " → Errors encountered      : $errorCount"

Write-Host "`nAll steps complete. See logs for full details."

# === END ===
Stop-Transcript
