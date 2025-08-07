#Author: Dharmik Pandya
#Purpose: This script checks which library/folder/file has unique permissions. Exports the output in a CSV file. Logs what was scanned and console output.
#Pre-Req: The account running the script must be a Site Collection Admin. Need PnP module. Works best with PowerShell 7.

# Connect to SharePoint
$siteUrl = "https://hotkilns.sharepoint.com/sites/Shop/"
Connect-PnPOnline -Url $siteUrl -ClientId 'ENTER CLIENT ID HERE' -Interactive

# Prepare files
$report = @()
$errorLogPath = "PermissionErrors.log"
$structureLogPath = "ScannedStructure.log"
if (Test-Path $errorLogPath) { Remove-Item $errorLogPath }
if (Test-Path $structureLogPath) { Remove-Item $structureLogPath }

# Progress tracking
$totalItems = 0
$currentItem = 0

# STEP 1: Count total items
Write-Host "`nEstimating total items to scan..."
$lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }

foreach ($list in $lists) {
    $rootFolderUrl = $list.RootFolder.ServerRelativeUrl
    $siteRelativeRoot = $rootFolderUrl -replace "^/sites/[^/]+", ""

    $stack = New-Object System.Collections.ArrayList
    [void]$stack.Add($siteRelativeRoot)

    while ($stack.Count -gt 0) {
        $url = $stack[$stack.Count - 1]
        $stack.RemoveAt($stack.Count - 1)
        $totalItems++

        try {
            $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $url -ItemType Folder -ErrorAction Stop
            foreach ($f in $folders) {
                [void]$stack.Add("$url/$($f.Name)")
            }
        } catch {}

        try {
            $files = Get-PnPFolderItem -FolderSiteRelativeUrl $url -ItemType File -ErrorAction Stop
            $totalItems += $files.Count
        } catch {}
    }
}

# STEP 2: Scan with minimal output and clear % progress bar
$totalLibraries = $lists.Count
$currentLibrary = 0

foreach ($list in $lists) {
    $currentLibrary++
    Write-Host "`nScanning Library [$currentLibrary of $totalLibraries]: $($list.Title)" -ForegroundColor Cyan

    $rootFolderUrl = $list.RootFolder.ServerRelativeUrl
    $siteRelativeRoot = $rootFolderUrl -replace "^/sites/[^/]+", ""

    # Check library-level permissions
    try {
        [void](Get-PnPProperty -ClientObject $list -Property "HasUniqueRoleAssignments")
        if ($list.HasUniqueRoleAssignments) {
            $report += [PSCustomObject]@{
                Type     = "Library"
                Path     = $rootFolderUrl
                Library  = $list.Title
                UniquePermissions = "Yes"
            }
        }
    } catch {
        Add-Content -Path $errorLogPath -Value "Failed to check library: $($list.Title) - $($_.Exception.Message)"
    }

    $stack = New-Object System.Collections.ArrayList
    [void]$stack.Add(@{ Url = $siteRelativeRoot; Depth = 1 })

    while ($stack.Count -gt 0) {
        $current = $stack[$stack.Count - 1]
        $stack.RemoveAt($stack.Count - 1)

        $url = $current.Url
        $depth = $current.Depth
        $indent = " " * ($depth * 2)

        $currentItem++
        $percent = [math]::Round(($currentItem / $totalItems) * 100, 2)
        Write-Progress -Activity "Scanning SharePoint Libraries ($currentLibrary of $totalLibraries)" `
                       -Status "Progress: $percent% - $currentItem of $totalItems items" `
                       -PercentComplete $percent

        Add-Content -Path $structureLogPath -Value "$indent $url"

        # Folder permission check
        try {
            $folder = Get-PnPFolder -Url $url -Includes ListItemAllFields.HasUniqueRoleAssignments -ErrorAction Stop
            if ($folder.ListItemAllFields.HasUniqueRoleAssignments) {
                $report += [PSCustomObject]@{
                    Type     = "Folder"
                    Path     = $url
                    Library  = $list.Title
                    UniquePermissions = "Yes"
                }
            }
        } catch {
            Add-Content -Path $errorLogPath -Value "Failed to get folder: $url - $($_.Exception.Message)"
        }

        # Subfolders
        try {
            $folders = Get-PnPFolderItem -FolderSiteRelativeUrl $url -ItemType Folder -ErrorAction Stop
            foreach ($f in $folders) {
                [void]$stack.Add(@{
                    Url = "$url/$($f.Name)"
                    Depth = $depth + 1
                })
            }
        } catch {
            Add-Content -Path $errorLogPath -Value "Failed to list subfolders at: $url - $($_.Exception.Message)"
        }

        # Files
        try {
            $files = Get-PnPFolderItem -FolderSiteRelativeUrl $url -ItemType File -ErrorAction Stop
            foreach ($file in $files) {
                $currentItem++
                $percent = [math]::Round(($currentItem / $totalItems) * 100, 2)
                Write-Progress -Activity "Scanning SharePoint Libraries ($currentLibrary of $totalLibraries)" `
                               -Status "Progress: $percent% - $currentItem of $totalItems items" `
                               -PercentComplete $percent

                Add-Content -Path $structureLogPath -Value "$indent   $($file.Name)"

                try {
                    $fileItem = Get-PnPFile -Url "$url/$($file.Name)" -AsListItem -ErrorAction Stop
                    [void](Get-PnPProperty -ClientObject $fileItem -Property "HasUniqueRoleAssignments")

                    if ($fileItem.HasUniqueRoleAssignments) {
                        $report += [PSCustomObject]@{
                            Type     = "File"
                            Path     = "$url/$($file.Name)"
                            Library  = $list.Title
                            UniquePermissions = "Yes"
                        }
                    }
                } catch {
                    Add-Content -Path $errorLogPath -Value "Failed to check file: $url/$($file.Name) - $($_.Exception.Message)"
                }
            }
        } catch {
            Add-Content -Path $errorLogPath -Value "Failed to list files at: $url - $($_.Exception.Message)"
        }
    }
}

# Export results
$report | Sort-Object Library, Path | Export-Csv -Path "Shop.csv" -NoTypeInformation

# Complete
Write-Progress -Activity "Completed" -Status "Finished scanning all libraries." -Completed
Write-Host "`nReport saved as 'Shop.csv'" -ForegroundColor Green
Write-Host "Structure logged in 'ScannedStructure.log'"
Write-Host "Any errors logged in 'PermissionErrors.log'"
