<#
.SYNOPSIS
    Find and optionally delete duplicate files in a folder and all subfolders.

.DESCRIPTION
    This script scans an input folder (and its subfolders) for duplicate files.
    A "duplicate" is defined as a file with the same name and extension as another 
    in the same folder, but with a numeric suffix in parentheses, e.g.:
        - song.mp3
        - song (1).mp3
        - song (2).mp3

    The script can:
        • Report duplicates
        • Run in debug mode to print all grouping decisions
        • Run in dry-run mode (-WhatIf) to preview deletions
        • Delete duplicates (keeping only the original without the (n) suffix)

.PARAMETER InputFolder
    The root folder to scan for duplicate files (includes all subfolders).

.PARAMETER DebugMode
    Prints extra information about how files are grouped.

.PARAMETER DeleteDuplicates
    Deletes duplicate files, keeping only the "original" one without suffix.

.PARAMETER WhatIf
    Works only with -DeleteDuplicates. Shows what would be deleted but does not delete.

.EXAMPLE
    # Only list duplicates
    .\Find-Duplicates.ps1 -InputFolder "C:\Users\alexg\OneDrive\Music"

.EXAMPLE
    # Show all debug information about grouping
    .\Find-Duplicates.ps1 -InputFolder "C:\Users\alexg\OneDrive\Music" -DebugMode

.EXAMPLE
    # Show what would be deleted (dry run, safe)
    .\Find-Duplicates.ps1 -InputFolder "C:\Users\alexg\OneDrive\Music" -DeleteDuplicates -WhatIf

.EXAMPLE
    # Actually delete duplicates
    .\Find-Duplicates.ps1 -InputFolder "C:\Users\alexg\OneDrive\Music" -DeleteDuplicates
#>

param (
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ })]
    [string]$InputFolder,

    # Show detailed logging
    [switch]$DebugMode,

    # Delete duplicate files (keep the first/original)
    [switch]$DeleteDuplicates,

    # Dry run mode (show what would be deleted without removing)
    [switch]$WhatIf
)

# Regex to match BaseName like: "file (1)" or "file(2)"
$dupBasePattern = '^(?<basename>.+?)\s*\((?<number>\d+)\)$'

# Collect all files recursively
$files = Get-ChildItem -Path $InputFolder -Recurse -File

if ($DebugMode) {
    Write-Host "Scanning: $InputFolder"
    Write-Host "Found $($files.Count) files`n"
}

# Group by: directory + (basename without (n)) + extension
$groups = @{}

foreach ($file in $files) {
    $dir  = $file.DirectoryName
    $base = $file.BaseName.Trim()
    $ext  = ($file.Extension ?? '').ToLowerInvariant()

    if ($base -match $dupBasePattern) {
        $rootBase = ($matches['basename']).Trim()
        $key = "$dir|$rootBase$ext"
        if ($DebugMode) { Write-Host "Duplicate candidate: $($file.Name)  → key: $key" }
    } else {
        $key = "$dir|$base$ext"
        if ($DebugMode) { Write-Host "Original candidate:  $($file.Name)  → key: $key" }
    }

    if (-not $groups.ContainsKey($key)) { $groups[$key] = @() }
    $groups[$key] += $file
}

# Show and optionally delete duplicates
$found = $false
foreach ($key in $groups.Keys) {
    $list = $groups[$key] | Sort-Object Name
    if ($list.Count -gt 1) {
        $found = $true
        Write-Host "`nDuplicate set in folder: $($list[0].DirectoryName)"
        foreach ($f in $list) {
            Write-Host "   $($f.Name)"
        }

        if ($DeleteDuplicates) {
            # Keep the first file (original), delete only suffixed ones
            $toDelete = $list | Where-Object { $_.BaseName -match $dupBasePattern }
            foreach ($f in $toDelete) {
                if ($WhatIf) {
                    Write-Host "Would delete: $($f.FullName)" -ForegroundColor Yellow
                } else {
                    Write-Host "Deleting: $($f.FullName)" -ForegroundColor Red
                    Remove-Item -LiteralPath $f.FullName -Force
                }
            }
        }
    }
}

if (-not $found) {
    Write-Host "No duplicates found."
}
