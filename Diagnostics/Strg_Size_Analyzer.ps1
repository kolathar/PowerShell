

# ============================================================
# Folder & File Size Analyzer
# Default Folders Size >= 5GB | Default Files Size >= 1GB
# Output: C:\Software\FolderAnalysis{MonthYear}_{ComputerName}.txt
# ============================================================
#Storage Size Analyzer 3.1 => [Standard Naming Scheme] - Storage Sizes Analyzer 1.2
# 2.0 - Changed size to be more precise Folder at 5GB (from 10), and File at 1Gb (from 5)
# 3.0 - Add a resparce eliminator , updated output file naming


#=========================================================================




Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# --- Configuration ---
$MinFolderSizeGB    = 5
$MinFileSizeGB      = 1
$MinFolderSizeBytes = $MinFolderSizeGB * 1GB
$MinFileSizeBytes   = $MinFileSizeGB * 1GB
$DateStamp          = Get-Date -Format "MMMM yyyy"
$TimeStamp          = Get-Date -Format "HHmm"
$ComputerName       = $env:COMPUTERNAME

# --- Global counters ---
$script:FolderCount = 0
$script:FileCount   = 0

# --- Determine output folder ---
# --- Determine output folder ---
$OutputFolder = "C:\Software\Diagnostics\Storage_Checks"

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}


$FileName   = "StrgAnalysis $DateStamp`_$TimeStamp`_$ComputerName.txt"
$OutputFile = "$OutputFolder\$FileName"

# --- Result collections ---
$script:FolderResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:FileResults   = [System.Collections.Generic.List[PSCustomObject]]::new()

# --- Function: Get total folder size ---
function Get-FolderSize {
    param ([string]$Path)
    try {
        $size = (Get-ChildItem -Path $Path -Recurse -File -Force -Attributes !ReparsePoint,!Offline -ErrorAction SilentlyContinue |
                 Measure-Object -Property Length -Sum).Sum
        return [long]($size)
    } catch {
        return 0L
    }
}

# --- Function: Get shortened file path (last 2 folders + filename) ---
# Root-level files keep full path
function Get-ShortFilePath {
    param ([string]$FullPath)
    $parts = $FullPath -split '\\'
    if ($parts.Count -le 2) {
        return $FullPath
    } elseif ($parts.Count -ge 4) {
        return "...\" + $parts[-3] + "\" + $parts[-2] + "\" + $parts[-1]
    } else {
        return "...\" + $parts[-2] + "\" + $parts[-1]
    }
}

# --- Function: Recursively scan folders >= $MinFolderSizeGB and files >= $MinFileSizeGB ---
function Search-LargeItems {
    param ([string]$Path)

    try {
        $subFolders = Get-ChildItem -Path $Path -Directory -Force -Attributes !ReparsePoint,!Offline -ErrorAction SilentlyContinue
    } catch {
        return
    }

    # Check subfolders
    foreach ($folder in $subFolders) {
        Write-Host "  Checking folder: $($folder.FullName)" -ForegroundColor DarkGray
        $size = Get-FolderSize -Path $folder.FullName

        if ($size -ge $MinFolderSizeBytes) {
            $script:FolderResults.Add([PSCustomObject]@{
                Drive   = $folder.FullName.Substring(0, 3)
                Path    = $folder.FullName
                SizeGB  = [math]::Round($size / 1GB, 2)
                SizeRaw = $size
            })
            $script:FolderCount += 1
            Write-Host "  [FOLDER MATCH] $($folder.FullName) - $([math]::Round($size / 1GB, 2)) GB" -ForegroundColor Green
            Search-LargeItems -Path $folder.FullName
        }
    }

    # Check files in current folder
    try {
        $files = Get-ChildItem -Path $Path -File -Force -Attributes !ReparsePoint,!Offline -ErrorAction SilentlyContinue
    } catch {
        return
    }

    foreach ($file in $files) {
        if ($file.Length -ge $MinFileSizeBytes) {
            $script:FileResults.Add([PSCustomObject]@{
                Drive     = $file.FullName.Substring(0, 3)
                Path      = $file.FullName
                ShortPath = Get-ShortFilePath -FullPath $file.FullName
                SizeGB    = [math]::Round($file.Length / 1GB, 2)
                SizeRaw   = $file.Length
            })
            $script:FileCount += 1
            Write-Host "  [FILE MATCH] $($file.FullName) - $([math]::Round($file.Length / 1GB, 2)) GB" -ForegroundColor Magenta
        }
    }
}

# --- Main Execution ---
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Folder & File Size Analyzer" -ForegroundColor Cyan
Write-Host "  Folders >= $MinFolderSizeGB GB | Files >= $MinFileSizeGB GB" -ForegroundColor Cyan
Write-Host "  Computer : $ComputerName" -ForegroundColor Cyan
Write-Host "  Output   : $OutputFile" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Scanning all local drives. This may take several minutes..." -ForegroundColor Yellow
Write-Host ""

# --- Get only local fixed drives (DriveType 3 = local fixed, excludes network) ---
$Drives = Get-WmiObject -Class Win32_LogicalDisk |
          Where-Object { $_.DriveType -eq 3 } |
          Select-Object -ExpandProperty DeviceID |
          Sort-Object

foreach ($drive in $Drives) {
    $drivePath = $drive + "\"
    Write-Host "----------------------------------------" -ForegroundColor Cyan
    Write-Host "Scanning Drive: $drivePath" -ForegroundColor Green
    Write-Host "----------------------------------------" -ForegroundColor Cyan
    Search-LargeItems -Path $drivePath
}

# --- Sort folders and files alphabetically by path, grouped by drive ---
$SortedFolders = $script:FolderResults | Sort-Object -Property Drive, Path
$SortedFiles   = $script:FileResults   | Sort-Object -Property Drive, Path

# --- Combine into one unified sorted list ---
$CombinedResults = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($folder in $SortedFolders) {
    $CombinedResults.Add([PSCustomObject]@{
        Type      = "Folder"
        Drive     = $folder.Drive
        Path      = $folder.Path
        ShortPath = $folder.Path
        SizeGB    = $folder.SizeGB
        SizeRaw   = $folder.SizeRaw
    })
}

foreach ($file in $SortedFiles) {
    $CombinedResults.Add([PSCustomObject]@{
        Type      = "File"
        Drive     = $file.Drive
        Path      = $file.Path
        ShortPath = $file.ShortPath
        SizeGB    = $file.SizeGB
        SizeRaw   = $file.SizeRaw
    })
}

$CombinedResults = $CombinedResults | Sort-Object -Property Drive, Path

# --- Calculate top-level folders (not a subfolder of another qualifying folder) ---
$TopLevelFolders = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($folder in $SortedFolders) {
    $isChild = $false
    foreach ($other in $SortedFolders) {
        if ($folder.Path -ne $other.Path -and $folder.Path.StartsWith($other.Path + "\")) {
            $isChild = $true
            break
        }
    }
    if (-not $isChild) {
        $TopLevelFolders.Add($folder)
    }
}

# --- Standalone files not inside a top-level qualifying folder ---
$StandaloneFiles = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($file in $SortedFiles) {
    $insideFolder = $false
    foreach ($folder in $TopLevelFolders) {
        if ($file.Path.StartsWith($folder.Path + "\")) {
            $insideFolder = $true
            break
        }
    }
    if (-not $insideFolder) {
        $StandaloneFiles.Add($file)
    }
}

# --- Total Data Found ---
$TotalDataBytes = (($TopLevelFolders | Measure-Object -Property SizeRaw -Sum).Sum) +
                  (($StandaloneFiles  | Measure-Object -Property SizeRaw -Sum).Sum)
$TotalDataGB    = [math]::Round($TotalDataBytes / 1GB, 2)

# --- Per-drive totals ---
$DriveBreakdown = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllDrives = ($TopLevelFolders | Select-Object -ExpandProperty Drive) +
             ($StandaloneFiles  | Select-Object -ExpandProperty Drive) |
             Sort-Object -Unique

foreach ($drv in $AllDrives) {
    $drvFolderBytes = ($TopLevelFolders | Where-Object { $_.Drive -eq $drv } |
                       Measure-Object -Property SizeRaw -Sum).Sum
    $drvFileBytes   = ($StandaloneFiles  | Where-Object { $_.Drive -eq $drv } |
                       Measure-Object -Property SizeRaw -Sum).Sum
    $drvTotal       = [math]::Round(($drvFolderBytes + $drvFileBytes) / 1GB, 2)
    $drvLetter      = $drv.TrimEnd('\').TrimEnd(':')
    $DriveBreakdown.Add([PSCustomObject]@{
        Drive   = $drv
        Letter  = $drvLetter
        TotalGB = $drvTotal
    })
}

# --- Display results in console ---
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  RESULTS" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

if ($CombinedResults.Count -eq 0) {
    Write-Host "No items found above the minimum thresholds." -ForegroundColor Red
} else {
    $currentDrive = ""
    foreach ($item in $CombinedResults) {
        if ($item.Drive -ne $currentDrive) {
            $currentDrive = $item.Drive
            Write-Host ""
            Write-Host "[ Drive: $currentDrive ]" -ForegroundColor Yellow
        }
        if ($item.Type -eq "Folder") {
            $line = "[DIR]  " + $item.Path.PadRight(80) + "  " + $item.SizeGB + " GB"
            Write-Host $line -ForegroundColor White
        } else {
            $line = "[FILE] " + $item.ShortPath.PadRight(80) + "  " + $item.SizeGB + " GB"
            Write-Host $line -ForegroundColor Magenta
        }
    }
}

Write-Host ""
Write-Host "Total Folders Found  : $($script:FolderCount)" -ForegroundColor Cyan
Write-Host "Total Files Found    : $($script:FileCount)" -ForegroundColor Cyan
Write-Host "Total Data Found     : $TotalDataGB GB" -ForegroundColor Cyan
foreach ($d in $DriveBreakdown) {
    Write-Host "    Drive $($d.Letter)          : $($d.TotalGB) GB" -ForegroundColor Cyan
}

# --- Build and write output file ---
$Output = [System.Collections.Generic.List[string]]::new()

$Output.Add("========================================")
$Output.Add("  Folder & File Size Analyzer - Results")
$Output.Add("  Computer             : $ComputerName")
$Output.Add("  Minimum Folder Size  : $MinFolderSizeGB GB")
$Output.Add("  Minimum File Size    : $MinFileSizeGB GB")
$Output.Add("  Scan Date            : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$Output.Add("  Total Folders Found  : $($script:FolderCount)")
$Output.Add("  Total Files Found    : $($script:FileCount)")
$Output.Add("  Total Data Found     : $TotalDataGB GB")
foreach ($d in $DriveBreakdown) {
    $Output.Add("      Drive " + $d.Letter + "          : " + $d.TotalGB + " GB")
}
$Output.Add("========================================")
$Output.Add("")

$currentDrive = ""
foreach ($item in $CombinedResults) {
    if ($item.Drive -ne $currentDrive) {
        $currentDrive = $item.Drive
        $Output.Add("")
        $Output.Add("[ Drive: $currentDrive ]")
    }
    if ($item.Type -eq "Folder") {
        $Output.Add("[DIR]  " + $item.Path.PadRight(80) + "  " + $item.SizeGB)
    } else {
        $Output.Add("[FILE] " + $item.ShortPath.PadRight(80) + "  " + $item.SizeGB)
    }
}

$Output | Out-File -FilePath $OutputFile -Encoding UTF8 -Force

Write-Host ""
Write-Host "Results saved to: $OutputFile" -ForegroundColor Green
Write-Host ""