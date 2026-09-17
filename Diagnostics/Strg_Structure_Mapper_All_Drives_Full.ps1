# ===========================================================================
# 
# ==============================  Notation ==================================
# 
# ===========================================================================

# Name of Program: Storage Structure Mapper - All Disks Version
# Creation Date: Sept 11, 2026
# Creator: Ethan Mathews

# 1.0 - Creates a txt file with a map of the folder structure of the C drive, with exclusion (see script contents). 
#       Primarily used for SERVERS and MULTI-DRIVE computers > Otherwise, USE the "C Drive Full" version,
#       which targets C:\ exclusively. 
#       EXCLUDES OneDrive Folders by default, as this folder can be caught by other smaller scripts. 
# 1.2 - Changed Output Folder to ...\All Drives and FileName to Strg_Struct_ADr
# 2.0 - Added a section to elminiate bloating from reparse points
#       Added section in Get-ChildItem to properly handle Exceptions area wildcards
#       Edited some misspelled variables 







# ===========================================================================
# 
# ================================= Setup ===================================
# 
# ===========================================================================


#------------------------------ Base Parameters -----------------------------


param(
    [string]$ComputerName = $env:COMPUTERNAME
)


#-------------------------Local Drives Detection --------------------------

# Only local fixed disks (DriveType 3) — excludes mapped network drives (DriveType 4), CDs, removable, etc.
$LocalDrives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" |
               Select-Object -ExpandProperty DeviceID   # e.g. "C:", "D:", "E:"



#------------------------------ Output Setup---------------------------------

$DateStamp = Get-Date -Format "MMddyyyy"
$TimeStamp = Get-Date -Format "HHmm"

$OutputFolder = "C:\Software\Diagnostics\Storage_Checks\Structure_Trees\All_Drives"
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
$FileName   = "Strg_Struct_ADr_$DateStamp`_$TimeStamp`_$ComputerName.html"
$OutputFile = "$OutputFolder\$FileName"



#-------------------------Excluded Locations Section --------------------------


# Folders to skip entirely (case-insensitive, matched by name at any level)
$ExcludeNames = @(
  "OneDrive -*",
  "OneDrive"
  )




# ===========================================================================
# 
# ============================ Main Functions ===============================
# 
# ===========================================================================



function Format-Size {
    param([long]$Bytes)

    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    else { return "$Bytes B" }
}

function Get-TreeHtml {
    param(
        [string]$Path
    )

    try {
        $items = Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop |
                  Where-Object {
                    $name = $_.Name
                    -not ($ExcludeNames | Where-Object { $name -like $_ })
                  }
    } catch {
        return [PSCustomObject]@{ Html = "<li>[ACCESS DENIED] $Path</li>"; Size = 0 }
    }

    $html = ""
    $totalSize = 0

    foreach ($item in $items) {
        if ($item.PSIsContainer) {
          
            # Skip reparse points (mount points, junctions, symlinks) — don't recurse into them
            if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                $html += "<li><em>$($item.Name) (reparse point / mount — not counted)</em></li>"
                continue
            }
            
            $child = Get-TreeHtml -Path $item.FullName
            $sizeLabel = Format-Size $child.Size
            $html += "<li><details><summary>$($item.Name) ($sizeLabel)</summary><ul>$($child.Html)</ul></details></li>"
            $totalSize += $child.Size
            
        } else {
          
            $fileSize = $item.Length
            $sizeLabel = Format-Size $fileSize
            $html += "<li class='file'>$($item.Name) ($sizeLabel)</li>"
            $totalSize += $fileSize
        }
    }

    return [PSCustomObject]@{ Html = $html; Size = $totalSize }
}



# ===========================================================================
# 
# ======================= Building the HTML File ============================
# 
# ===========================================================================



$allDrivesHtml = ""
$grandTotalSize = 0

foreach ($drive in $LocalDrives) {
    $rootPath = "$drive\"
    $driveResult = Get-TreeHtml -Path $rootPath
    $driveSizeLabel = Format-Size $driveResult.Size
    $grandTotalSize += $driveResult.Size

    $allDrivesHtml += "<li><details open><summary>$rootPath ($driveSizeLabel)</summary><ul>$($driveResult.Html)</ul></details></li>"
}

$bodyHtml = $allDrivesHtml
$rootSize = Format-Size $grandTotalSize


$htmlDoc = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Structure Map - $ComputerName - $DateStamp</title>
<style>
  body { font-family: Consolas, monospace; font-size: 14px; }
  ul { list-style-type: none; padding-left: 20px; }
  summary { cursor: pointer; font-weight: bold; }
  li.file { color: #444; }
  #search { margin-bottom: 15px; padding: 6px; width: 300px; font-size: 14px; }
</style>
</head>
<body>
  <h2>Structure map of all local drives - $ComputerName - generated $(Get-Date) - Total size: $rootSize</h2>
  <input type="text" id="search" placeholder="Type to highlight matches...">
  <ul>$bodyHtml</ul>

  <script>
    document.getElementById('search').addEventListener('input', function() {
      const term = this.value.toLowerCase();
      const items = document.querySelectorAll('li');
      items.forEach(li => {
        const summary = li.querySelector(':scope > details > summary');
        const label = summary ? summary.textContent : li.textContent;

        if (term && label.toLowerCase().includes(term)) {
          li.style.backgroundColor = 'yellow';
          let el = li.closest('details');
          while (el) { el.open = true; el = el.parentElement.closest('details'); }
        } else {
          li.style.backgroundColor = '';
        }
      });
    });
  </script>
</body>
</html>
"@




# ===========================================================================
# 
# =========================== Write Outs ===============================
# 
# ===========================================================================


Set-Content -Path $OutputFile -Value $htmlDoc -Encoding UTF8

Write-Host "Structure map written to $OutputFile"

$OutputDate = Get-Date -Format "MMddyyyy"
$OutputTime = Get-Date -Format "HHmm"

Ninja-Property-Set strgAnalysisAllDrives "$OutputDate $OutputTime"





# ===========================================================================
# +++++++++++EM+++++++++++++++                   +++++++++++EM+++++++++++++++    
# ============================= END OF PROGRAM ==============================
# +++++++++++EM+++++++++++++++                   +++++++++++EM+++++++++++++++                        
# ===========================================================================
