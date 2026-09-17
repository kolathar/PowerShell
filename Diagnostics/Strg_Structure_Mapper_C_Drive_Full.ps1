# ===========================================================================
# 
# ==============================  Notation ==================================
# 
# ===========================================================================

# Name of Program: Storage Structure Mapper - Full C Drive Version
# Creation Date: Sept 1, 2026
# Creator: Ethan Mathews

# 1.0 Creates a txt file with a map of the folder structure of the C drive, no exclusions. 
# 2.0 Creates the same but with Folder/File sizes added 
# 2.1 - Switched the Date and Time order filename from [Time] [Date] to [Date] [Time]
# 2.2 - Added an output to custom feild " storageAnalysisFullCompleted " with output: "$DateStamp_$TimeStamp". 
#        This is due to the script taking a lenghty amount of time to complete and helps verify completion. 
#        Also helpful when running in bulk; instead of checking each indiv device for completion, the tech 
#        can add the field to table view. 
#     - 2.2.3 > Updated output to "DateStamp $TimeStamp"
#     - 2.2.4 > Changed custom field to " strgAnalysisFull "
# 2.3 - Changed the output $CompleteDate/$CompleteTime, as DateStamp/TimeStamp log beginning values not end. 
# 2.4 - Changed OutputFoler to ""...\Drive_C" to distinguish from the comprehensive checker and FileName to 
#       Strg_Struct_DrC
# 3.0 - Added a section to elminiate bloating from reparse points



# ===========================================================================
# 
# ================================= Setup ===================================
# 
# ===========================================================================


#------------------------------ Base Parameters -----------------------------


param(
    [string]$RootPath = "C:\",
    [string]$ComputerName = $env:COMPUTERNAME
)


#------------------------------ Output Setup---------------------------------

$DateStamp = Get-Date -Format "MMddyyyy"
$TimeStamp = Get-Date -Format "HHmm"

$OutputFolder = "C:\Software\Diagnostics\Storage_Checks\Structure_Trees\Drive_C\"
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
$FileName   = "Strg_Struct_DrC_$DateStamp`_$TimeStamp`_$ComputerName.html"
$OutputFile = "$OutputFolder\$FileName"



#-------------------------Excluded Locations Section --------------------------


# Folders to skip entirely (case-insensitive, matched by name at any level)
$ExcludeNames = @()


# ===========================================================================
# 
# ============================ Main Functions ===============================
# 
# ===========================================================================



#------------------------------ Size Formatting ------------------------------

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
        $items = Get-ChildItem -LiteralPath $Path -ErrorAction Stop |
                  Where-Object { $ExcludeNames -notcontains $_.Name }
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

$rootResult = Get-TreeHtml -Path $RootPath
$bodyHtml   = $rootResult.Html
$rootSize   = Format-Size $rootResult.Size


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
  <h2>Structure map of $RootPath - $ComputerName - generated $(Get-Date)</h2>
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

$CompleteDate = Get-Date -Format "MMddyyyy"
$CompleteTime = Get-Date -Format "HHmm"


Ninja-Property-Set strgAnalysisFull "$CompleteDate $CompleteTime"



# ===========================================================================
# +++++++++++EM+++++++++++++++                   +++++++++++EM+++++++++++++++    
# ============================= END OF PROGRAM ==============================
# +++++++++++EM+++++++++++++++                   +++++++++++EM+++++++++++++++                        
# ===========================================================================
