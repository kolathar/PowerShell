# Event Veiwer Reporter 1.2 => [Standard Naming Scheme] Event Viewer Reporter (48hrs - Sorted) (1.2)
# Examines criticals and errors in System and Application logs in Event Viewer and outputs to an  Excel file.  
# Must manually change time in $StartTime = (Get-Date).AddHours(-48)
# Creates a formatted Excel workbook without Excel, ImportExcel, NuGet, or external modules.

# 
# =========================== Tracking Changes ==============================
# 
# 1.3 - improved script formatting > added comments > section headers
# 2.0 - added a log file and logging
# 2.1 - improved output file > Sheet formatting > custom columns widths
#
# ~ Updates will now follow these rules of documentation in Ninja: 
#     first number changes with a new components (always #.0.0)
#     second number changes with a improved component
#     third number changes with each iteration saved 
#           As in, every time I save and ready a test run, I update this third 
#             number so that i can tell the difference between tests when 
#             scrolling through Ninja's activity log
#
#  2.2 - changed the columns of Applications to remove the message prefix and apply color the source cells
#           (sheets still use message prefix to sort and determine color)
#  2.3 - changed message prefix to be 30 instead of 15
#
#
#
#





#++++++++++++++++++++++++++++++ PROGRAM START +++++++++++++++++++++++++++++++



# ===========================================================================
# 
# =============================== Parameters ================================
# 
# ===========================================================================


$ErrorActionPreference = 'Stop'

$ReportRoot = 'C:\Software\Diagnostics\Event_Viewer_Reports'
$LogRoot = 'C:\Software\Diagnostics\Event_Viewer_Reports\logs'

$StartTime = (Get-Date).AddHours(-48)
$Timestamp = Get-Date -Format 'MMddyyyy_HHmm'

$Computer = $env:COMPUTERNAME
$OutputFile = Join-Path $ReportRoot ("{0}_EventVRep_{1}.xlsx" -f $Computer, $Timestamp)
$TempRoot = Join-Path $env:TEMP ("EventVRep_{0}" -f ([guid]::NewGuid().ToString()))

# --- Column Width Configuration (Excel character-width units) ---
$SysTopWidths     = @(40, 8, 20, 12, 12, 12, 150)   # Source, Count, RecentDate, EventID, Level, LogName, Message
$SysRawWidths     = @(20, 40, 12, 12, 12, 150)       # Date, Source, EventID, Level, LogName, Message
$SysSortedWidths  = @(20, 40, 12, 12, 12, 150)
$AppTopWidths     = @(40, 8, 20, 12, 12, 12, 150)    # MessagePrefix, Count, RecentDate, EventID, Level, LogName, Message
$AppRawWidths     = @(20, 30, 12, 12, 12, 150)   # Date, MessagePrefix, Source, EventID, Level, LogName, Message
$AppSortedWidths  = @(20, 30, 12, 12, 12, 150)


$LightColors = @(
    'FCE4D6','E2F0D9','DDEBF7','FFF2CC','EADCF8','D9EAD3','F4CCCC','CFE2F3','FCE5CD','D9D2E9',
    'D0E0E3','E6B8AF','B6D7A8','A4C2F4','FFD966','C9DAF8','D5A6BD','B4A7D6','CFE2F3','D9EAD3',
    'F9CB9C','B7E1CD','D5E8D4','F4CCCC','D0E0E3','EADCF8','FFF2CC','CFE2F3','E2F0D9','FCE4D6'
)



# ===========================================================================
#
# =============================== Logging ===================================
#
# ===========================================================================


$LogFile = Join-Path $LogRoot ("{0}_EventViewerRep_log_{1}.log" -f $Computer, $Timestamp)

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error','Success')]
        [string]$Level = 'Info'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        'Info'    { Write-Host $logEntry -ForegroundColor Cyan }
        'Warning' { Write-Host $logEntry -ForegroundColor Yellow }
        'Error'   { Write-Host $logEntry -ForegroundColor Red }
        'Success' { Write-Host $logEntry -ForegroundColor Green }
    }
    
    try {
        Add-Content -Path $LogFile -Value $logEntry -ErrorAction SilentlyContinue
    } catch {}
}

Write-Log "Starting diagnostic report collection on $Computer" -Level Info
$ParametersMessage = @"
  
  Computer Name = $Computer, 
  Start Time = $StartTime, 
  Time Stamp = $Timestamp, 
  ...no unique information for this report. 
"@
Write-Log "Parameters: $ParametersMessage " -Level Info




# ===========================================================================
# 
# =============================== Functions =================================
# 
# ===========================================================================

function Escape-XmlText {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Security.SecurityElement]::Escape($Text)
}

function Limit-Text {
    param([AllowNull()][string]$Text, [int]$Length = 500)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $Clean = ($Text -replace "`r`n", ' ' -replace "`n", ' ' -replace "`r", ' ')
    if ($Clean.Length -gt $Length) { return $Clean.Substring(0, $Length) }
    return $Clean
}

function Get-MessagePrefix {
    param([AllowNull()][string]$Message)
    $Clean = Limit-Text -Text $Message -Length 500
    if ([string]::IsNullOrEmpty($Clean)) { return '[No Message]' }
    if ($Clean.Length -gt 30) { return $Clean.Substring(0, 30) }
    return $Clean
}

function Get-ColumnName {
    param([int]$Index)
    $Name = ''
    while ($Index -gt 0) {
        $Modulo = ($Index - 1) % 26
        $Name = [char](65 + $Modulo) + $Name
        $Index = [math]::Floor(($Index - $Modulo) / 26)
    }
    return $Name
}

function New-StyleArray {
    param([int]$Count, [int]$Style)
    return @(for ($i = 0; $i -lt $Count; $i++) { $Style })
}

function New-BlankArray {
    param([int]$Count)
    return @(for ($i = 0; $i -lt $Count; $i++) { '' })
}

function New-RowXml {
    param(
        [int]$RowNumber,
        [object[]]$Values,
        [int[]]$Styles
    )

    $Cells = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $Column = Get-ColumnName ($i + 1)
        $Ref = "{0}{1}" -f $Column, $RowNumber
        $Value = Escape-XmlText ([string]$Values[$i])
        $Style = 0
        if ($Styles -and $i -lt $Styles.Count) { $Style = $Styles[$i] }

        if ($Style -gt 0) {
            $Cells.Add("<c r=`"$Ref`" s=`"$Style`" t=`"inlineStr`"><is><t>$Value</t></is></c>")
        }
        else {
            $Cells.Add("<c r=`"$Ref`" t=`"inlineStr`"><is><t>$Value</t></is></c>")
        }
    }

    return "<row r=`"$RowNumber`">$($Cells -join '')</row>"
}


#------------------------------ New-WorksheetXml --------------------------------

function New-WorksheetXml {
    param(
        [object[]]$Rows,
        [object[]]$StyleRows,
        [int]$ColumnCount,
        [int[]]$ColumnWidths
    )

    $RowXml = New-Object System.Collections.Generic.List[string]
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        $RowNumber = $r + 1
        $Values = [object[]]$Rows[$r]
        $Styles = [int[]]$StyleRows[$r]
        $RowXml.Add((New-RowXml -RowNumber $RowNumber -Values $Values -Styles $Styles))
    }

    $LastCol = Get-ColumnName $ColumnCount
    $LastRow = [math]::Max($Rows.Count, 1)
    
    $ColumnsBuilder = New-Object System.Text.StringBuilder
    $null = $ColumnsBuilder.Append("<cols>")
    for ($i = 0; $i -lt $ColumnWidths.Count; $i++) {
        $ColNum = $i + 1
        $null = $ColumnsBuilder.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f $ColNum, $ColumnWidths[$i]))
    }
    $null = $ColumnsBuilder.Append("</cols>")
    
@"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:$LastCol$LastRow"/>
<sheetViews>
<sheetView workbookViewId="0">
<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
</sheetView>
</sheetViews>
$($ColumnsBuilder.ToString())
<sheetData>
$($RowXml -join "`r`n")
</sheetData>
<autoFilter ref="A1:$LastCol$LastRow"/>
</worksheet>
"@
}

#------------------------------ Add-DataRows --------------------------------

function Add-DataRows {
    param(
        [object[]]$Headers,
        [object[]]$Items,
        [string]$GroupProperty,
        [int]$ColorColumnIndex,
        [hashtable]$ColorMap
    )

    $Rows = New-Object System.Collections.Generic.List[object]
    $Styles = New-Object System.Collections.Generic.List[object]

    $Rows.Add($Headers)
    $Styles.Add((New-StyleArray -Count $Headers.Count -Style 1))

    foreach ($Item in $Items) {
        $GroupValue = [string]$Item.$GroupProperty

        $Values = @()
        foreach ($Header in $Headers) {
            $Values += [string]$Item.$Header
        }

        $CellStyles = New-StyleArray -Count $Headers.Count -Style 0
        if ($ColorMap.ContainsKey($GroupValue)) {
            $CellStyles[$ColorColumnIndex] = $ColorMap[$GroupValue]
        }

        $Rows.Add($Values)
        $Styles.Add($CellStyles)
    }

    return [PSCustomObject]@{ Rows = $Rows.ToArray(); Styles = $Styles.ToArray() }
}



#------------------------------ Add-SortedDataRows --------------------------------


function Add-SortedDataRows {
    param(
        [object[]]$Headers,
        [object[]]$Items,
        [string]$GroupProperty,
        [int]$ColorColumnIndex,
        [hashtable]$ColorMap
    )

    $Rows = New-Object System.Collections.Generic.List[object]
    $Styles = New-Object System.Collections.Generic.List[object]

    $Rows.Add($Headers)
    $Styles.Add((New-StyleArray -Count $Headers.Count -Style 1))

    $LastGroup = $null
    foreach ($Item in $Items) {
        $GroupValue = [string]$Item.$GroupProperty

        if ($null -ne $LastGroup -and $GroupValue -ne $LastGroup) {
            $Rows.Add((New-BlankArray -Count $Headers.Count))
            $Styles.Add((New-StyleArray -Count $Headers.Count -Style 0))
        }

        $Values = @()
        foreach ($Header in $Headers) {
            $Values += [string]$Item.$Header
        }

        $CellStyles = New-StyleArray -Count $Headers.Count -Style 0
        if ($ColorMap.ContainsKey($GroupValue)) {
            $CellStyles[$ColorColumnIndex] = $ColorMap[$GroupValue]
        }

        $Rows.Add($Values)
        $Styles.Add($CellStyles)
        $LastGroup = $GroupValue
    }

    return [PSCustomObject]@{ Rows = $Rows.ToArray(); Styles = $Styles.ToArray() }
}



#--------------------------------- New-ColorMap --------------------------------


function New-ColorMap {
    param([string[]]$Keys)
    $Map = @{}
    $Index = 0
    foreach ($Key in ($Keys | Where-Object { $_ } | Select-Object -Unique)) {
        $Map[$Key] = 2 + ($Index % 30)
        $Index++
    }
    return $Map
}



#------------------------------ Write-TextFileUtf8 --------------------------------


function Write-TextFileUtf8 {
    param([string]$Path, [string]$Content)
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.UTF8Encoding]::new($false))
}


#--------------------------------- Get-EventRows --------------------------------



function Get-EventRows {
    param([string]$LogName)

    $Events = Get-WinEvent -FilterHashtable @{ LogName = $LogName; Level = 1,2; StartTime = $StartTime } -ErrorAction SilentlyContinue

    foreach ($Event in $Events) {
        $Message = Limit-Text -Text $Event.Message -Length 500
        [PSCustomObject]@{
            Date = $Event.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            Source = [string]$Event.ProviderName
            MessagePrefix = Get-MessagePrefix -Message $Message
            EventID = [string]$Event.Id
            Level = [string]$Event.LevelDisplayName
            LogName = [string]$Event.LogName
            Message = $Message
            SortDate = $Event.TimeCreated
        }
    }
}


#------------------------------ New-StylesXml ---------------------------------

function New-StylesXml {
    param([string[]]$Colors)

    $Fills = New-Object System.Collections.Generic.List[string]
    $Fills.Add('<fill><patternFill patternType="none"/></fill>')
    $Fills.Add('<fill><patternFill patternType="gray125"/></fill>')
    $Fills.Add('<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>')
    foreach ($Color in $Colors) {
        $Fills.Add("<fill><patternFill patternType=`"solid`"><fgColor rgb=`"FF$Color`"/><bgColor indexed=`"64`"/></patternFill></fill>")
    }

    $CellXfs = New-Object System.Collections.Generic.List[string]
    $CellXfs.Add('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')
    $CellXfs.Add('<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>')
    for ($i = 0; $i -lt $Colors.Count; $i++) {
        $FillId = 3 + $i
        $CellXfs.Add("<xf numFmtId=`"0`" fontId=`"0`" fillId=`"$FillId`" borderId=`"0`" xfId=`"0`" applyFill=`"1`"/>")
    }

@"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
</fonts>
<fills count="$($Fills.Count)">
$($Fills -join "`r`n")
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="$($CellXfs.Count)">
$($CellXfs -join "`r`n")
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"@
}



# ===========================================================================
# 
# =========================== Build this Thing ==============================
# 
# ===========================================================================

Write-Log "Beginning Data Collection Phase." -Level Info

New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null
New-Item -ItemType Directory -Path "$TempRoot\_rels" -Force | Out-Null
New-Item -ItemType Directory -Path "$TempRoot\xl\_rels" -Force | Out-Null
New-Item -ItemType Directory -Path "$TempRoot\xl\worksheets" -Force | Out-Null

$SystemEvents = @(Get-EventRows -LogName 'System')
$ApplicationEvents = @(Get-EventRows -LogName 'Application')

if (($SystemEvents.Count + $ApplicationEvents.Count) -eq 0) {
    throw 'There were no Events found in the last 48 hours'
}

$SystemTop = @(
    $SystemEvents |
    Group-Object Source |
    Sort-Object Count -Descending |
    Select-Object -First 5 |
    ForEach-Object {
        $Latest = $_.Group | Sort-Object SortDate -Descending | Select-Object -First 1
        [PSCustomObject]@{
            Source = $_.Name
            Count = [string]$_.Count
            RecentDate = $Latest.Date
            EventID = $Latest.EventID
            Level = $Latest.Level
            LogName = $Latest.LogName
            Message = $Latest.Message
        }
    }
)

$ApplicationTop = @(
    $ApplicationEvents |
    Group-Object MessagePrefix |
    Sort-Object Count -Descending |
    Select-Object -First 5 |
    ForEach-Object {
        $Latest = $_.Group | Sort-Object SortDate -Descending | Select-Object -First 1
        [PSCustomObject]@{
            Source = $Latest.Source
            MessagePrefix = $_.Name
            Count = [string]$_.Count
            RecentDate = $Latest.Date
            EventID = $Latest.EventID
            Level = $Latest.Level
            LogName = $Latest.LogName
            Message = $Latest.Message
        }
    }
)

$SystemRaw = @($SystemEvents | Sort-Object SortDate -Descending)
$SystemSorted = @($SystemEvents | Sort-Object Source, @{Expression='SortDate'; Descending=$true})
$ApplicationRaw = @($ApplicationEvents | Sort-Object SortDate -Descending)
$ApplicationSorted = @($ApplicationEvents | Sort-Object MessagePrefix, @{Expression='SortDate'; Descending=$true})

$SystemColorMap = New-ColorMap -Keys @($SystemEvents.Source)
$ApplicationColorMap = New-ColorMap -Keys @($ApplicationEvents.MessagePrefix)

$SysTopData = Add-DataRows -Headers @('Source','Count','RecentDate','EventID','Level','LogName','Message') -Items $SystemTop -GroupProperty 'Source' -ColorColumnIndex 0 -ColorMap $SystemColorMap
$SysRawData = Add-DataRows -Headers @('Date','Source','EventID','Level','LogName','Message') -Items $SystemRaw -GroupProperty 'Source' -ColorColumnIndex 1 -ColorMap $SystemColorMap
$SysSortedData = Add-SortedDataRows -Headers @('Date','Source','EventID','Level','LogName','Message') -Items $SystemSorted -GroupProperty 'Source' -ColorColumnIndex 1 -ColorMap $SystemColorMap

$AppTopData = Add-DataRows -Headers @('Source','Count','RecentDate','EventID','Level','LogName','Message') -Items $ApplicationTop -GroupProperty 'MessagePrefix' -ColorColumnIndex 0 -ColorMap $ApplicationColorMap
$AppRawData = Add-DataRows -Headers @('Date','Source','EventID','Level','LogName','Message') -Items $ApplicationRaw -GroupProperty 'MessagePrefix' -ColorColumnIndex 1 -ColorMap $ApplicationColorMap
$AppSortedData = Add-SortedDataRows -Headers @('Date','Source','EventID','Level','LogName','Message') -Items $ApplicationSorted -GroupProperty 'MessagePrefix' -ColorColumnIndex 1 -ColorMap $ApplicationColorMap

$Sheets = @(
    [PSCustomObject]@{ Name='Sys - Top Events'; Data=$SysTopData; Columns=7; Widths=$SysTopWidths }
    [PSCustomObject]@{ Name='Sys - Raw'; Data=$SysRawData; Columns=6; Widths=$SysRawWidths }
    [PSCustomObject]@{ Name='Sys - Sorted'; Data=$SysSortedData; Columns=6; Widths=$SysSortedWidths }
    [PSCustomObject]@{ Name='App - Top Events'; Data=$AppTopData; Columns=7; Widths=$AppTopWidths }
    [PSCustomObject]@{ Name='App - Raw'; Data=$AppRawData; Columns=6; Widths=$AppRawWidths }
    [PSCustomObject]@{ Name='App - Sorted'; Data=$AppSortedData; Columns=6; Widths=$AppSortedWidths }
)

for ($i = 0; $i -lt $Sheets.Count; $i++) {
    $SheetNumber = $i + 1
    $Xml = New-WorksheetXml -Rows $Sheets[$i].Data.Rows -StyleRows $Sheets[$i].Data.Styles -ColumnCount $Sheets[$i].Columns -ColumnWidths $Sheets[$i].Widths
    Write-TextFileUtf8 -Path "$TempRoot\xl\worksheets\sheet$SheetNumber.xml" -Content $Xml
}


$WorkbookSheets = New-Object System.Collections.Generic.List[string]
for ($i = 0; $i -lt $Sheets.Count; $i++) {
    $SheetNumber = $i + 1
    $Name = Escape-XmlText $Sheets[$i].Name
    $WorkbookSheets.Add("<sheet name=`"$Name`" sheetId=`"$SheetNumber`" r:id=`"rId$SheetNumber`"/>")
}

$WorkbookXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
$($WorkbookSheets -join "`r`n")
</sheets>
</workbook>
"@
Write-TextFileUtf8 -Path "$TempRoot\xl\workbook.xml" -Content $WorkbookXml
Write-Log "Checkpoint A." -Level Info

$WorkbookRels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/>
<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"@
Write-TextFileUtf8 -Path "$TempRoot\xl\_rels\workbook.xml.rels" -Content $WorkbookRels
Write-Log "Checkpoint B." -Level Info


$RootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
Write-TextFileUtf8 -Path "$TempRoot\_rels\.rels" -Content $RootRels
Write-Log "Checkpoint C." -Level Info

$ContentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
'@
Write-TextFileUtf8 -Path "$TempRoot\[Content_Types].xml" -Content $ContentTypes
Write-Log "Checkpoint D." -Level Info

$StylesXml = New-StylesXml -Colors $LightColors
Write-TextFileUtf8 -Path "$TempRoot\xl\styles.xml" -Content $StylesXml
Write-Log "Checkpoint E." -Level Info

if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($TempRoot, $OutputFile)
Remove-Item $TempRoot -Recurse -Force

Write-Host "Created report: $OutputFile"
Write-Log "Created report: $OutputFile"


# ===========================================================================
# +++++++++++EM+++++++++++++++                   +++++++++++EM+++++++++++++++    
# ============================= END OF PROGRAM ==============================
# +++++++++++EM+++++++++++++++                   +++++++++++EM+++++++++++++++                        
# ===========================================================================




