param(
  [string]$Source = "$env:USERPROFILE\Downloads\Fleet_Operations_Master_Final.xlsx",
  [string]$OutDir = (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Source)) { Write-Error "Excel file not found: $Source"; exit 1 }
$tmp = Join-Path $env:TEMP ("fleet_xlsx_" + [guid]::NewGuid().ToString('N').Substring(0,8))
if (Test-Path $tmp) { Remove-Item $tmp -Recurse -Force }
New-Item -ItemType Directory -Force $tmp | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($Source, $tmp)
$outDir = $OutDir
New-Item -ItemType Directory -Force $outDir | Out-Null
Write-Host "Importing '$Source' -> '$outDir\data.js'"

# ---- shared strings (DOM, UTF-8 safe via Load) ----
$ssx = New-Object System.Xml.XmlDocument
$ssx.Load("$tmp\xl\sharedStrings.xml")
$ns = New-Object System.Xml.XmlNamespaceManager($ssx.NameTable)
$ns.AddNamespace('a','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
$siNodes = $ssx.SelectNodes('//a:sst/a:si', $ns)
$SS = New-Object 'System.Collections.Generic.List[string]'
foreach ($si in $siNodes) { $SS.Add($si.InnerText) }
Write-Host "Shared strings:" $SS.Count

# ---- cellXfs -> numFmtId ----
$st = Get-Content "$tmp\xl\styles.xml" -Raw
$m = [regex]::Match($st, '<cellXfs count="\d+">(.*?)</cellXfs>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
$xfMatches = [regex]::Matches($m.Groups[1].Value, '<xf\b[^>]*?numFmtId="(\d+)"')
$XF = New-Object 'System.Collections.Generic.List[int]'
foreach ($x in $xfMatches) { $XF.Add([int]$x.Groups[1].Value) }

# date/time categories
$dateTimeIds = @{}
foreach($i in 22,165,175,176,185,187){ $dateTimeIds[$i]=$true }
$dateIds = @{}
foreach($i in 14,15,16,17,167,168,169,171,173,174,180,181,182,184,186,188,189,177){ $dateIds[$i]=$true }
$timeIds = @{}
foreach($i in 18,19,20,21,45,46,47,164,166,170,172,178,179){ $timeIds[$i]=$true }
$base = [datetime]'1899-12-30'

function Get-Cat([int]$styleIdx){
  if ($styleIdx -lt 0 -or $styleIdx -ge $XF.Count){ return 'n' }
  $f = $XF[$styleIdx]
  if ($dateTimeIds.ContainsKey($f)){ return 'dt' }
  if ($dateIds.ContainsKey($f)){ return 'd' }
  if ($timeIds.ContainsKey($f)){ return 't' }
  return 'n'
}

function ColToIndex([string]$ref){
  $col = 0
  foreach($ch in $ref.ToCharArray()){
    if ($ch -ge 'A' -and $ch -le 'Z'){ $col = $col*26 + ([int][char]$ch - 64) }
    elseif ($ch -ge '0' -and $ch -le '9'){ break }
    else { break }
  }
  return $col - 1
}

# ---- read sheet order + names from workbook.xml, map r:id -> worksheet file via rels ----
$wb = New-Object System.Xml.XmlDocument; $wb.Load("$tmp\xl\workbook.xml")
$wbns = New-Object System.Xml.XmlNamespaceManager($wb.NameTable)
$wbns.AddNamespace('a','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
$wbns.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
$rels = New-Object System.Xml.XmlDocument; $rels.Load("$tmp\xl\_rels\workbook.xml.rels")
$relMap = @{}
foreach($rel in $rels.DocumentElement.ChildNodes){ $relMap[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target') }
$sheetDefs = New-Object 'System.Collections.Generic.List[object]'
foreach($sh in $wb.SelectNodes('//a:sheets/a:sheet', $wbns)){
  $rid = $sh.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $target = $relMap[$rid]
  $file = Join-Path $tmp ("xl\" + ($target -replace '/','\'))
  $sheetDefs.Add([pscustomobject]@{ name = $sh.GetAttribute('name'); file = $file })
}
Write-Host "Sheets in workbook:" $sheetDefs.Count

$allSheets = New-Object 'System.Collections.Generic.List[object]'

for ($sn=1; $sn -le $sheetDefs.Count; $sn++){
  $path = $sheetDefs[$sn-1].file
  $settings = New-Object System.Xml.XmlReaderSettings
  $settings.IgnoreWhitespace = $true
  $reader = [System.Xml.XmlReader]::Create($path, $settings)
  $rows = New-Object 'System.Collections.Generic.List[object]'
  $curRow = $null
  $maxCol = 0
  $advance = $true
  while ($true){
    if ($advance){ if (-not $reader.Read()){ break } } else { $advance = $true }
    if ($reader.NodeType -eq 'Element' -and $reader.LocalName -eq 'row'){
      if ($curRow -ne $null){ $rows.Add($curRow) }
      $curRow = New-Object 'System.Collections.Generic.Dictionary[int,string]'
    }
    elseif ($reader.NodeType -eq 'Element' -and $reader.LocalName -eq 'c'){
      $ref = $reader.GetAttribute('r')
      $t = $reader.GetAttribute('t')
      $sAttr = $reader.GetAttribute('s')
      $colIdx = if($ref){ ColToIndex $ref } else { 0 }
      $val = $null
      if (-not $reader.IsEmptyElement){
        $inner = $reader.ReadInnerXml()   # reader now positioned on node AFTER </c>
        $advance = $false
        $vText = $null
        $isText = $null
        $mv = [regex]::Match($inner, '<v[^>]*>(.*?)</v>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($mv.Success){ $vText = [System.Net.WebUtility]::HtmlDecode($mv.Groups[1].Value) }
        $mt = [regex]::Matches($inner, '<t[^>]*>(.*?)</t>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($mt.Count -gt 0){ $sbT = New-Object System.Text.StringBuilder; foreach($tm in $mt){ [void]$sbT.Append([System.Net.WebUtility]::HtmlDecode($tm.Groups[1].Value)) }; $isText = $sbT.ToString() }
        if ($t -eq 's'){ if($vText -match '^\d+$'){ $i=[int]$vText; if($i -ge 0 -and $i -lt $SS.Count){ $val = $SS[$i] } } }
        elseif ($t -eq 'inlineStr'){ $val = $isText }
        elseif ($t -eq 'str'){ $val = $vText }
        elseif ($t -eq 'b'){ $val = if($vText -eq '1'){'TRUE'}else{'FALSE'} }
        elseif ($t -eq 'e'){ $val = $vText }
        else {
          # numeric
          if ($vText -ne $null -and $vText -ne ''){
            $cat = if($sAttr){ Get-Cat ([int]$sAttr) } else { 'n' }
            if ($cat -ne 'n'){
              try {
                $d = $base.AddDays([double]$vText)
                switch($cat){
                  'dt' { $val = $d.ToString('dd/MM/yyyy HH:mm') }
                  'd'  { $val = $d.ToString('dd/MM/yyyy') }
                  't'  { $val = $d.ToString('HH:mm') }
                }
              } catch { $val = $vText }
            } else { $val = $vText }
          }
        }
      }
      if ($val -ne $null){ $val = $val.Trim() }
      if ($val -ne $null -and $val -ne ''){
        $curRow[$colIdx] = $val
        if ($colIdx+1 -gt $maxCol){ $maxCol = $colIdx+1 }
      }
    }
  }
  if ($curRow -ne $null){ $rows.Add($curRow) }
  $reader.Close()

  # Build grid as list of string arrays, dropping fully-empty rows
  $grid = New-Object 'System.Collections.Generic.List[object]'
  foreach($r in $rows){
    if ($r.Count -eq 0){ continue }
    $arr = New-Object 'string[]' $maxCol
    for($c=0;$c -lt $maxCol;$c++){ $arr[$c] = if($r.ContainsKey($c)){ $r[$c] } else { '' } }
    $grid.Add($arr)
  }

  # Convert serial-date columns detected by header name (e.g. Timestamp/Date/Time)
  if ($grid.Count -gt 1){
    $hdr = $grid[0]
    for($c=0;$c -lt $maxCol;$c++){
      $h = ("" + $hdr[$c]).ToLower()
      if ($h -match 'timestamp|time stamp|date|time'){
        # check majority of column values are excel serials
        $serialCnt=0; $checked=0
        for($ri=1; $ri -lt [Math]::Min($grid.Count, 40); $ri++){
          $v = $grid[$ri][$c]; if($v -ne ''){ $checked++; if($v -match '^\d{4,5}(\.\d+)?$'){ $serialCnt++ } }
        }
        if ($checked -gt 0 -and $serialCnt -ge [Math]::Ceiling($checked*0.7)){
          for($ri=1; $ri -lt $grid.Count; $ri++){
            $v = $grid[$ri][$c]
            if ($v -match '^\d{4,5}(\.\d+)?$'){
              try { $d=$base.AddDays([double]$v); $grid[$ri][$c] = $d.ToString('dd/MM/yyyy HH:mm') } catch {}
            }
          }
        }
      }
    }
  }
  $allSheets.Add([ordered]@{ name=$sheetDefs[$sn-1].name; rows=$grid.Count; cols=$maxCol; grid=$grid })
  Write-Host ("sheet{0,-2} {1,-22} rows={2,-5} cols={3}" -f $sn, $sheetDefs[$sn-1].name, $grid.Count, $maxCol)
}

$payload = [ordered]@{ generated = (Get-Date).ToString('yyyy-MM-dd HH:mm'); sheets = $allSheets }
$json = $payload | ConvertTo-Json -Depth 6 -Compress
$out = "window.FLEET_DATA = " + $json + ";"
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText("$outDir\data.js", $out, $utf8)
Write-Host "`nWrote data.js size MB:" ([math]::Round((Get-Item "$outDir\data.js").Length/1MB,2))
try { Remove-Item $tmp -Recurse -Force } catch {}
Write-Host "Done. Open index.html to view."
