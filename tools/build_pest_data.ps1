# Builds pest_data.json from the Biofuturo MIP visit folders:
#   - RESUMEN PROSPECCION *.xlsx  -> numeric readings per pest x sector
#   - *.kml / *.kmz               -> georeferenced points, joined to the sector polygons
#   - Manifold *.pdf              -> merged in from observations.json (curated narrative)
# Requires Excel (COM) to read the workbooks. Run from PowerShell 5.1.
param(
  # Carpeta con una subcarpeta por visita (dd-MM-yyyy). Por defecto la sincronizada
  # por OneDrive; se puede pasar cualquier ruta con -Src.
  [string]$Src      = (Join-Path $env:USERPROFILE 'OneDrive - Sembrador\Biofuturo San Gerardo'),
  # index.html of the dashboard: source of truth for the 23 sector polygons
  [string]$IndexHtml = (Join-Path $PSScriptRoot '..\index.html'),
  # curated narrative extracted from the Manifold PDFs
  [string]$Observations = (Join-Path $PSScriptRoot 'observations.json'),
  # scratch folder for KMZ extraction
  [string]$Work     = $env:TEMP,
  [string]$OutFile  = (Join-Path $PSScriptRoot '..\pest_data.json')
)

$ErrorActionPreference = 'Stop'
$inv = [Globalization.CultureInfo]::InvariantCulture
$issues = [System.Collections.Generic.List[object]]::new()
function AddIssue($visit, $severity, $source, $field, $detail) {
  $issues.Add([ordered]@{ visit=$visit; severity=$severity; source=$source; field=$field; detail=$detail })
}

# ---------------------------------------------------------------- helpers
function Slug([string]$t) {
  if (-not $t) { return $null }
  $t = $t.Trim().ToLower()
  $n = $t.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object Text.StringBuilder
  foreach ($ch in $n.ToCharArray()) {
    if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($ch) }
  }
  $t = $sb.ToString() -replace '[^a-z0-9]+','-'
  return $t.Trim('-')
}
function Norm([string]$t) { if ($null -eq $t) { return $null }; ($t -replace '\s+',' ').Trim() }

$MONTHS = @('ene.','feb.','mar.','abr.','may.','jun.','jul.','ago.','sep.','oct.','nov.','dic.')
function DateLabel([datetime]$d) { "$($d.Day) $($MONTHS[$d.Month-1]) $($d.Year)" }

# ---------------------------------------------------------------- sectors: read the GEOJSON literal out of index.html
$html = Get-Content -Raw $IndexHtml -Encoding UTF8
$gi = $html.IndexOf('const GEOJSON = ')
if ($gi -lt 0) { throw "No se encontro 'const GEOJSON =' en $IndexHtml" }
$gs = $html.IndexOf('{', $gi); $depth = 0; $ge = -1
for ($p = $gs; $p -lt $html.Length; $p++) {
  if ($html[$p] -eq '{') { $depth++ } elseif ($html[$p] -eq '}') { $depth--; if ($depth -eq 0) { $ge = $p; break } }
}
$geo = $html.Substring($gs, $ge - $gs + 1) | ConvertFrom-Json
$sectors = @()
foreach ($f in $geo.features) {
  $p = $f.properties
  $sectors += [ordered]@{
    key       = "E$($p.equipo_num)-S$($p.sector_num)"
    code      = "$($p.equipo_num).$($p.sector_num)"
    equipo    = [int]$p.equipo_num
    sector    = [int]$p.sector_num
    name      = "Equipo $($p.equipo_num) - S$($p.sector_num)"
    hectareas = [double]$p.hectareas
    ring      = $f.geometry.coordinates[0]
  }
}
$byCode = @{}; foreach ($s in $sectors) { $byCode[$s.code] = $s }

function PointInRing($lon, $lat, $ring) {
  $inside = $false; $n = $ring.Count
  for ($i=0; $i -lt $n; $i++) {
    $j = ($i + $n - 1) % $n
    $xi = [double]$ring[$i][0]; $yi = [double]$ring[$i][1]
    $xj = [double]$ring[$j][0]; $yj = [double]$ring[$j][1]
    if ((($yi -gt $lat) -ne ($yj -gt $lat)) -and
        ($lon -lt ($xj - $xi) * ($lat - $yi) / ($yj - $yi) + $xi)) { $inside = -not $inside }
  }
  return $inside
}
function NearestVertexM($lon, $lat, $ring) {
  $best = [double]::MaxValue
  $mLat = 111132.0
  $mLon = 111320.0 * [Math]::Cos($lat * [Math]::PI / 180)
  foreach ($c in $ring) {
    $dx = ([double]$c[0] - $lon) * $mLon
    $dy = ([double]$c[1] - $lat) * $mLat
    $d = [Math]::Sqrt($dx*$dx + $dy*$dy)
    if ($d -lt $best) { $best = $d }
  }
  return [Math]::Round($best, 1)
}

$allLon = @(); $allLat = @()
foreach ($s in $sectors) { foreach ($c in $s.ring) { $allLon += [double]$c[0]; $allLat += [double]$c[1] } }
$bbox = @{ minLon = ($allLon | Measure-Object -Minimum).Minimum - 0.01
           maxLon = ($allLon | Measure-Object -Maximum).Maximum + 0.01
           minLat = ($allLat | Measure-Object -Minimum).Minimum - 0.01
           maxLat = ($allLat | Measure-Object -Maximum).Maximum + 0.01 }

# ---------------------------------------------------------------- KML label -> canonical pest
$LABEL_MAP = @(
  @{ match=@('acaros fitofagos','acaros','acaro','acsors','acaros fitofagos moviles'); pest='acaros-fitofagos-moviles'; species=$null },
  @{ match=@('aranita roja europea','foco panonychus ulmi','panonychus ulmi');          pest='acaros-fitofagos-moviles'; species='Panonychus ulmi' },
  @{ match=@('aranita bimaculada','tetranychus urticae');                               pest='acaros-fitofagos-moviles'; species='Tetranychus urticae' },
  @{ match=@('pulgon del nogal','pulgon nogal','pulgon','pulgones','ninfas de pulgon','ninfas de pulgon del nogal'); pest='pulgon-del-nogal'; species='Chromaphis juglandicola' },
  @{ match=@('pulgon del avellano');                                                    pest='pulgon-del-avellano'; species='Myzocallis coryli' },
  @{ match=@('escamas','escama morada','escama');                                       pest='escamas'; species='Lepidosaphes ulmi' },
  @{ match=@('escama de san jose');                                                     pest='escamas'; species='Diaspidiotus perniciosus' },
  @{ match=@('conchuela grande cafe','conchuela blanda cafe','conchuelas');             pest='conchuelas'; species='Parthenolecanium corni' },
  @{ match=@('capachitos','capachito de los frutales');                                 pest='capachitos'; species='Geniocremnus chilensis' },
  @{ match=@('chanchito blanco');                                                       pest='chanchito-blanco'; species='Familia Pseudococcidae' },
  @{ match=@('trips');                                                                  pest='trips'; species='Frankliniella occidentalis' }
)
# scientific names the Excel catalog leaves empty but the Manifold reports state
$SCI_FALLBACK = @{
  'pulgon-del-nogal'  = 'Chromaphis juglandicola'
  'capachitos'        = 'Geniocremnus chilensis (Familia Curculionidae)'
  'larvas-xilofagas'  = $null
}

function MapLabel([string]$raw) {
  $k = Slug $raw
  if (-not $k) { return $null }
  $k = $k -replace '-',' '
  foreach ($e in $LABEL_MAP) { if ($e.match -contains $k) { return $e } }
  return $null
}

# Display taxonomy for the point layer. The points use their own key space: the field app
# records the pest observed at the point, which is coarser than the spreadsheet rows
# (e.g. one 'acaros-fitofagos-moviles' point vs. the '…-perimetro' / '…-interior' rows).
$PEST_TAXA = [ordered]@{
  'acaros-fitofagos-moviles' = @{ name='Ácaros fitófagos móviles'; scientific='Familia Tetranychidae' }
  'pulgon-del-nogal'         = @{ name='Pulgón del nogal';          scientific='Chromaphis juglandicola' }
  'pulgon-del-avellano'      = @{ name='Pulgón del avellano';       scientific='Myzocallis coryli' }
  'escamas'                  = @{ name='Escamas';                   scientific='Familia Diaspididae' }
  'conchuelas'               = @{ name='Conchuelas';                scientific='Familia Coccidae' }
  'capachitos'               = @{ name='Capachitos';                scientific='Familia Curculionidae' }
  'trips'                    = @{ name='Trips';                     scientific='Frankliniella occidentalis' }
  'chanchito-blanco'         = @{ name='Chanchito blanco';          scientific='Familia Pseudococcidae' }
}

# Cada fila de la planilla se adscribe a un taxón de $PEST_TAXA. Es lo que permite
# que el mapa, la leyenda y las tarjetas del informe compartan el mismo color:
# 'acaros-fitofagos-moviles-perimetro' y '-interior' son la misma plaga, y
# 'viabilidad-escamas' es una métrica de 'escamas', no otra plaga.
$TAXON_OF = @{
  'acaros-fitofagos-moviles-perimetro'             = 'acaros-fitofagos-moviles'
  'acaros-fitofagos-moviles-interior'              = 'acaros-fitofagos-moviles'
  'acaros-fitofagos-moviles'                       = 'acaros-fitofagos-moviles'
  'estructuras-con-huevos-de-acaros-fitofagos'     = 'acaros-fitofagos-moviles'
  'rangos-estimados-de-huevos-de-acaros-fitofagos' = 'acaros-fitofagos-moviles'
  'viabilidad-huevos-de-acaros-fitofagos'          = 'acaros-fitofagos-moviles'
  'huevos-a-nivel-foliar'                          = 'acaros-fitofagos-moviles'
  'escamas'                                        = 'escamas'
  'viabilidad-escamas'                             = 'escamas'
  'conchuelas'                                     = 'conchuelas'
  'viabilidad-conchuelas'                          = 'conchuelas'
  'pulgon-del-nogal'                               = 'pulgon-del-nogal'
  'pulgon-del-avellano'                            = 'pulgon-del-avellano'
  'afidos'                                         = 'pulgon-del-nogal'
  'trips'                                          = 'trips'
  'capachitos'                                     = 'capachitos'
  'chanchito-blanco'                               = 'chanchito-blanco'
  'control-natural-chanchito-blanco'               = 'chanchito-blanco'
}

# ---------------------------------------------------------------- workbooks
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false

$visits   = @()
$readings = [System.Collections.Generic.List[object]]::new()
$pestDefs = [ordered]@{}
$catalog  = [ordered]@{}

foreach ($f in (Get-ChildItem -Path $Src -Recurse -Filter '*.xlsx' | Sort-Object { $_.Directory.Name })) {
  $folder = $f.Directory.Name
  $wb = $xl.Workbooks.Open($f.FullName, 0, $true)

  if ($catalog.Count -eq 0) {
    $ws = $wb.Worksheets.Item('Hoja1'); $ur = $ws.UsedRange
    for ($r=2; $r -le $ur.Rows.Count; $r++) {
      $name = Norm $ur.Cells.Item($r,1).Text
      if (-not $name) { continue }
      $key = Slug $name
      if ($catalog.Contains($key)) { continue }
      $catalog[$key] = [ordered]@{
        key           = $key
        name          = $name
        unit          = Norm $ur.Cells.Item($r,2).Text
        scientific    = Norm $ur.Cells.Item($r,3).Text
        isParasitism  = (Norm $ur.Cells.Item($r,4).Text) -eq 'SI'
        isViability   = (Norm $ur.Cells.Item($r,5).Text) -eq 'SI'
        isNaturalCtrl = (Norm $ur.Cells.Item($r,6).Text) -eq 'SI'
      }
    }
  }

  $ws = $wb.Worksheets.Item('Biofuturo Ltda.'); $ur = $ws.UsedRange
  $rows = $ur.Rows.Count; $cols = $ur.Columns.Count

  $visitDate = $null
  for ($r=1; $r -le 6 -and -not $visitDate; $r++) {
    for ($c=1; $c -le $cols; $c++) {
      if ((Norm $ur.Cells.Item($r,$c).Text) -match '^Fecha') {
        for ($k=$c+1; $k -le $cols; $k++) {
          $v = $ur.Cells.Item($r,$k).Value2
          if ($v -is [double]) { $visitDate = [datetime]::FromOADate($v); break }
          $t = Norm $ur.Cells.Item($r,$k).Text
          if ($t -match '^\d{2}/\d{2}/\d{4}$') { $visitDate = [datetime]::ParseExact($t,'dd/MM/yyyy',$inv); break }
        }
        break
      }
    }
  }
  if (-not $visitDate) {
    $parts = $folder -split '-'
    $visitDate = Get-Date -Year ([int]$parts[2]) -Month ([int]$parts[1]) -Day ([int]$parts[0]) -Hour 0 -Minute 0 -Second 0
    AddIssue $folder 'warn' 'xlsx' 'Fecha' 'La celda "Fecha :" no traía un valor legible; se usó la fecha del nombre de la carpeta.'
  }
  $vkey = $visitDate.ToString('yyyy-MM-dd')

  $hdr = 0
  for ($r=1; $r -le $rows; $r++) { if ((Norm $ur.Cells.Item($r,1).Text) -eq 'Frutal') { $hdr = $r; break } }
  if ($hdr -eq 0) { throw "No se encontro la fila de encabezado en $folder" }

  $colToSector = @{}
  for ($c=1; $c -le $cols; $c++) {
    $t = Norm $ur.Cells.Item($hdr,$c).Text
    if ($t -match '^([1-5])\.([1-5])$') { $colToSector[$c] = "$($matches[1]).$($matches[2])" }
  }
  if ($colToSector.Count -ne 23) { AddIssue $vkey 'warn' 'xlsx' 'sectores' "Se detectaron $($colToSector.Count) columnas de sector (se esperaban 23)." }

  $crop = $null; $seen = @{}; $pestRows = 0
  for ($r=$hdr+1; $r -le $rows; $r++) {
    $pestName = Norm $ur.Cells.Item($r,3).Text
    if (-not $pestName) { continue }
    if (-not $crop) { $crop = Norm $ur.Cells.Item($r,1).Text }
    $pestRows++

    $key = Slug $pestName
    if ($seen.ContainsKey($key)) {
      $seen[$key]++
      AddIssue $vkey 'error' 'xlsx' $pestName "La plaga aparece repetida en la misma planilla con valores distintos; la segunda fila se conservó como '$key#$($seen[$key])'."
      $key = "$key#$($seen[$key])"
    } else { $seen[$key] = 1 }

    $unit = Norm $ur.Cells.Item($r,4).Text
    if (-not $pestDefs.Contains($key)) {
      $base = ($key -split '#')[0]
      $cat  = if ($catalog.Contains($base)) { $catalog[$base] } else { $null }
      $sci  = Norm $ur.Cells.Item($r,5).Text
      if (-not $sci -and $cat) { $sci = $cat.scientific }
      if (-not $sci -and $SCI_FALLBACK.ContainsKey($base)) { $sci = $SCI_FALLBACK[$base] }
      $taxon = if ($TAXON_OF.ContainsKey($base)) { $TAXON_OF[$base] } else { $null }
      if (-not $taxon) { AddIssue $vkey 'warn' 'xlsx' $pestName "La fila no está adscrita a ningún taxón en `$TAXON_OF; en el mapa y el informe saldrá con el color neutro." }
      $pestDefs[$key] = [ordered]@{
        key           = $key
        taxon         = $taxon
        name          = $pestName
        unit          = $unit
        metric        = $(if ($unit -eq 'Porcentaje') { 'porcentaje' } elseif ($unit -eq 'Promedios') { 'promedio' } else { Slug $unit })
        unitLabel     = $(if ($unit -eq 'Porcentaje') { '% de plantas o estructuras con presencia' } else { 'individuos promedio por hoja' })
        scientific    = $sci
        isParasitism  = (Norm $ur.Cells.Item($r,6).Text) -eq 'SI'
        isViability   = (Norm $ur.Cells.Item($r,7).Text) -eq 'SI'
        isNaturalCtrl = (Norm $ur.Cells.Item($r,8).Text) -eq 'SI'
        inCatalog     = [bool]$cat
      }
    }

    foreach ($c in $colToSector.Keys) {
      $code = $colToSector[$c]
      $raw  = Norm $ur.Cells.Item($r,$c).Text
      $value = $null; $rangeText = $null
      if ($raw) {
        $cand = $raw -replace ',','.'
        $ok = 0.0
        if ([double]::TryParse($cand, [Globalization.NumberStyles]::Float, $inv, [ref]$ok)) {
          $value = [Math]::Round($ok, 4)
        } elseif ($raw -match '^\s*\d+\s*a\s*\d+\s*$') {
          $rangeText = $raw
        } else {
          AddIssue $vkey 'error' 'xlsx' "$pestName / cuartel $code" "Valor no numérico '$raw'; se registró como nulo."
        }
      }
      $sec = $byCode[$code]
      $readings.Add([ordered]@{
        visit = $vkey; pest = $key; sector = $sec.key
        equipo = $sec.equipo; sectorNum = $sec.sector
        value = $value; rangeText = $rangeText; raw = $raw
      })
    }
  }

  $visits += [ordered]@{ folder = $folder; date = $vkey; pestRows = $pestRows; crop = $crop; file = $f.Name }
  $wb.Close($false)
}
$xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)

# ---------------------------------------------------------------- georeferenced points
$points = [System.Collections.Generic.List[object]]::new()
$kmlFiles = @()
Add-Type -AssemblyName System.IO.Compression.FileSystem
foreach ($d in (Get-ChildItem -Path $Src -Directory)) {
  foreach ($k in (Get-ChildItem -Path $d.FullName -Filter '*.kml' -ErrorAction SilentlyContinue)) {
    $kmlFiles += @{ folder=$d.Name; path=$k.FullName; name=$k.Name }
  }
  foreach ($k in (Get-ChildItem -Path $d.FullName -Filter '*.kmz' -ErrorAction SilentlyContinue)) {
    $tmp = Join-Path $Work ("kmz_" + $d.Name)
    if (Test-Path $tmp) { Remove-Item -Recurse -Force $tmp }
    New-Item -ItemType Directory -Force $tmp | Out-Null
    $zip = Join-Path $tmp 'a.zip'; Copy-Item $k.FullName $zip
    [IO.Compression.ZipFile]::ExtractToDirectory($zip, $tmp)
    foreach ($inner in (Get-ChildItem -Path $tmp -Filter '*.kml' -Recurse)) {
      $kmlFiles += @{ folder=$d.Name; path=$inner.FullName; name=$k.Name }
    }
  }
}

$folderToVisit = @{}; foreach ($v in $visits) { $folderToVisit[$v.folder] = $v.date }
$foldersWithKml = @{}

foreach ($kf in $kmlFiles) {
  $vkey = $folderToVisit[$kf.folder]
  $foldersWithKml[$kf.folder] = $true
  $txt = Get-Content -Raw $kf.path -Encoding UTF8
  if ($txt -notmatch '</kml>' -or $txt -match '<na\}') {
    AddIssue $vkey 'error' 'kml' $kf.name 'El archivo KML está truncado o corrupto: el primer bloque de Placemarks no es XML válido y esos puntos no se pudieron recuperar.'
  }
  foreach ($m in [regex]::Matches($txt, '(?s)<Placemark>(.*?)</Placemark>')) {
    $b = $m.Groups[1].Value
    $id    = if ($b -match '<name>(.*?)</name>') { Norm $matches[1] } else { $null }
    $label = if ($b -match '<description>(.*?)</description>') { Norm $matches[1] } else { $null }
    if ($label -eq 'no description') { $label = $null }
    if (-not ($b -match '<coordinates>(.*?)</coordinates>')) { continue }
    $co = ($matches[1].Trim() -split ',')
    $lon = [double]::Parse($co[0], $inv); $lat = [double]::Parse($co[1], $inv)

    $inBounds = ($lon -ge $bbox.minLon -and $lon -le $bbox.maxLon -and $lat -ge $bbox.minLat -and $lat -le $bbox.maxLat)
    if (-not $inBounds) {
      # Coordinates from another property: dropped from points[] so the map layer never
      # has to filter them out at runtime. The reason stays recorded here.
      AddIssue $vkey 'error' 'kml' $id "El punto ($lat, $lon) queda fuera del predio (a decenas de km); no corresponde a San Gerardo. Se excluyó de points[]."
      continue
    }

    $secKey = $null
    foreach ($s in $sectors) { if (PointInRing $lon $lat $s.ring) { $secKey = $s.key; break } }
    if (-not $secKey) {
      $best = [double]::MaxValue; $nearest = $null
      foreach ($s in $sectors) { $dd = NearestVertexM $lon $lat $s.ring; if ($dd -lt $best) { $best = $dd; $nearest = $s.key } }
      AddIssue $vkey 'info' 'kml' $id "El punto no cae dentro de ningún polígono de cuartel; el más cercano es $nearest a $best m (probable hilera perimetral). Queda con sector nulo."
    }

    $map = MapLabel $label
    if ($label -and -not $map) { AddIssue $vkey 'warn' 'kml' $id "Etiqueta de plaga no reconocida: '$label'." }
    if (-not $label)           { AddIssue $vkey 'warn' 'kml' $id 'Punto sin descripción de plaga.' }

    $points.Add([ordered]@{
      visit      = $vkey
      id         = $id
      lat        = [Math]::Round($lat, 7)
      lon        = [Math]::Round($lon, 7)
      rawLabel   = $label
      pest       = $(if ($map) { $map.pest } else { $null })
      species    = $(if ($map) { $map.species } else { $null })
      stage      = $(if ($label -and (Slug $label) -like 'ninfas*') { 'ninfas' } else { $null })
      sector     = $secKey
      equipo     = $(if ($secKey) { [int]($secKey -replace 'E(\d)-S\d','$1') } else { $null })
      sourceFile = $kf.name
    })
  }
}
foreach ($v in $visits) {
  if (-not $foldersWithKml.ContainsKey($v.folder)) {
    AddIssue $v.date 'warn' 'kml' '-' 'La visita no tiene archivo de puntos georreferenciados (KML/KMZ) en su carpeta.'
  }
}

# ---------------------------------------------------------------- narrative reports
$reports = @()
if (Test-Path $Observations) {
  # ConvertFrom-Json in WinPS 5.1 emits a JSON array as a single object, not enumerated
  $reports = @(Get-Content -Raw $Observations -Encoding UTF8 | ConvertFrom-Json)
  if ($reports.Count -eq 1 -and $reports[0] -is [Array]) { $reports = @($reports[0]) }
} else {
  AddIssue '-' 'warn' 'manifold' 'observations.json' "No se encontró $Observations; el JSON queda sin la parte cualitativa."
}

# issues declared alongside the curated narrative (contradictions found when cross-checking
# the Manifold text against the numeric summary)
foreach ($r in $reports) {
  if ($r.PSObject.Properties.Name -contains 'issues' -and $r.issues) {
    foreach ($i in $r.issues) { AddIssue $r.visit $i.severity $i.source $i.field $i.detail }
  }
}

# ---------------------------------------------------------------- assemble
$repByVisit = @{}; foreach ($r in $reports) { $repByVisit[$r.visit] = $r }
$ptsByVisit = @{}
foreach ($p in $points) { if (-not $ptsByVisit.ContainsKey($p.visit)) { $ptsByVisit[$p.visit] = 0 }; $ptsByVisit[$p.visit]++ }

$visitOut = @()
# Sort-Object -Property no funciona sobre OrderedDictionary: hay que usar un script block,
# si no visits[] queda en orden de carpeta y deja de ser una serie temporal.
foreach ($v in ($visits | Sort-Object { $_.date })) {
  $d = [datetime]::ParseExact($v.date, 'yyyy-MM-dd', $inv)
  $rep = $repByVisit[$v.date]
  $visitOut += [ordered]@{
    key          = $v.date
    label        = DateLabel $d
    date         = $v.date
    year         = $d.Year
    month        = $d.Month
    folder       = $v.folder
    crop         = $v.crop
    phenology    = $(if ($rep) { $rep.phenology } else { $null })
    harvestStart = $(if ($rep) { $rep.harvestStart } else { $null })
    harvestEnd   = $(if ($rep) { $rep.harvestEnd } else { $null })
    author       = $(if ($rep) { $rep.author } else { $null })
    pestRows     = $v.pestRows
    pointCount   = $(if ($ptsByVisit.ContainsKey($v.date)) { $ptsByVisit[$v.date] } else { 0 })
    hasPoints    = $ptsByVisit.ContainsKey($v.date)
    hasReport    = [bool]$rep
    sources      = [ordered]@{ summary = $v.file; manifold = $(if ($rep) { $rep.source } else { $null }) }
  }
}

# catalog of the pest taxa that actually appear in points[], for the map layer's filter+legend
$pointPestOut = @()
foreach ($k in $PEST_TAXA.Keys) {
  $n = @($points | Where-Object { $_.pest -eq $k }).Count
  if ($n -eq 0) { continue }
  $pointPestOut += [ordered]@{
    key = $k; name = $PEST_TAXA[$k].name; scientific = $PEST_TAXA[$k].scientific; pointCount = $n
  }
}
$unmapped = @($points | Where-Object { -not $_.pest }).Count
if ($unmapped -gt 0) {
  $pointPestOut += [ordered]@{
    key = '_default'; name = 'Sin clasificar'; scientific = $null; pointCount = $unmapped
  }
}

$sectorOut = @()
foreach ($s in ($sectors | Sort-Object { $_.equipo }, { $_.sector })) {
  $sectorOut += [ordered]@{
    key=$s.key; code=$s.code; equipo=$s.equipo; sector=$s.sector
    name=$s.name; mapUnit=$s.name; hectareas=$s.hectareas
  }
}

# readings y points tambien se emiten en orden cronologico
$readingsOut = @($readings | Sort-Object { $_.visit }, { $_.pest }, { $_.equipo }, { $_.sectorNum })
$pointsOut   = @($points   | Sort-Object { $_.visit }, { $_.id })

$doc = [ordered]@{
  meta = [ordered]@{
    farm      = 'Fundo San Gerardo'
    company   = 'Agrícola San Gerardo SpA.'
    orchard   = 'Huerto de Nogales'
    crop      = 'Nogales'
    varieties = @('Chandler','Cisco','Franquette')
    surfaceHa = 123
    provider  = 'Biofuturo Ltda.'
    program   = 'Monitoreo MIP / prospección fitosanitaria'
    season    = '2025-26'
    protocol  = [ordered]@{
      plantsPerSector = 30
      leavesPerPlant  = 5
      metrics = @(
        ([ordered]@{ unit='Promedios';  description='Promedio de hojas con presencia por cuartel (individuos promedio por hoja)' }),
        ([ordered]@{ unit='Porcentaje'; description='Porcentaje de plantas o estructuras con presencia por cuartel' })
      )
    }
    generatedFrom = 'RESUMEN PROSPECCIÓN *.xlsx + KML/KMZ de puntos + Manifold Post Visita *.pdf'
  }
  visits     = $visitOut
  sectors    = $sectorOut
  pests      = @($pestDefs.Values)
  pointPests = @($pointPestOut)
  catalog  = @($catalog.Values)
  readings = $readingsOut
  points   = $pointsOut
  reports  = $reports
  issues   = @($issues)
}

$json = $doc | ConvertTo-Json -Depth 12 -Compress
[IO.File]::WriteAllText($OutFile, $json, [Text.UTF8Encoding]::new($false))
"pest_data.json -> $OutFile"
"  visitas=$($visitOut.Count) sectores=$($sectorOut.Count) plagas=$($pestDefs.Count) catalogo=$($catalog.Count) lecturas=$($readings.Count) puntos=$($points.Count) informes=$($reports.Count) incidencias=$($issues.Count)"
"  tamano=$([Math]::Round($json.Length/1024,1)) KB"
