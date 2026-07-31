# Minimal PDF text extractor for the Biofuturo "Manifold" reports (Word-generated PDFs).
# Inflates FlateDecode content streams, replays enough of the text state machine to
# rebuild visual lines, and inserts a space when the horizontal gap implies one.
param([string]$Path, [string]$OutFile)

$latin = [Text.Encoding]::GetEncoding(28591)

function Inflate([byte[]]$data) {
  try {
    $ms = New-Object IO.MemoryStream(,$data); $ms.Position = 2
    $ds = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Decompress)
    $out = New-Object IO.MemoryStream; $ds.CopyTo($out); $ds.Dispose(); $ms.Dispose()
    return $out.ToArray()
  } catch { return $null }
}

$bytes = [IO.File]::ReadAllBytes($Path)
$s = $latin.GetString($bytes)

$chunks = @()
$pos = 0
while ($true) {
  $i = $s.IndexOf('stream', $pos)
  if ($i -lt 0) { break }
  if ($i -ge 3 -and $s.Substring($i-3,3) -eq 'end') { $pos = $i + 6; continue }
  $dictStart = $s.LastIndexOf(' obj', $i)
  $dict = if ($dictStart -gt 0) { $s.Substring($dictStart, $i - $dictStart) } else { '' }
  $d = $i + 6
  if ($s[$d] -eq "`r") { $d++ }
  if ($s[$d] -eq "`n") { $d++ }
  $e = $s.IndexOf('endstream', $d)
  if ($e -lt 0) { break }
  $len = $e - $d
  $isFont = $dict -match '/FontFile|/Length1|/Subtype\s*/(Type1C|CIDFontType0C|TrueType|Image)'
  if ($dict -match 'FlateDecode' -and -not $isFont) {
    $raw = New-Object byte[] $len
    [Array]::Copy($bytes, $d, $raw, 0, $len)
    $inf = Inflate $raw
    if ($inf) {
      $txt = $latin.GetString($inf)
      $printable = ([regex]::Matches($txt, '[\x20-\x7E\r\n\t]')).Count
      if ($txt.Length -gt 0 -and ($printable / $txt.Length) -gt 0.92 -and
          ($txt -match '\bBT\b' -or $txt -match 'beginbfchar|beginbfrange')) { $chunks += ,$txt }
    }
  }
  $pos = $e + 9
}

# --- ToUnicode CMaps (merged; these PDFs only use them for a couple of symbol fonts)
$cmap = @{}
foreach ($c in $chunks) {
  if ($c -notmatch 'beginbfchar|beginbfrange') { continue }
  foreach ($m in [regex]::Matches($c, '(?s)beginbfchar(.*?)endbfchar')) {
    foreach ($p in [regex]::Matches($m.Groups[1].Value, '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>')) {
      $dst = $p.Groups[2].Value
      $ch = -join (0..([int]($dst.Length/4) - 1) | ForEach-Object { [char][Convert]::ToInt32($dst.Substring($_*4,4),16) })
      $cmap[$p.Groups[1].Value.ToUpper()] = $ch
    }
  }
  foreach ($m in [regex]::Matches($c, '(?s)beginbfrange(.*?)endbfrange')) {
    foreach ($p in [regex]::Matches($m.Groups[1].Value, '<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>')) {
      $lo = [Convert]::ToInt32($p.Groups[1].Value,16); $hi = [Convert]::ToInt32($p.Groups[2].Value,16)
      $base = [Convert]::ToInt32($p.Groups[3].Value.Substring(0,[Math]::Min(4,$p.Groups[3].Value.Length)),16)
      if ($hi - $lo -gt 5000) { continue }
      for ($k=$lo; $k -le $hi; $k++) { $cmap[('{0:X4}' -f $k)] = [char]($base + ($k - $lo)) }
    }
  }
}

function DecodeLiteral([string]$lit) {
  $sb = New-Object Text.StringBuilder
  for ($i=0; $i -lt $lit.Length; $i++) {
    $c = $lit[$i]
    if ($c -eq '\') {
      $i++; if ($i -ge $lit.Length) { break }
      $n = $lit[$i]
      switch ($n) {
        'n' { [void]$sb.Append(' ') } 'r' { [void]$sb.Append(' ') } 't' { [void]$sb.Append(' ') }
        'b' { } 'f' { }
        '(' { [void]$sb.Append('(') } ')' { [void]$sb.Append(')') } '\' { [void]$sb.Append('\') }
        default {
          if ($n -match '[0-7]') {
            $oct = "$n"
            while ($oct.Length -lt 3 -and ($i+1) -lt $lit.Length -and $lit[$i+1] -match '[0-7]') { $i++; $oct += $lit[$i] }
            [void]$sb.Append([char][Convert]::ToInt32($oct,8))
          } else { [void]$sb.Append($n) }
        }
      }
    } else { [void]$sb.Append($c) }
  }
  return $sb.ToString()
}

function DecodeHex([string]$hex) {
  $hex = ($hex -replace '\s','').ToUpper()
  if ($hex.Length % 4 -ne 0) { $hex = $hex.PadRight([int][Math]::Ceiling($hex.Length/4)*4, '0') }
  $sb = New-Object Text.StringBuilder
  for ($i=0; $i -lt $hex.Length; $i+=4) {
    $code = $hex.Substring($i,4)
    if ($cmap.ContainsKey($code)) { [void]$sb.Append($cmap[$code]) }
    else { [void]$sb.Append([char][Convert]::ToInt32($code,16)) }
  }
  return $sb.ToString()
}

$num = '-?[\d]*\.?[\d]+'
$tokRx = "(?s)" +
  "\((?<lit>(?:\\.|[^\\()])*)\)|" +
  "<(?<hex>[0-9A-Fa-f][0-9A-Fa-f\s]*)>|" +
  "(?<tm>$num)\s+($num)\s+($num)\s+($num)\s+(?<tmx>$num)\s+(?<tmy>$num)\s+Tm|" +
  "(?<tdx>$num)\s+(?<tdy>$num)\s+(?<tdop>Td|TD)|" +
  "(?<tstar>T\*)|(?<bt>\bBT\b)|(?<et>\bET\b)|" +
  "/(?<font>F\d+)\s+(?<size>$num)\s+Tf|" +
  "(?<lead>$num)\s+TL"

# Each Word text run is its own BT/Tm/TJ/ET block with an absolute position, so lines are
# rebuilt by grouping consecutive runs that share a baseline (Y) and concatenating by X.
# Spaces are already present as explicit "( )" runs; only column jumps need a separator.
function Inv([string]$v) { [double]::Parse($v, [Globalization.CultureInfo]::InvariantCulture) }

$lines = New-Object Collections.Generic.List[string]
$page = 0
foreach ($c in $chunks) {
  if ($c -notmatch '\bBT\b') { continue }
  # drop marked-content property strings (e.g. /Span<</Lang(es-CL)/MCID 3>>BDC) so the
  # language tags of accessibility-tagged PDFs do not leak into the text
  $c = [regex]::Replace($c, '/(Lang|ActualText|Alt|TU|E)\s*\((?:\\.|[^\\()])*\)', '')
  $page++
  $lines.Add("=== PAGE $page ===")
  $x = 0.0; $y = 0.0; $lastY = [double]::NaN; $size = 11.0; $leading = 13.0
  $runStartX = 0.0; $runChars = 0; $newRun = $false
  $cur = New-Object Text.StringBuilder
  foreach ($m in [regex]::Matches($c, $tokRx)) {
    if ($m.Groups['lit'].Success -or $m.Groups['hex'].Success) {
      $piece = if ($m.Groups['lit'].Success) { DecodeLiteral $m.Groups['lit'].Value } else { DecodeHex $m.Groups['hex'].Value }
      if ($piece.Length -eq 0) { continue }
      if ($newRun) {
        # column jump on the same baseline (table label -> value) vs. justification slack
        $estEnd = $runStartX + ($runChars * $size * 0.5)
        if ($cur.Length -gt 0 -and ($x - $estEnd) -gt 25) { [void]$cur.Append("`t") }
        $runStartX = $x; $runChars = 0; $newRun = $false
      }
      [void]$cur.Append($piece)
      $runChars += $piece.Length
      continue
    }
    if ($m.Groups['tm'].Success) {
      $x = Inv $m.Groups['tmx'].Value; $newY = Inv $m.Groups['tmy'].Value
    } elseif ($m.Groups['tdop'].Success) {
      $x = $x + (Inv $m.Groups['tdx'].Value); $newY = $y + (Inv $m.Groups['tdy'].Value)
      if ($m.Groups['tdop'].Value -eq 'TD') { $leading = -(Inv $m.Groups['tdy'].Value) }
    } elseif ($m.Groups['tstar'].Success) {
      $newY = $y - $leading
    } elseif ($m.Groups['font'].Success) {
      $size = Inv $m.Groups['size'].Value; continue
    } elseif ($m.Groups['lead'].Success) {
      $leading = Inv $m.Groups['lead'].Value; continue
    } else { continue }   # BT / ET do not break a line here

    if (-not [double]::IsNaN($lastY) -and [Math]::Abs($newY - $lastY) -gt 1.0) {
      if ($cur.Length -gt 0) { $lines.Add($cur.ToString().Trim()); [void]$cur.Clear() }
      $runStartX = $x; $runChars = 0
    }
    $y = $newY; $lastY = $newY; $newRun = $true
  }
  if ($cur.Length -gt 0) { $lines.Add($cur.ToString().Trim()) }
}

$text = (($lines | Where-Object { $_ -ne '' }) -join "`n")
if ($OutFile) { [IO.File]::WriteAllText($OutFile, $text, [Text.UTF8Encoding]::new($false)); "wrote $($text.Length) chars -> $OutFile" }
else { $text }
