# Assembles a standalone index.html (full HTML document) for GitHub Pages
# from part1.html (title+style+body), stock_data.json, and part2.html (scripts).
param(
    [string]$Dir = $PSScriptRoot,
    [string]$OutFile = (Join-Path $PSScriptRoot "..\index.html")
)
$ErrorActionPreference = "Stop"

$part1 = [System.IO.File]::ReadAllText((Join-Path $Dir "part1.html"))
$data  = [System.IO.File]::ReadAllText((Join-Path $Dir "stock_data.json"))
$part2 = [System.IO.File]::ReadAllText((Join-Path $Dir "part2.html"))

# Split part1 at the end of </style>: head goes to <head>, the rest (body markup) to <body>.
$marker = "</style>"
$idx = $part1.IndexOf($marker)
if ($idx -lt 0) { throw "No </style> found in part1.html" }
$headInner = $part1.Substring(0, $idx + $marker.Length)   # includes <title>, <style>...</style>
$bodyInner = $part1.Substring($idx + $marker.Length)        # the <div class="wrap">... markup

# bodyInner already ends with the opening `<script id="stock-data" ...>` tag (from part1),
# and part2.html begins with the matching `</script>` — so concatenate bodyInner + data + part2
# exactly like the Claude-artifact build, just wrapped in a proper head/body skeleton.
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('<!doctype html>')
[void]$sb.AppendLine('<html lang="es">')
[void]$sb.AppendLine('<head>')
[void]$sb.AppendLine('<meta charset="utf-8">')
[void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
[void]$sb.AppendLine('<meta name="robots" content="noindex, nofollow">')
[void]$sb.AppendLine('<meta name="referrer" content="no-referrer">')
[void]$sb.AppendLine($headInner)
[void]$sb.AppendLine('</head>')
[void]$sb.AppendLine('<body>')
[void]$sb.Append($bodyInner)
[void]$sb.Append($data)
[void]$sb.AppendLine()
[void]$sb.Append($part2)
[void]$sb.AppendLine()
[void]$sb.AppendLine('</body>')
[void]$sb.AppendLine('</html>')

$outDir = Split-Path $OutFile -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
[System.IO.File]::WriteAllText($OutFile, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
Write-Output "Wrote $OutFile ($([math]::Round((Get-Item $OutFile).Length/1KB,1)) KB)"
