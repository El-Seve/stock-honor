# actualizar_stock.ps1 - UN SOLO LLAMADO para actualizar el panel Stock Honor.
#
#   powershell -ExecutionPolicy Bypass -File tools\actualizar_stock.ps1 -Path "<archivo .xlsb/.xlsx>"
#
# Hace todo, en orden, y aborta con un mensaje claro si algo no cuadra:
#   1. Guardas previas (rama main, arbol limpio, archivo existe, plantillas presentes)
#   2. Copia el Excel a una carpeta temporal (evita el bloqueo si lo tienes abierto)
#   3. Extrae los datos (extract.ps1: valida encabezados y filtra a Honor)
#   4. Verifica los datos y compara contra lo publicado (corte mas antiguo / cifras raras)
#   5. Ensambla index.html (plantillas del repo = unica fuente de verdad) y lo verifica
#   6. Commit + push a main
#   7. Espera el despliegue en GitHub Pages y confirma que la pagina en vivo coincide
#   8. Limpia temporales y muestra el resumen
#
# -DryRun : hace 1-5 sin tocar el repo (construye a un HTML temporal) y compara con lo publicado.
# -Force  : salta las guardas (corte mas antiguo, cifras muy distintas, rama != main, arbol sucio).
#
# NOTA: mantener este archivo en ASCII puro (PowerShell 5.1 lee .ps1 sin BOM como ANSI).
param(
    [Parameter(Mandatory = $true, Position = 0)][string]$Path,
    [switch]$DryRun,
    [switch]$Force,
    [switch]$Coverage,
    [string]$Trailer = "",
    [int]$DeployTimeoutSec = 300
)

$ErrorActionPreference = "Stop"
$sw = [System.Diagnostics.Stopwatch]::StartNew()
function Step($m) { Write-Host ("[{0,5:n1}s] {1}" -f $sw.Elapsed.TotalSeconds, $m) }

$Tools     = $PSScriptRoot
$Repo      = Split-Path -Parent $Tools
$Live      = "https://el-seve.github.io/stock-honor/"
$Short     = "https://tinyurl.com/Consulta-Stock-Honor"
$JsonPath  = Join-Path $Tools "stock_data.json"
$IndexPath = Join-Path $Repo "index.html"
$Work      = Join-Path $env:TEMP "stock_honor_work"
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# ---------------------------------------------------------------- utilidades
function Invoke-Git {
    $old = $ErrorActionPreference; $ErrorActionPreference = "Continue"
    try { $o = & git -C $Repo @args 2>&1 | ForEach-Object { "$_" } } finally { $ErrorActionPreference = $old }
    if ($LASTEXITCODE -ne 0) { throw ("git " + ($args -join " ") + " fallo:`n" + ($o -join "`n")) }
    return $o
}

function Get-Summary([string]$jsonText) {
    $d = $jsonText | ConvertFrom-Json
    $units  = @{ Moviles = 0.0; Accesorios = 0.0 }
    $models = @{ Moviles = @{}; Accesorios = @{} }
    $pdvs   = @{ Moviles = @{}; Accesorios = @{} }
    $allM = @{}; $allP = @{}
    foreach ($row in $d.stock) {
        $prod = $d.products[[int]$row[0]]
        $f = [string]$prod.familia
        if (-not $units.ContainsKey($f)) { continue }
        $units[$f] += [double]$row[2]
        $models[$f][[string]$prod.modelo] = 1; $pdvs[$f][[int]$row[1]] = 1
        $allM[[string]$prod.modelo] = 1; $allP[[int]$row[1]] = 1
    }
    $per = [string]$d.periodo; $dia = ([string]$d.dia).PadLeft(2, "0")
    $s = New-Object PSObject -Property @{
        Cut        = [int]($per + $dia)
        Label      = ("{0}/{1}/{2}" -f $dia, $per.Substring(4, 2), $per.Substring(0, 4))
        UnitsMov   = $units["Moviles"];  UnitsAcc = $units["Accesorios"]
        UnitsAll   = $units["Moviles"] + $units["Accesorios"]
        ModelsMov  = $models["Moviles"].Count; ModelsAcc = $models["Accesorios"].Count; ModelsAll = $allM.Count
        PdvMov     = $pdvs["Moviles"].Count;   PdvAcc = $pdvs["Accesorios"].Count;     PdvAll = $allP.Count
        Products   = $d.products.Count; Pairs = $d.stock.Count
        HasCoverage = [bool]$d.coverage
    }
    $s | Add-Member -NotePropertyName Fingerprint -NotePropertyValue ("{0}|{1}|{2}|{3}|{4}|{5}|{6}" -f $s.Cut, $s.Pairs, $s.Products, $s.PdvAll, $s.UnitsAll, $s.UnitsMov, $s.ModelsAll)
    return $s
}

function Get-EmbeddedJson([string]$html) {
    $m = [regex]::Match($html, '(?s)<script id="stock-data" type="application/json">\s*(.*?)\s*</script>')
    if ($m.Success) { return $m.Groups[1].Value } else { return $null }
}

function Get-LiveSummary {
    $url = $Live + "index.html?n=" + [guid]::NewGuid().ToString("N")
    $r = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 60 -Headers @{ "Cache-Control" = "no-cache" }
    $j = Get-EmbeddedJson $r.Content
    if (-not $j) { return $null }
    return (Get-Summary $j)
}

function Show-Table($s) {
    Write-Host ""
    Write-Host ("  {0,-13}{1,10}{2,9}{3,7}" -f "Familia", "Unidades", "Modelos", "PDV")
    Write-Host ("  {0,-13}{1,10:n0}{2,9}{3,7}" -f "Todas",       $s.UnitsAll, $s.ModelsAll, $s.PdvAll)
    Write-Host ("  {0,-13}{1,10:n0}{2,9}{3,7}" -f "Smartphones", $s.UnitsMov, $s.ModelsMov, $s.PdvMov)
    Write-Host ("  {0,-13}{1,10:n0}{2,9}{3,7}" -f "Accesorios",  $s.UnitsAcc, $s.ModelsAcc, $s.PdvAcc)
    Write-Host ""
}

# ---------------------------------------------------------------- flujo
try {
    Write-Host "=== Actualizar stock Honor ==="
    if ($DryRun) { Write-Host "(modo DryRun: no se toca el repo ni se publica nada)" }

    # 1. guardas previas
    if (-not (Test-Path -LiteralPath $Path)) { throw "No existe el archivo: $Path" }
    foreach ($f in @("part1.html", "part2.html", "extract.ps1", "build_standalone.ps1")) {
        if (-not (Test-Path -LiteralPath (Join-Path $Tools $f))) { throw "Falta tools\$f en el repo." }
    }
    [void](Invoke-Git rev-parse --is-inside-work-tree)
    $branch = (Invoke-Git branch --show-current) -join ""
    $dirty  = @(Invoke-Git status --porcelain --untracked-files=no)
    if (-not $DryRun) {
        if ($branch -ne "main" -and -not $Force) { throw "El repo esta en la rama '$branch', no en 'main'. Cambia a main (o usa -Force)." }
        if ($dirty.Count -gt 0 -and -not $Force) { throw ("Hay cambios sin commitear en el repo (se abortan para no mezclarlos):`n  " + ($dirty -join "`n  ")) }
    }
    $part2 = Get-Content -LiteralPath (Join-Path $Tools "part2.html") -Raw
    $needCov = $Coverage.IsPresent -or ($part2 -match "RAW\.coverage")
    Step ("Guardas OK (rama {0}, cobertura {1})" -f $branch, $(if ($needCov) { "SI" } else { "no" }))

    # lo publicado hoy (para comparar); si no hay red, se sigue con aviso
    $liveBefore = $null
    try { $liveBefore = Get-LiveSummary } catch { Write-Host ("  AVISO: no pude leer lo publicado ({0}); sigo sin comparar." -f $_.Exception.Message) }
    if ($liveBefore) { Step ("Publicado hoy: corte {0}, {1:n0} u." -f $liveBefore.Label, $liveBefore.UnitsAll) }

    # 2. copia de trabajo (evita el bloqueo si el archivo esta abierto en Excel)
    if (Test-Path -LiteralPath $Work) { Remove-Item -LiteralPath $Work -Recurse -Force -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Force -Path $Work | Out-Null
    $tmp = Join-Path $Work ("bd_corte" + [System.IO.Path]::GetExtension($Path))
    Copy-Item -LiteralPath $Path -Destination $tmp -Force
    Step ("Copia de trabajo lista ({0:n1} MB)" -f ((Get-Item -LiteralPath $tmp).Length / 1MB))

    # 3. extraer datos
    & (Join-Path $Tools "extract.ps1") -Path $tmp -OutJson $JsonPath -Coverage:$needCov
    Step "Datos extraidos"

    # 4. verificar datos y comparar con lo publicado
    $local = Get-Summary (Get-Content -LiteralPath $JsonPath -Raw)
    if ($local.UnitsAll -le 0 -or $local.PdvAll -le 0 -or $local.ModelsAll -le 0) { throw "El resultado quedo vacio (0 unidades/PDV/modelos). Revisa el archivo." }
    Step ("Corte {0}: {1:n0} u., {2} modelos, {3} PDV" -f $local.Label, $local.UnitsAll, $local.ModelsAll, $local.PdvAll)
    if ($liveBefore) {
        if ($local.Cut -lt $liveBefore.Cut -and -not $Force) { throw ("El archivo es del {0}, MAS ANTIGUO que lo publicado ({1}). Usa -Force si es intencional." -f $local.Label, $liveBefore.Label) }
        if ($local.Cut -eq $liveBefore.Cut) { Write-Host ("  AVISO: mismo corte que lo publicado ({0}); se tratara como version corregida." -f $local.Label) }
        $ratio = $local.UnitsAll / [math]::Max(1, $liveBefore.UnitsAll)
        if (($ratio -lt 0.5 -or $ratio -gt 2.0) -and -not $Force) { throw ("Cifras muy distintas a lo publicado ({0:n0} vs {1:n0} u.). Verifica que sea el archivo correcto o usa -Force." -f $local.UnitsAll, $liveBefore.UnitsAll) }
    }

    # 5. ensamblar y verificar el HTML
    $outHtml = if ($DryRun) { Join-Path $Work "index_dryrun.html" } else { $IndexPath }
    & (Join-Path $Tools "build_standalone.ps1") -OutFile $outHtml | Out-Null
    $html = [System.IO.File]::ReadAllText($outHtml)
    $emb = Get-EmbeddedJson $html
    if (-not $emb) { throw "El HTML generado no contiene el bloque de datos." }
    $built = Get-Summary $emb
    if ($built.Fingerprint -ne $local.Fingerprint) { throw "El HTML generado no coincide con los datos extraidos (huella distinta)." }
    foreach ($id in @("statsRow", "familiaFilter", "deptoFilter", "canalFilter", "pdvFilter", "modeloFilter", "shareBtn", "resultsArea", "snapshotDate")) {
        if ($html -notmatch ('id="' + $id + '"')) { throw "Al HTML generado le falta el elemento #$id (plantilla danada)." }
    }
    if (([regex]::Matches($html, "<script")).Count -ne ([regex]::Matches($html, "</script>")).Count) { throw "El HTML generado tiene etiquetas <script> sin cerrar." }
    Step ("index.html ensamblado y verificado ({0:n0} KB)" -f ($html.Length / 1KB))

    if ($DryRun) {
        $same = ($liveBefore -and $liveBefore.Fingerprint -eq $local.Fingerprint)
        Show-Table $local
        Write-Host ("DryRun OK. Coincide con lo publicado: {0}" -f $(if ($same) { "SI" } else { "NO (habria cambios)" }))
        Write-Host ("Tiempo total: {0:n0} s" -f $sw.Elapsed.TotalSeconds)
        return
    }

    # 6. commit + push
    [void](Invoke-Git add index.html)
    $staged = @(Invoke-Git diff --cached --name-only)
    if ($staged.Count -eq 0) {
        Step "index.html no cambio: no hay nada que publicar."
    }
    else {
        $msg = ("Actualiza datos al corte {0}`n`n{1:n0} u. (Smartphones {2:n0} / Accesorios {3:n0}), {4} modelos, {5} PDV. Mismo alcance de filtros; solo cambia el corte.`n`n{6}`n" -f `
            $local.Label, $local.UnitsAll, $local.UnitsMov, $local.UnitsAcc, $local.ModelsAll, $local.PdvAll, $Trailer)
        $msg = $msg.TrimEnd() + "`n"     # sin lineas en blanco sobrantes si no hay trailer
        $msgFile = Join-Path $Work "commit_msg.txt"
        [System.IO.File]::WriteAllText($msgFile, $msg, (New-Object System.Text.UTF8Encoding($false)))
        [void](Invoke-Git commit -q -F $msgFile)
        [void](Invoke-Git push -q origin main)
        Step "Commit + push a main hechos"

        # 7. esperar despliegue y confirmar en vivo
        $deadline = (Get-Date).AddSeconds($DeployTimeoutSec)
        $ok = $false
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds 10
            try {
                $lv = Get-LiveSummary
                if ($lv -and $lv.Fingerprint -eq $local.Fingerprint) { $ok = $true; break }
            } catch {}
            Write-Host "  ... esperando el despliegue de GitHub Pages"
        }
        if (-not $ok) { throw ("Publicado en GitHub, pero NO pude confirmar el despliegue en {0} s. Revisa {1} en un minuto." -f $DeployTimeoutSec, $Live) }
        Step "Despliegue confirmado en vivo"
    }

    # 8. resumen
    Write-Host ""
    Write-Host ("=== LISTO: corte {0} publicado ===" -f $local.Label)
    Show-Table $local
    Write-Host ("En vivo: {0}   (corto: {1})" -f $Live, $Short)
    Write-Host ("Tiempo total: {0:n0} s" -f $sw.Elapsed.TotalSeconds)
}
catch {
    Write-Host ""
    Write-Host ("ERROR: {0}" -f $_.Exception.Message)
    exit 1
}
finally {
    if (Test-Path -LiteralPath $Work) { Remove-Item -LiteralPath $Work -Recurse -Force -ErrorAction SilentlyContinue }
}
