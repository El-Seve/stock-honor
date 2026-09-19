# make_fixtures.ps1 - genera libros Excel SINTETICOS (sin datos reales) y sus resultados esperados.
#
#   tests/fixtures/synthetic_stock.xlsx   (PERIODO/DIA numericos, como en los .xlsx reales)
#   tests/fixtures/synthetic_stock.xlsb   (PERIODO/DIA como texto, como en los .xlsb reales)
#   tests/golden/{xlsx,xlsb}[_cov].json   (salida del extractor local extract.ps1 = referencia)
#
# Sirven para comprobar que el extractor de la nube (Python) da EXACTAMENTE lo mismo que el
# extractor local (Excel + C#). Requiere Excel instalado; solo se corre en la PC de desarrollo.
# NOTA: ASCII puro (PowerShell 5.1).
param(
    [string]$FixturesDir = (Join-Path $PSScriptRoot "fixtures"),
    [string]$GoldenDir = (Join-Path $PSScriptRoot "golden")
)
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $FixturesDir | Out-Null
New-Item -ItemType Directory -Force -Path $GoldenDir | Out-Null

function UChar([int]$code) { return [char]::ConvertFromUtf32($code) }
$E_ACUTE = UChar 0xC9      # E con tilde
$N_TILDE = UChar 0xD1      # enie
$U_ACUTE = UChar 0xDA
$EMOJI   = UChar 0x1F4F1   # fuera del plano basico (par sustituto)
$NBSP    = UChar 0xA0      # espacio de no separacion

$headers = @('PERIODO','DIA','UBICACIONFISICA','ALMACEN','CODSUBALMACEN','CODIGOORACLE','ITEMS','VALORIZADO','TIPOALMACENAMIENTO','DESPLEGADO','CANAL','ESTADO','MODALIDAD','PROCESO','DISPONIBILIDAD','FAMILIA','MARCAMODELO','PRODUCTO','USO','CONDICION','TAG1','TAG2','TIPO','SUBTIPO','STATUS','BLOQUE','MARCA','CLASIFICACION','GAMA','SUBGAMA','RANGOAGING','SEMANA','ACTIVOREPOSICION','PUNTODEVENTA','ESTADO_PDV','CATEGORIA_EQUIPO','CATEGORIA_ACCE','REGION','DEPARTAMENTO')

# ------------------------------------------------------------ catalogo ficticio
# modelos: lista con repeticiones = pesos para elegir la variante de mayusculas
$catalog = @(
    @{ fam='Moviles';    tipo='Celular';    sub='Smartphone'; modelos=@('Honor Test A1 128GB 5G');                                      productos=@('HONOR TEST A1 128GB NEGRO 5G');                                   cod='TESTA1BK'; uso='Normal' },
    @{ fam='Moviles';    tipo='Celular';    sub='Smartphone'; modelos=@('Honor Test A1 128GB 5G');                                      productos=@(("HONOR TEST A1 128GB AZUL OC" + $E_ACUTE + "ANO 5G")); cod='TESTA1BL'; uso='Normal' },
    @{ fam='Moviles';    tipo='Celular';    sub='Smartphone'; modelos=@('Honor Test B2 256GB');                                         productos=@('HONOR TEST B2 256GB PLATA');                                      cod=12345;      uso='Normal' },
    @{ fam='Moviles';    tipo='Celular';    sub='Smartphone'; modelos=@('Honor Test C3 Plus','Honor Test C3 Plus','Honor Test C3 Plus','Honor Test C3 plus'); productos=@('HONOR TEST C3 PLUS 256GB VERDE','Honor Test C3 Plus 256GB Verde','Honor Test C3 Plus 256GB Verde'); cod='TESTC3GR'; uso='Normal' },
    @{ fam='Moviles';    tipo='Celular';    sub='Smartphone'; modelos=@("Honor <&> E5 'Pro'");                                          productos=@("HONOR <&> E5 'PRO' NEGRO");                                       cod='TESTE5';   uso='Normal' },
    @{ fam='Moviles';    tipo='Celular';    sub='Smartphone'; modelos=@('Honor Test A1 128GB 5G');                                      productos=@('PACK HONOR TEST A1 128GB NEGRO 5G INC EARBUDS');                  cod='PKTESTA1'; uso='Pack' },
    @{ fam='Moviles';    tipo='Celular';    sub='Smartphone'; modelos=@("Honor Test Emoji $EMOJI");                                     productos=@("HONOR TEST EMOJI 64GB $EMOJI");                                   cod='TESTEMO';  uso='Normal' },
    @{ fam='Accesorios'; tipo='AUDIO';      sub='AUDIFONOS';  modelos=@('Honor Test Earbuds');                                          productos=@('HONOR TEST EARBUDS BLANCO');                                      cod='ACCTESTEARBUDWT'; uso='Normal' },
    @{ fam='Accesorios'; tipo='WEARABLES';  sub='SMARTWATCHES'; modelos=@('Honor Test Watch');                                          productos=@('HONOR TEST WATCH NEGRO');                                         cod='ACCTESTWATCHBK'; uso='Normal' },
    @{ fam='Accesorios'; tipo='AUDIO';      sub='PARLANTES';  modelos=@(("Honor Test " + $N_TILDE + "and" + $U_ACUTE + " Speaker"));  productos=@(("HONOR TEST " + $N_TILDE + "AND" + $U_ACUTE + " SPEAKER"));       cod='ACCTESTSPK'; uso='Normal' },
    @{ fam='Accesorios'; tipo='TABLET';     sub='TABLET';     modelos=@('Honor Test Pad X');                                            productos=@('HONOR TEST PAD X 128GB');                                         cod='TESTPADX'; uso='Normal' }
)

$deptos  = @('LIMA', ('LIMA' + ' '), 'CUSCO', 'PIURA', 'ND', 'AREQUIPA')
$canales = @('Tiendas Express', 'Retail', 'Central', 'Islas', 'Tiendas Propias Franquicias')
$rnd = New-Object System.Random 20260919

# PDV ficticios: 40 puntos de venta con estados variados
$pdvCount = 40
function Get-PdvInfo([int]$i) {
    $n = '{0:000}' -f ($i + 1)
    $name = "PDV_TEST_$n"
    if ($i -eq 1) { $name = "PDV_TEST_" + $N_TILDE + "AND" + $U_ACUTE + "_002" }
    if ($i -eq 2) { $name = "PDV_TEST_<&>_'003'" }
    if ($i -eq 3) { $name = "PDV_TEST_004" + $NBSP }          # espacio de no separacion final
    $estado = 'OPERATIVO'
    if ($i -ge 34 -and $i -le 36) { $estado = 'NO OPERATIVO' }
    if ($i -eq 37) { $estado = 'ND' }
    if ($i -eq 38) { $estado = $null }
    $info = @{
        name = $name; ubic = ("001.TEST UBIC " + $n); estado = $estado
        depto = $deptos[$i % $deptos.Count]; canal = $canales[$i % $canales.Count]
        region = $(if ($i % 2 -eq 0) { 'ND' } else { $null })
        onlyNoDisp = ($i -eq 39)         # PDV que solo tiene stock "No Disponible" (universo de cobertura)
    }
    return $info
}

$rows = New-Object System.Collections.Generic.List[object]
function Add-Row($pdv, $marca, $fam, $modelo, $producto, $cod, $tipo, $sub, $estado, $disp, $uso, $qty) {
    $r = New-Object object[] 39
    $r[2] = $pdv.ubic; $r[3] = 'ALM TEST'; $r[5] = $cod; $r[6] = $qty; $r[10] = $pdv.canal; $r[11] = $estado
    $r[14] = $disp; $r[15] = $fam; $r[16] = $modelo; $r[17] = $producto; $r[18] = $uso; $r[22] = $tipo; $r[23] = $sub
    $r[26] = $marca; $r[33] = $pdv.name; $r[34] = $pdv.estado; $r[35] = 'B'; $r[36] = 'B'; $r[37] = $pdv.region; $r[38] = $pdv.depto
    $rows.Add($r)
}

# ---- filas explicitas (orden fijo al inicio) ----
# empate de variantes de MARCAMODELO (2 vs 2): gana la primera vista
$p0 = Get-PdvInfo 0; $p1 = Get-PdvInfo 1; $p4 = Get-PdvInfo 4; $p5 = Get-PdvInfo 5
Add-Row $p0 'Honor' 'Moviles' 'Honor Tie D4' 'HONOR TIE D4 128GB' 'TESTD4' 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' 3
Add-Row $p1 'Honor' 'Moviles' 'Honor Tie D4' 'HONOR TIE D4 128GB' 'TESTD4' 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' 2
Add-Row $p4 'Honor' 'Moviles' 'HONOR TIE D4' 'HONOR TIE D4 128GB' 'TESTD4' 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' 4
Add-Row $p5 'Honor' 'Moviles' 'HONOR TIE D4' 'HONOR TIE D4 128GB' 'TESTD4' 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' 1
# mismo PDV escrito distinto (minusculas + espacio final): debe fusionarse con PDV_TEST_001
$p0v = @{ name = 'pdv_test_001 '; ubic = $p0.ubic; estado = $p0.estado; depto = $p0.depto; canal = $p0.canal; region = $p0.region }
Add-Row $p0v 'Honor' 'Moviles' 'Honor Test A1 128GB 5G' 'HONOR TEST A1 128GB NEGRO 5G' 'TESTA1BK' 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' 4
# rarezas que SI deben pasar el filtro: marca/familia con espacios y mayusculas, "disponible" en minuscula
Add-Row $p0 'HONOR ' ' moviles ' 'Honor Test A1 128GB 5G' 'HONOR TEST A1 128GB NEGRO 5G' 'TESTA1BK' 'celular' 'Smartphone' 'nuevos' 'disponible' 'normal' 5
# cantidades raras: vacia (0), texto numerico, decimal
Add-Row $p1 'Honor' 'Moviles' 'Honor Test B2 256GB' 'HONOR TEST B2 256GB PLATA' 12345 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' $null
Add-Row $p4 'Honor' 'Moviles' 'Honor Test B2 256GB' 'HONOR TEST B2 256GB PLATA' 12345 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' "'7"
Add-Row $p5 'Honor' 'Moviles' 'Honor Test B2 256GB' 'HONOR TEST B2 256GB PLATA' 12345 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' 2.5
# exclusiones que NO deben pasar
Add-Row $p0 'Honor' 'Moviles' 'Honor Test A1 128GB 5G' 'HONOR TEST A1 128GB NEGRO 5G' 'TESTA1BK' 'Tablet' 'Tablet' 'Nuevos' 'Disponible' 'Normal' 9      # Moviles con TIPO != Celular
Add-Row $p0 'Honor' 'Chips-Sims' 'Chip Test' 'CHIP TEST' 'CHIPT' 'Chips' '0' 'Nuevos' 'Disponible' 'Normal' 9                                        # familia fuera de alcance
Add-Row $p0 'Honor' 'Otros' 'Honor Test Otros' 'HONOR TEST OTROS' 'OTR' 'ND' 'ND' 'Nuevos' 'Disponible' 'Normal' 9
Add-Row $p0 'Samsung' 'Moviles' 'Samsung Test S1' 'SAMSUNG TEST S1 NEGRO' 'SAMS1' 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' 9             # otra marca
Add-Row $p0 'Honor' 'Moviles' 'Honor Test A1 128GB 5G' 'HONOR TEST A1 128GB NEGRO 5G' 'TESTA1BK' 'Celular' 'Smartphone' 'Usados' 'Disponible' 'Normal' 9
Add-Row $p0 'Honor' 'Moviles' 'Honor Test A1 128GB 5G' 'HONOR TEST A1 128GB NEGRO 5G' 'TESTA1BK' 'Celular' 'Smartphone' 'Nuevos' 'No Disponible' 'Normal' 9
Add-Row $p0 'Honor' 'Moviles' 'Honor Test A1 128GB 5G' 'HONOR TEST A1 128GB NEGRO 5G' 'TESTA1BK' 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Dummie' 9
Add-Row $p0 'Honor' 'Moviles' 'Honor Test A1 128GB 5G' 'HONOR TEST A1 128GB NEGRO 5G' 'TESTA1BK' 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Livedemo' 9
$pBlank = @{ name = $null; ubic = 'SIN PDV'; estado = 'OPERATIVO'; depto = 'ND'; canal = 'Central'; region = $null }
Add-Row $pBlank 'Honor' 'Moviles' 'Honor Test A1 128GB 5G' 'HONOR TEST A1 128GB NEGRO 5G' 'TESTA1BK' 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' 9   # PDV vacio

# ---- filas aleatorias (semilla fija => reproducible) ----
for ($i = 0; $i -lt $pdvCount; $i++) {
    $pdv = Get-PdvInfo $i
    foreach ($item in $catalog) {
        if ($rnd.NextDouble() -gt 0.45) { continue }
        $copies = 1 + $rnd.Next(2)                 # 1-2 filas para el mismo par (prueba la suma)
        for ($k = 0; $k -lt $copies; $k++) {
            $modelo = $item.modelos[$rnd.Next($item.modelos.Count)]
            $producto = $item.productos[$rnd.Next($item.productos.Count)]
            $disp = 'Disponible'; if ($pdv.onlyNoDisp -or $rnd.NextDouble() -lt 0.08) { $disp = 'No Disponible' }
            $estado = 'Nuevos';   if ($rnd.NextDouble() -lt 0.06) { $estado = 'Usados' }
            $uso = $item.uso
            if ($item.fam -eq 'Moviles') { $x = $rnd.NextDouble(); if ($x -lt 0.05) { $uso = 'Dummie' } elseif ($x -lt 0.07) { $uso = 'Livedemo' } }
            $qty = 1 + $rnd.Next(30)
            Add-Row $pdv 'Honor' $item.fam $modelo $producto $item.cod $item.tipo $item.sub $estado $disp $uso $qty
        }
    }
    if ($rnd.NextDouble() -lt 0.15) {   # ruido de otras marcas / familias
        Add-Row $pdv 'Samsung' 'Moviles' 'Samsung Test S1' 'SAMSUNG TEST S1 NEGRO' 'SAMS1' 'Celular' 'Smartphone' 'Nuevos' 'Disponible' 'Normal' (1 + $rnd.Next(20))
    }
}
Write-Host ("Filas sinteticas: {0}" -f $rows.Count)

# ------------------------------------------------------------ escribir libros
function Build-Workbook([string]$path, [int]$fileFormat, [bool]$periodoAsText) {
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Add()
        while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add() }
        $ws = $wb.Worksheets.Item(1); $ws.Name = "Base Stocks"
        $ws2 = $wb.Worksheets.Item(2); $ws2.Name = "Hoja1"
        $n = $rows.Count
        $arr = [object[,]]::new(($n + 1), 39)
        for ($j = 0; $j -lt 39; $j++) { $arr[0, $j] = $headers[$j] }
        for ($i = 0; $i -lt $n; $i++) {
            $r = $rows[$i]
            for ($j = 2; $j -lt 39; $j++) { $arr[($i + 1), $j] = $r[$j] }
            if ($periodoAsText) { $arr[($i + 1), 0] = '202609'; $arr[($i + 1), 1] = '15' }
            else { $arr[($i + 1), 0] = [double]202609; $arr[($i + 1), 1] = [double]15 }
        }
        if ($periodoAsText) { $ws.Range("A:B").NumberFormat = "@" }
        $ws.Range($ws.Cells(1, 1), $ws.Cells($n + 1, 39)).Value2 = $arr
        $ws2.Cells(1, 1).Value2 = "hoja auxiliar sintetica"
        $wb.SaveAs($path, $fileFormat)
        $wb.Close($false)
    }
    finally { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    Write-Host ("Creado: {0} ({1:n0} KB)" -f $path, ((Get-Item -LiteralPath $path).Length / 1KB))
}

$xlsx = Join-Path $FixturesDir "synthetic_stock.xlsx"
$xlsb = Join-Path $FixturesDir "synthetic_stock.xlsb"
Build-Workbook $xlsx 51 $false     # xlOpenXMLWorkbook
Build-Workbook $xlsb 50 $true      # xlExcelBinaryWorkbook

# ------------------------------------------------------------ resultados esperados (extractor local)
$extract = Join-Path $PSScriptRoot "..\tools\extract.ps1"
foreach ($t in @(@{ f = $xlsx; n = "xlsx" }, @{ f = $xlsb; n = "xlsb" })) {
    & $extract -Path $t.f -OutJson (Join-Path $GoldenDir ($t.n + ".json")) | Out-Null
    & $extract -Path $t.f -OutJson (Join-Path $GoldenDir ($t.n + "_cov.json")) -Coverage | Out-Null
    Write-Host ("Golden: {0}.json y {0}_cov.json" -f $t.n)
}
Write-Host "Listo."
