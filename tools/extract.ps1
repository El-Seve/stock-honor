param(
    [string]$Path = "C:\Users\IseverinoS\Downloads\Bases_Stock&Movs_Honor_CONF.xlsx",
    [string]$OutJson = (Join-Path $PSScriptRoot "stock_data.json")
)

$ErrorActionPreference = "Stop"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $wb = $excel.Workbooks.Open($Path, 0, $true)
    $ws = $wb.Worksheets.Item("Base Stocks")
    $usedRange = $ws.UsedRange
    $rows = $usedRange.Rows.Count
    $data = $usedRange.Value2

    # Column indices (1-based)
    $COL_PERIODO=1; $COL_DIA=2; $COL_CODORACLE=6; $COL_ITEMS=7
    $COL_CANAL=11; $COL_ESTADO=12; $COL_DISPONIBILIDAD=15
    $COL_FAMILIA=16; $COL_MARCAMODELO=17; $COL_PRODUCTO=18; $COL_USO=19
    $COL_TIPO=23; $COL_SUBTIPO=24
    $COL_MARCA=27; $COL_CATEQ=36; $COL_CATACC=37
    $COL_PDV=34; $COL_ESTADOPDV=35
    $COL_REGION=38; $COL_DEPTO=39; $COL_UBIC=3

    # ---- Filtros de alcance (todos case-insensitive) ----
    $famAllowedLower   = @{ "moviles"=$true; "accesorios"=$true }  # Moviles y Accesorios
    $marcaAllowedLower = "honor"                              # solo Honor
    $estadoAllowedLower= "nuevos"                             # ESTADO = Nuevos
    $tipoMovilLower    = "celular"                            # TIPO=Celular SOLO aplica a Moviles
    $usoAllowedLower   = @{ "normal"=$true; "pack"=$true }    # USO Normal/Pack (excluye Dummie y Livedemo)

    function TrimStr($v) {
        if ($null -eq $v) { return "" }
        return ([string]$v).Trim()
    }

    # Predicado unico de alcance, usado en ambas pasadas
    function RowPasses($r) {
        $fam = (TrimStr $data[$r, $COL_FAMILIA]).ToLowerInvariant()
        if ($fam -eq "" -or -not $famAllowedLower.ContainsKey($fam)) { return $false }
        if ((TrimStr $data[$r, $COL_DISPONIBILIDAD]) -ne "Disponible") { return $false }
        if ((TrimStr $data[$r, $COL_ESTADOPDV]) -ne "OPERATIVO") { return $false }
        if ((TrimStr $data[$r, $COL_MARCA]).ToLowerInvariant() -ne $marcaAllowedLower) { return $false }
        if ((TrimStr $data[$r, $COL_ESTADO]).ToLowerInvariant() -ne $estadoAllowedLower) { return $false }
        # TIPO=Celular solo se exige a Moviles (Accesorios son AUDIO/TABLET/WEARABLES)
        if ($fam -eq "moviles" -and (TrimStr $data[$r, $COL_TIPO]).ToLowerInvariant() -ne $tipoMovilLower) { return $false }
        $uso = (TrimStr $data[$r, $COL_USO]).ToLowerInvariant()
        if (-not $usoAllowedLower.ContainsKey($uso)) { return $false }
        return $true
    }

    # ---- Pass 1: build case-normalization maps (ordinal, pick most frequent variant) ----
    function BuildCanonicalMap($colIdx) {
        $counts = New-Object 'System.Collections.Generic.Dictionary[string,int]'([System.StringComparer]::Ordinal)
        for ($r=2; $r -le $rows; $r++) {
            if (-not (RowPasses $r)) { continue }
            $v = TrimStr $data[$r, $colIdx]
            if ($v -eq "") { continue }
            if ($counts.ContainsKey($v)) { $counts[$v]++ } else { $counts[$v] = 1 }
        }
        $canonical = @{}  # lowercase -> canonical original-case string
        $best = @{}       # lowercase -> best count so far
        foreach ($kv in $counts.GetEnumerator()) {
            $lk = $kv.Key.ToLowerInvariant()
            if (-not $best.ContainsKey($lk) -or $kv.Value -gt $best[$lk]) {
                $best[$lk] = $kv.Value
                $canonical[$lk] = $kv.Key
            }
        }
        return $canonical
    }

    $mapFamilia = BuildCanonicalMap $COL_FAMILIA
    $mapTipo = BuildCanonicalMap $COL_TIPO
    $mapSubtipo = BuildCanonicalMap $COL_SUBTIPO
    $mapModelo = BuildCanonicalMap $COL_MARCAMODELO
    $mapProducto = BuildCanonicalMap $COL_PRODUCTO

    function Norm($map, $v) {
        $s = TrimStr $v
        $lk = $s.ToLowerInvariant()
        if ($map.ContainsKey($lk)) { return $map[$lk] }
        return $s
    }

    # ---- Pass 2: aggregate ----
    $productIndex = @{}   # normalized producto (lowercase) -> idx
    $productList = New-Object System.Collections.Generic.List[object]
    $pdvIndex = @{}
    $pdvList = New-Object System.Collections.Generic.List[object]
    $stockMap = New-Object 'System.Collections.Generic.Dictionary[string,double]'([System.StringComparer]::Ordinal)

    $periodo = $data[2, $COL_PERIODO]
    $dia = $data[2, $COL_DIA]

    $kept = 0
    for ($r = 2; $r -le $rows; $r++) {
        if (-not (RowPasses $r)) { continue }
        $famRaw = TrimStr $data[$r, $COL_FAMILIA]

        $fam = Norm $mapFamilia $famRaw
        $producto = Norm $mapProducto $data[$r, $COL_PRODUCTO]
        $marcamodelo = Norm $mapModelo $data[$r, $COL_MARCAMODELO]
        $tipo = Norm $mapTipo $data[$r, $COL_TIPO]
        $subtipo = Norm $mapSubtipo $data[$r, $COL_SUBTIPO]

        $pdv = TrimStr $data[$r, $COL_PDV]
        if ($pdv -eq "") { continue }

        $qty = $data[$r, $COL_ITEMS]
        if ($null -eq $qty) { $qty = 0 }

        $pKey = $producto.ToLowerInvariant()
        if (-not $productIndex.ContainsKey($pKey)) {
            $obj = [PSCustomObject]@{
                producto = $producto
                modelo = $marcamodelo
                marca = TrimStr $data[$r, $COL_MARCA]
                familia = $fam
                tipo = $tipo
                subtipo = $subtipo
                catEq = TrimStr $data[$r, $COL_CATEQ]
                catAcc = TrimStr $data[$r, $COL_CATACC]
                codOracle = TrimStr $data[$r, $COL_CODORACLE]
            }
            $productList.Add($obj)
            $productIndex[$pKey] = $productList.Count - 1
        }
        $pIdx = $productIndex[$pKey]

        $sKey = $pdv.ToLowerInvariant()
        if (-not $pdvIndex.ContainsKey($sKey)) {
            $obj = [PSCustomObject]@{
                pdv = $pdv
                ubicacion = TrimStr $data[$r, $COL_UBIC]
                region = TrimStr $data[$r, $COL_REGION]
                departamento = TrimStr $data[$r, $COL_DEPTO]
                canal = TrimStr $data[$r, $COL_CANAL]
            }
            $pdvList.Add($obj)
            $pdvIndex[$sKey] = $pdvList.Count - 1
        }
        $sIdx = $pdvIndex[$sKey]

        $mapKey = "$pIdx|$sIdx"
        if ($stockMap.ContainsKey($mapKey)) {
            $stockMap[$mapKey] += $qty
        } else {
            $stockMap[$mapKey] = $qty
        }
        $kept++
    }

    Write-Output "Filtered rows kept: $kept"
    Write-Output "Distinct products: $($productList.Count)"
    Write-Output "Distinct PDVs: $($pdvList.Count)"
    Write-Output "Distinct stock pairs: $($stockMap.Count)"

    $stockArr = New-Object System.Collections.Generic.List[object]
    foreach ($kv in $stockMap.GetEnumerator()) {
        $parts = $kv.Key -split '\|'
        $stockArr.Add(@([int]$parts[0], [int]$parts[1], $kv.Value))
    }

    # ---- Pass 3: Cobertura - TODOS los PDV Honor (universo completo) x que modelos tienen ----
    # Universo: cualquier fila con MARCA=Honor, sin exigir Disponible/Operativo/Estado/Uso
    # (asi tambien aparecen tiendas cerradas o sin stock vendible). "Tiene" un modelo usa
    # la MISMA definicion de disponibilidad que el resto del panel (RowPasses).
    $modelSeen = @{}
    $modelListRaw = New-Object System.Collections.Generic.List[object]
    foreach ($p in $productList) {
        $mk = $p.modelo.ToLowerInvariant()
        if (-not $modelSeen.ContainsKey($mk)) {
            $modelSeen[$mk] = $true
            $modelListRaw.Add([PSCustomObject]@{ modelo = $p.modelo; familia = $p.familia })
        }
    }
    $modelListOrdered = New-Object System.Collections.Generic.List[object]
    $modelIndex = @{}
    foreach ($m in ($modelListRaw | Sort-Object @{Expression={ if ($_.familia -eq 'Moviles') { 0 } else { 1 } }}, modelo)) {
        $modelListOrdered.Add($m)
        $modelIndex[$m.modelo.ToLowerInvariant()] = $modelListOrdered.Count - 1
    }

    $univIndex = @{}   # pdv(lower) -> idx
    $univList = New-Object System.Collections.Generic.List[object]
    for ($r = 2; $r -le $rows; $r++) {
        if ((TrimStr $data[$r, $COL_MARCA]).ToLowerInvariant() -ne $marcaAllowedLower) { continue }
        $pdvU = TrimStr $data[$r, $COL_PDV]
        if ($pdvU -eq "") { continue }
        $ukey = $pdvU.ToLowerInvariant()
        if (-not $univIndex.ContainsKey($ukey)) {
            $uobj = [PSCustomObject]@{
                pdv = $pdvU
                ubicacion = TrimStr $data[$r, $COL_UBIC]
                departamento = TrimStr $data[$r, $COL_DEPTO]
                canal = TrimStr $data[$r, $COL_CANAL]
                estadoPdv = TrimStr $data[$r, $COL_ESTADOPDV]
            }
            $univList.Add($uobj)
            $univIndex[$ukey] = $univList.Count - 1
        }
    }

    $have = New-Object 'System.Collections.Generic.Dictionary[int,System.Collections.Generic.HashSet[int]]'
    for ($r = 2; $r -le $rows; $r++) {
        if (-not (RowPasses $r)) { continue }
        $pdvU = TrimStr $data[$r, $COL_PDV]
        if ($pdvU -eq "") { continue }
        $ukey = $pdvU.ToLowerInvariant()
        if (-not $univIndex.ContainsKey($ukey)) { continue }
        $uIdx = $univIndex[$ukey]
        $mKey = (Norm $mapModelo $data[$r, $COL_MARCAMODELO]).ToLowerInvariant()
        if (-not $modelIndex.ContainsKey($mKey)) { continue }
        $mIdx = $modelIndex[$mKey]
        if (-not $have.ContainsKey($uIdx)) { $have[$uIdx] = New-Object 'System.Collections.Generic.HashSet[int]' }
        [void]$have[$uIdx].Add($mIdx)
    }

    $coveragePdvs = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $univList.Count; $i++) {
        $u = $univList[$i]
        $haveArr = @()
        if ($have.ContainsKey($i)) { $haveArr = @($have[$i] | Sort-Object) }
        $cobj = [PSCustomObject]@{
            pdv = $u.pdv
            ubicacion = $u.ubicacion
            departamento = $u.departamento
            canal = $u.canal
            estadoPdv = $u.estadoPdv
            have = $haveArr
        }
        $coveragePdvs.Add($cobj)
    }
    Write-Output "Cobertura: modelos=$($modelListOrdered.Count) universo PDV=$($coveragePdvs.Count)"

    $result = [PSCustomObject]@{
        periodo = $periodo
        dia = $dia
        products = $productList
        pdvs = $pdvList
        stock = $stockArr
        coverage = [PSCustomObject]@{
            models = $modelListOrdered
            pdvs = $coveragePdvs
        }
    }

    $json = $result | ConvertTo-Json -Depth 6 -Compress
    [System.IO.File]::WriteAllText($OutJson, $json)
    Write-Output "Wrote JSON to $OutJson ($([math]::Round((Get-Item $OutJson).Length/1MB,2)) MB)"
}
finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
