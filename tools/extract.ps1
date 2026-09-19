# extract.ps1 - genera stock_data.json a partir de la pestana "Base Stocks" (Excel).
# Version rapida: Excel solo entrega las columnas necesarias y el procesamiento
# corre en C# compilado (antes: ~9 pasadas en PowerShell, ~4 min; ahora segundos).
#
# Salida: JSON ASCII-only (todo caracter no ASCII y < > & ' van como \uXXXX), asi
# ninguna codificacion puede corromperlo y es seguro dentro de <script>.
#
# NOTA: este archivo debe mantenerse en ASCII puro (PowerShell 5.1 lee .ps1 sin BOM
# como ANSI y los acentos/guiones largos pueden romper el parser).
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$OutJson = (Join-Path $PSScriptRoot "stock_data.json"),
    [switch]$Coverage,          # agrega el bloque "coverage" (universo completo de PDV x modelos)
    [int]$TimeoutSec = 300      # watchdog: si Excel se cuelga, se mata SOLO esta instancia
)

$ErrorActionPreference = "Stop"
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()
function Lap($msg) { Write-Host ("  [{0,5:n1}s] {1}" -f $swTotal.Elapsed.TotalSeconds, $msg) }

if (-not (Test-Path -LiteralPath $Path)) { throw "No existe el archivo: $Path" }

# ---------------------------------------------------------------- C# (compatible con C# 5 / PS 5.1)
if (-not ('StockExtractor' -as [type])) {
    $src = @'
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

public static class StockExtractor
{
    public const int C_PERIODO = 1, C_DIA = 2, C_UBIC = 3, C_CODORACLE = 6, C_ITEMS = 7, C_CANAL = 11,
        C_ESTADO = 12, C_DISP = 15, C_FAM = 16, C_MODELO = 17, C_PROD = 18, C_USO = 19, C_TIPO = 23,
        C_SUBTIPO = 24, C_MARCA = 27, C_PDV = 34, C_ESTADOPDV = 35, C_CATEQ = 36, C_CATACC = 37,
        C_REGION = 38, C_DEPTO = 39;

    public static int[] NeededColumns()
    {
        return new int[] { 1, 2, 3, 6, 7, 11, 12, 15, 16, 17, 18, 19, 23, 24, 27, 34, 35, 36, 37, 38, 39 };
    }

    public class Result
    {
        public int Kept, Products, Pdvs, StockPairs, CovModels, CovPdvs;
        public double TotalQty;
        public string Json;
    }

    static string T(object[,] col, int r)
    {
        object v = col[r, 1];
        if (v == null) return "";
        string s = v as string;
        if (s == null) s = Convert.ToString(v, CultureInfo.InvariantCulture);
        return s.Trim();
    }

    static bool Eq(string a, string b) { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }

    // Variante canonica por valor (case-insensitive): gana la mas frecuente; en empate, la primera vista.
    static Dictionary<string, string> CanonMap(object[,] col, bool[] pass, int rows)
    {
        Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.Ordinal);
        List<string> order = new List<string>();
        for (int r = 2; r <= rows; r++)
        {
            if (!pass[r]) continue;
            string v = T(col, r);
            if (v.Length == 0) continue;
            int c;
            if (counts.TryGetValue(v, out c)) counts[v] = c + 1;
            else { counts[v] = 1; order.Add(v); }
        }
        Dictionary<string, string> canon = new Dictionary<string, string>();
        Dictionary<string, int> best = new Dictionary<string, int>();
        foreach (string k in order)
        {
            string lk = k.ToLowerInvariant();
            int cnt = counts[k];
            int b;
            if (!best.TryGetValue(lk, out b) || cnt > b) { best[lk] = cnt; canon[lk] = k; }
        }
        return canon;
    }

    static string Norm(Dictionary<string, string> map, string s)
    {
        string c;
        if (map.TryGetValue(s.ToLowerInvariant(), out c)) return c;
        return s;
    }

    // ---------- JSON (ASCII-only) ----------
    static void Q(StringBuilder sb, string s)
    {
        sb.Append('"');
        for (int i = 0; i < s.Length; i++)
        {
            char ch = s[i];
            switch (ch)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                default:
                    if (ch < 0x20 || ch > 0x7E || ch == '<' || ch == '>' || ch == '&' || ch == '\'')
                        sb.Append("\\u").Append(((int)ch).ToString("x4"));
                    else sb.Append(ch);
                    break;
            }
        }
        sb.Append('"');
    }

    static string Num(double v)
    {
        if (v == Math.Floor(v) && Math.Abs(v) < 1e15) return ((long)v).ToString(CultureInfo.InvariantCulture);
        return v.ToString("R", CultureInfo.InvariantCulture);
    }

    // PERIODO / DIA: numero si la celda es numerica, texto si viene como texto (.xlsb)
    static void PD(StringBuilder sb, object v)
    {
        if (v == null) { sb.Append("null"); return; }
        if (v is double) { sb.Append(Num((double)v)); return; }
        Q(sb, Convert.ToString(v, CultureInfo.InvariantCulture));
    }

    static void KV(StringBuilder sb, string name, string val, bool first)
    {
        if (!first) sb.Append(',');
        sb.Append('"').Append(name).Append("\":");
        Q(sb, val);
    }

    public static Result Run(object[] cols, int rows, bool coverage)
    {
        if (rows < 2) throw new Exception("La hoja Base Stocks no tiene filas de datos.");

        object[,] cPer = (object[,])cols[C_PERIODO], cDia = (object[,])cols[C_DIA];
        object[,] cUbic = (object[,])cols[C_UBIC], cCod = (object[,])cols[C_CODORACLE], cItems = (object[,])cols[C_ITEMS];
        object[,] cCanal = (object[,])cols[C_CANAL], cEstado = (object[,])cols[C_ESTADO], cDisp = (object[,])cols[C_DISP];
        object[,] cFam = (object[,])cols[C_FAM], cModelo = (object[,])cols[C_MODELO], cProd = (object[,])cols[C_PROD];
        object[,] cUso = (object[,])cols[C_USO], cTipo = (object[,])cols[C_TIPO], cSub = (object[,])cols[C_SUBTIPO];
        object[,] cMarca = (object[,])cols[C_MARCA], cPdv = (object[,])cols[C_PDV], cEstPdv = (object[,])cols[C_ESTADOPDV];
        object[,] cCatEq = (object[,])cols[C_CATEQ], cCatAcc = (object[,])cols[C_CATACC];
        object[,] cRegion = (object[,])cols[C_REGION], cDepto = (object[,])cols[C_DEPTO];

        // ---- Predicado unico de alcance (una sola vez por fila) ----
        // Honor | Moviles+Accesorios | Nuevos | Disponible | PDV OPERATIVO |
        // TIPO=Celular solo para Moviles | USO Normal/Pack
        bool[] pass = new bool[rows + 1];
        for (int r = 2; r <= rows; r++)
        {
            string fam = T(cFam, r).ToLowerInvariant();
            bool isMov = (fam == "moviles");
            if (!isMov && fam != "accesorios") continue;
            if (!Eq(T(cDisp, r), "Disponible")) continue;
            if (!Eq(T(cEstPdv, r), "OPERATIVO")) continue;
            if (!Eq(T(cMarca, r), "honor")) continue;
            if (!Eq(T(cEstado, r), "nuevos")) continue;
            if (isMov && !Eq(T(cTipo, r), "celular")) continue;
            string uso = T(cUso, r).ToLowerInvariant();
            if (uso != "normal" && uso != "pack") continue;
            pass[r] = true;
        }

        Dictionary<string, string> mapFam = CanonMap(cFam, pass, rows);
        Dictionary<string, string> mapTipo = CanonMap(cTipo, pass, rows);
        Dictionary<string, string> mapSub = CanonMap(cSub, pass, rows);
        Dictionary<string, string> mapModelo = CanonMap(cModelo, pass, rows);
        Dictionary<string, string> mapProd = CanonMap(cProd, pass, rows);

        // ---- Agregacion ----
        Dictionary<string, int> productIndex = new Dictionary<string, int>();
        List<string[]> products = new List<string[]>();   // producto, modelo, marca, familia, tipo, subtipo, catEq, catAcc, codOracle
        Dictionary<string, int> pdvIndex = new Dictionary<string, int>();
        List<string[]> pdvs = new List<string[]>();       // pdv, ubicacion, region, departamento, canal
        Dictionary<long, int> pairIndex = new Dictionary<long, int>();
        List<int> pairP = new List<int>(), pairS = new List<int>();
        List<double> pairQ = new List<double>();
        int kept = 0;
        double totalQty = 0;

        for (int r = 2; r <= rows; r++)
        {
            if (!pass[r]) continue;
            string pdv = T(cPdv, r);
            if (pdv.Length == 0) continue;

            string producto = Norm(mapProd, T(cProd, r));
            string pKey = producto.ToLowerInvariant();
            int pIdx;
            if (!productIndex.TryGetValue(pKey, out pIdx))
            {
                products.Add(new string[] {
                    producto, Norm(mapModelo, T(cModelo, r)), T(cMarca, r), Norm(mapFam, T(cFam, r)),
                    Norm(mapTipo, T(cTipo, r)), Norm(mapSub, T(cSub, r)), T(cCatEq, r), T(cCatAcc, r), T(cCod, r) });
                pIdx = products.Count - 1;
                productIndex[pKey] = pIdx;
            }

            string sKey = pdv.ToLowerInvariant();
            int sIdx;
            if (!pdvIndex.TryGetValue(sKey, out sIdx))
            {
                pdvs.Add(new string[] { pdv, T(cUbic, r), T(cRegion, r), T(cDepto, r), T(cCanal, r) });
                sIdx = pdvs.Count - 1;
                pdvIndex[sKey] = sIdx;
            }

            double qty = 0;
            object q = cItems[r, 1];
            if (q is double) qty = (double)q;
            else if (q != null)
            {
                double parsed;
                if (double.TryParse(Convert.ToString(q, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out parsed)) qty = parsed;
            }

            long key = ((long)pIdx << 32) | (uint)sIdx;
            int pi;
            if (pairIndex.TryGetValue(key, out pi)) pairQ[pi] = pairQ[pi] + qty;
            else { pairIndex[key] = pairP.Count; pairP.Add(pIdx); pairS.Add(sIdx); pairQ.Add(qty); }
            totalQty += qty;
            kept++;
        }

        StringBuilder sb = new StringBuilder(1 << 20);
        sb.Append("{\"periodo\":"); PD(sb, cPer[2, 1]);
        sb.Append(",\"dia\":"); PD(sb, cDia[2, 1]);

        sb.Append(",\"products\":[");
        for (int i = 0; i < products.Count; i++)
        {
            string[] p = products[i];
            if (i > 0) sb.Append(',');
            sb.Append('{');
            KV(sb, "producto", p[0], true); KV(sb, "modelo", p[1], false); KV(sb, "marca", p[2], false);
            KV(sb, "familia", p[3], false); KV(sb, "tipo", p[4], false); KV(sb, "subtipo", p[5], false);
            KV(sb, "catEq", p[6], false); KV(sb, "catAcc", p[7], false); KV(sb, "codOracle", p[8], false);
            sb.Append('}');
        }
        sb.Append("],\"pdvs\":[");
        for (int i = 0; i < pdvs.Count; i++)
        {
            string[] p = pdvs[i];
            if (i > 0) sb.Append(',');
            sb.Append('{');
            KV(sb, "pdv", p[0], true); KV(sb, "ubicacion", p[1], false); KV(sb, "region", p[2], false);
            KV(sb, "departamento", p[3], false); KV(sb, "canal", p[4], false);
            sb.Append('}');
        }
        sb.Append("],\"stock\":[");
        for (int i = 0; i < pairP.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append('[').Append(pairP[i]).Append(',').Append(pairS[i]).Append(',').Append(Num(pairQ[i])).Append(']');
        }
        sb.Append(']');

        Result res = new Result();
        res.Kept = kept; res.Products = products.Count; res.Pdvs = pdvs.Count;
        res.StockPairs = pairP.Count; res.TotalQty = totalQty;

        // ---- Cobertura (opcional): TODOS los PDV Honor x que modelos tienen ----
        if (coverage)
        {
            // catalogo de modelos: orden fijo (Moviles primero, luego alfabetico)
            List<string[]> models = new List<string[]>();     // modelo, familia
            HashSet<string> seen = new HashSet<string>();
            for (int i = 0; i < products.Count; i++)
            {
                string mk = products[i][1].ToLowerInvariant();
                if (seen.Add(mk)) models.Add(new string[] { products[i][1], products[i][3] });
            }
            models.Sort(delegate(string[] a, string[] b)
            {
                int ga = (a[1] == "Moviles") ? 0 : 1, gb = (b[1] == "Moviles") ? 0 : 1;
                if (ga != gb) return ga.CompareTo(gb);
                return string.Compare(a[0], b[0], StringComparison.OrdinalIgnoreCase);
            });
            Dictionary<string, int> modelIndex = new Dictionary<string, int>();
            for (int i = 0; i < models.Count; i++) modelIndex[models[i][0].ToLowerInvariant()] = i;

            // universo: cualquier fila Honor con PDV, sin exigir disponibilidad/operatividad/estado/uso
            Dictionary<string, int> univIndex = new Dictionary<string, int>();
            List<string[]> univ = new List<string[]>();       // pdv, ubicacion, departamento, canal, estadoPdv
            for (int r = 2; r <= rows; r++)
            {
                if (!Eq(T(cMarca, r), "honor")) continue;
                string pdv = T(cPdv, r);
                if (pdv.Length == 0) continue;
                string uk = pdv.ToLowerInvariant();
                if (univIndex.ContainsKey(uk)) continue;
                univ.Add(new string[] { pdv, T(cUbic, r), T(cDepto, r), T(cCanal, r), T(cEstPdv, r) });
                univIndex[uk] = univ.Count - 1;
            }
            List<SortedSet<int>> have = new List<SortedSet<int>>();
            for (int i = 0; i < univ.Count; i++) have.Add(new SortedSet<int>());
            for (int r = 2; r <= rows; r++)
            {
                if (!pass[r]) continue;
                string pdv = T(cPdv, r);
                if (pdv.Length == 0) continue;
                int ui, mi;
                if (!univIndex.TryGetValue(pdv.ToLowerInvariant(), out ui)) continue;
                if (!modelIndex.TryGetValue(Norm(mapModelo, T(cModelo, r)).ToLowerInvariant(), out mi)) continue;
                have[ui].Add(mi);
            }

            sb.Append(",\"coverage\":{\"models\":[");
            for (int i = 0; i < models.Count; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append('{'); KV(sb, "modelo", models[i][0], true); KV(sb, "familia", models[i][1], false); sb.Append('}');
            }
            sb.Append("],\"pdvs\":[");
            for (int i = 0; i < univ.Count; i++)
            {
                string[] u = univ[i];
                if (i > 0) sb.Append(',');
                sb.Append('{');
                KV(sb, "pdv", u[0], true); KV(sb, "ubicacion", u[1], false); KV(sb, "departamento", u[2], false);
                KV(sb, "canal", u[3], false); KV(sb, "estadoPdv", u[4], false);
                sb.Append(",\"have\":[");
                bool f = true;
                foreach (int m in have[i]) { if (!f) sb.Append(','); sb.Append(m); f = false; }
                sb.Append("]}");
            }
            sb.Append("]}");
            res.CovModels = models.Count; res.CovPdvs = univ.Count;
        }

        sb.Append('}');
        res.Json = sb.ToString();
        return res;
    }
}
'@
    $refs = @()
    if ($PSVersionTable.PSEdition -ne 'Core') { $refs = @('System.Core') }
    Add-Type -TypeDefinition $src -Language CSharp -ReferencedAssemblies $refs
}
Lap "C# compilado"

# ---------------------------------------------------------------- Excel (solo lectura de columnas)
Add-Type -Namespace Win32 -Name ProcUtil -ErrorAction SilentlyContinue -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);
'@

$excel = $null; $wb = $null; $xlPid = 0; $watchdog = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try { $excel.ScreenUpdating = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 3 } catch {}

    # PID de ESTA instancia (nunca se toca la instancia del usuario)
    $pidOut = [uint32]0
    [void][Win32.ProcUtil]::GetWindowThreadProcessId([System.IntPtr][int]$excel.Hwnd, [ref]$pidOut)
    $xlPid = [int]$pidOut

    # watchdog: si algo se cuelga (dialogo oculto, archivo bloqueado), mata solo esta instancia
    if ($xlPid -gt 0) {
        $watchdog = Start-Process -FilePath "powershell" -WindowStyle Hidden -PassThru -ArgumentList @(
            "-NoProfile", "-Command", "Start-Sleep -Seconds $TimeoutSec; Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue")
    }

    $wb = $excel.Workbooks.Open($Path, 0, $true)
    try { $excel.Calculation = -4135 } catch {}
    Lap "Excel abrio el libro"

    $sheetNames = @(); foreach ($s in $wb.Worksheets) { $sheetNames += $s.Name }
    if ($sheetNames -notcontains "Base Stocks") { throw "El archivo no tiene la hoja 'Base Stocks'. Hojas: $($sheetNames -join ', ')" }
    Write-Host ("Hojas del libro: {0}" -f ($sheetNames -join ', '))
    $ws = $wb.Worksheets.Item("Base Stocks")

    # --- validar encabezados (39 columnas, mismo orden) ---
    $expected = @('PERIODO','DIA','UBICACIONFISICA','ALMACEN','CODSUBALMACEN','CODIGOORACLE','ITEMS','VALORIZADO','TIPOALMACENAMIENTO','DESPLEGADO','CANAL','ESTADO','MODALIDAD','PROCESO','DISPONIBILIDAD','FAMILIA','MARCAMODELO','PRODUCTO','USO','CONDICION','TAG1','TAG2','TIPO','SUBTIPO','STATUS','BLOQUE','MARCA','CLASIFICACION','GAMA','SUBGAMA','RANGOAGING','SEMANA','ACTIVOREPOSICION','PUNTODEVENTA','ESTADO_PDV','CATEGORIA_EQUIPO','CATEGORIA_ACCE','REGION','DEPARTAMENTO')
    $hdr = $ws.Range($ws.Cells(1, 1), $ws.Cells(1, 39)).Value2
    $bad = @()
    for ($i = 1; $i -le 39; $i++) {
        $got = ([string]$hdr[1, $i]).Trim()
        if ($got -ne $expected[$i - 1]) { $bad += ("col {0}: '{1}' (esperado '{2}')" -f $i, $got, $expected[$i - 1]) }
    }
    if ($bad.Count -gt 0) { throw ("Encabezados de Base Stocks distintos al esquema:`n  " + ($bad -join "`n  ")) }

    $ur = $ws.UsedRange
    $rows = $ur.Row + $ur.Rows.Count - 1
    Lap ("Encabezados OK (39 columnas), filas usadas: {0}" -f $rows)

    # --- leer solo las columnas necesarias ---
    $cols = New-Object 'object[]' 40
    foreach ($c in [StockExtractor]::NeededColumns()) {
        $cols[$c] = $ws.Range($ws.Cells(1, $c), $ws.Cells($rows, $c)).Value2
    }
    Lap "Columnas leidas desde Excel"
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($excel) { try { $excel.Quit() } catch {} }
    if ($ws) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) }
    if ($wb) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
    if ($excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($watchdog) { try { Stop-Process -Id $watchdog.Id -Force -ErrorAction SilentlyContinue } catch {} }
    if ($xlPid -gt 0) {   # si Excel no cerro solo, se cierra unicamente el PID propio
        Start-Sleep -Milliseconds 800
        $p = Get-Process -Id $xlPid -ErrorAction SilentlyContinue
        if ($p) { try { $p.Kill() } catch {} }
    }
}
Lap "Excel cerrado"

# ---------------------------------------------------------------- procesar (C#) y escribir
$res = [StockExtractor]::Run($cols, [int]$rows, [bool]$Coverage.IsPresent)
Lap "Datos procesados"

[System.IO.File]::WriteAllText($OutJson, $res.Json, (New-Object System.Text.UTF8Encoding($false)))

Write-Host ("Filtered rows kept: {0}" -f $res.Kept)
Write-Host ("Distinct products: {0}" -f $res.Products)
Write-Host ("Distinct PDVs: {0}" -f $res.Pdvs)
Write-Host ("Distinct stock pairs: {0}" -f $res.StockPairs)
if ($Coverage) { Write-Host ("Cobertura: modelos={0} universo PDV={1}" -f $res.CovModels, $res.CovPdvs) }
Write-Host ("Wrote JSON to {0} ({1} KB)" -f $OutJson, [math]::Round((Get-Item -LiteralPath $OutJson).Length / 1KB, 1))
Lap "Listo"
