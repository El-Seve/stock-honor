# Stock Honor — consulta por modelo y punto de venta

Panel web estático para consultar el **stock disponible de productos Honor** (smartphones y accesorios) de Entel Perú, filtrando en cascada y compartiendo el resultado por WhatsApp.

🔗 **App en vivo:** https://el-seve.github.io/stock-honor/

## Qué hace

- **Filtros en cascada:** Familia → Departamento → Canal → Punto de venta → Modelo. Cada nivel acota al siguiente.
- **Resultado progresivo:**
  - Sin modelo elegido → lista los modelos con stock del ámbito (con desglose por SKU).
  - Con un punto de venta elegido → todos sus modelos.
  - Con un modelo elegido → en qué puntos de venta está disponible.
- **Compartir por WhatsApp:** arma el mensaje del ámbito seleccionado y abre WhatsApp para elegir contacto.

## Alcance de los datos

Generado desde la pestaña **Base Stocks** del Excel Honor (Oracle). Filtros de alcance aplicados en `tools/extract.ps1`:

- Marca **Honor**.
- **ESTADO** = Nuevos.
- **FAMILIA** ∈ {Moviles, Accesorios} — seleccionable con el combo Familia.
- **TIPO** = Celular **solo para Moviles** (Accesorios: AUDIO/TABLET/WEARABLES).
- **USO** ∈ {Normal, Pack} (excluye equipos Dummie de exhibición y Livedemo).
- Solo stock **Disponible** en puntos de venta **OPERATIVO**.
- Nivel de detalle: **PRODUCTO** (SKU con color/capacidad), agrupado por modelo.

> ⚠️ Los datos de stock están incrustados en `index.html`. Este repositorio es **público**; el enlace se comparte solo con el equipo.

## Cómo actualizar (cada nuevo Excel)

Los datos son una foto fija; se refrescan con **un solo llamado** (requiere Excel instalado, usa COM):

```powershell
powershell -ExecutionPolicy Bypass -File tools/actualizar_stock.ps1 -Path "C:/ruta/archivo.xlsb"
```

O arrastra el archivo `.xlsb`/`.xlsx` sobre `tools/actualizar_stock.cmd`.

El script: valida el archivo (hoja Base Stocks y 39 columnas), extrae y filtra a Honor, verifica los datos contra lo publicado (aborta si el corte es más antiguo o las cifras son muy distintas), ensambla `index.html`, hace commit + push a `main`, espera el despliegue de GitHub Pages y confirma que la página en vivo coincide. Tarda ~1 min.

Opciones: `-DryRun` (todo salvo publicar), `-Force` (salta las guardas).

### Actualizar desde el celular (sin PC)

El Excel **no se sube a este repo** (es público y el archivo puede traer otras marcas). Se sube desde el navegador del
celular al buzón **privado** `El-Seve/stock-honor-inbox`, carpeta `inbox/`: una GitHub Action extrae los datos con Python
(sin Excel), los verifica igual que el flujo local (aborta si el corte es más antiguo o las cifras son muy distintas),
publica el resultado aquí con una *deploy key* limitada a este repo y confirma el despliegue. Los pasos para el celular
están en el README del repo privado.

- `tools/cloud/extract_cloud.py`: extractor en Python; replica **exactamente** `tools/extract.ps1`.
- `tools/cloud/selftest.py` + `.github/workflows/selftest.yml`: comprueban que ambos extractores dan el **mismo resultado byte a byte**
  sobre libros **sintéticos** (`tests/fixtures`, esperados en `tests/golden`, generados con `tests/make_fixtures.ps1`).
  La Action de publicación corre este autotest antes de cada corrida y no publica si falla.
- Si se cambia la lógica de filtro/extracción hay que cambiarla **en los dos extractores** y regenerar `tests/golden` (`make_fixtures.ps1`).
- Cotejo con datos reales: ambos flujos imprimen una **huella de datos** (`corte|filas|productos|PDV|unidades|...`); el mismo Excel
  procesado en la nube y con `-DryRun` local debe dar la misma huella.

## Estructura

| Archivo | Rol |
|---|---|
| `index.html` | App completa (datos incrustados). Servido por GitHub Pages. |
| `tools/actualizar_stock.ps1` / `.cmd` | **Un solo llamado**: extrae, verifica, ensambla, publica y confirma el despliegue. |
| `tools/extract.ps1` | Lee la pestaña Base Stocks (solo las columnas necesarias) y genera `stock_data.json` con procesamiento en C# (segundos). |
| `tools/cloud/extract_cloud.py`, `selftest.py` | Extractor en Python (sin Excel) para la Action del celular y su autotest contra libros sintéticos. |
| `tests/` | Libros sintéticos (`fixtures`), resultados esperados (`golden`) y su generador (`make_fixtures.ps1`). |
| `tools/build_standalone.ps1` | Ensambla `index.html` a partir de las plantillas + datos. |
| `tools/part1.html`, `tools/part2.html` | Plantillas de estilo/markup y de lógica. |
