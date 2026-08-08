# Stock Honor — consulta por modelo y punto de venta

Panel web estático para consultar el **stock disponible de productos Honor** (smartphones y accesorios) de Entel Perú, filtrando en cascada y compartiendo el resultado por WhatsApp.

🔗 **App en vivo:** https://el-seve.github.io/stock-honor/

## Qué hace

- **Filtros en cascada:** Departamento → Canal → Punto de venta → Modelo. Cada nivel acota al siguiente.
- **Filtro de familia:** Smartphones / Accesorios / Otros.
- **Resultado progresivo:**
  - Sin modelo elegido → lista los modelos con stock del ámbito (con desglose por SKU).
  - Con un punto de venta elegido → todos sus modelos.
  - Con un modelo elegido → en qué puntos de venta está disponible.
- **Compartir por WhatsApp:** arma el mensaje del ámbito seleccionado y abre WhatsApp para elegir contacto.

## Alcance de los datos

Generado desde la pestaña **Base Stocks** del Excel `BD Oracle2` (Oracle). Se incluye:

- Marca **Honor** únicamente.
- Familias **Moviles, Accesorios y Otros**.
- Solo stock **Disponible** en puntos de venta **OPERATIVO**.
- Nivel de detalle: **PRODUCTO** (SKU con color/capacidad), agrupado por modelo.

> ⚠️ Los datos de stock están incrustados en `index.html`. Este repositorio es **público**; el enlace se comparte solo con el equipo.

## Cómo actualizar (cada nuevo Excel)

Los datos son una foto fija; para refrescarlos hay que regenerar `index.html` y hacer push.

1. Coloca el nuevo `.xlsb` y ajusta la ruta en `tools/extract.ps1` (parámetro `-Path`).
2. Ejecuta el generador (requiere Excel instalado — usa COM):
   ```powershell
   powershell -ExecutionPolicy Bypass -File tools/extract.ps1
   powershell -ExecutionPolicy Bypass -File tools/build_standalone.ps1
   ```
   `extract.ps1` filtra y agrega el stock a `stock_data.json`; `build_standalone.ps1` ensambla `index.html`.
3. Commit + push a `main`. GitHub Pages se actualiza en ~1 minuto.

## Estructura

| Archivo | Rol |
|---|---|
| `index.html` | App completa (datos incrustados). Servido por GitHub Pages. |
| `tools/extract.ps1` | Lee el Excel (pestaña Base Stocks) y genera `stock_data.json`. |
| `tools/build_standalone.ps1` | Ensambla `index.html` a partir de las plantillas + datos. |
| `tools/part1.html`, `tools/part2.html` | Plantillas de estilo/markup y de lógica. |
