# Excel de Compras: stock, pedidos y rotación

Arma `Compras - Stock, pedidos y rotación.xlsx` (en el Escritorio) con: días de stock por proveedor y por categoría,
productos sin compra del mes, sugerido de pedido (seguridad A 7 / B 5 / C 2 días), stock valorizado a costo y stock
alto con venta lenta. Se pide con "actualizame el Excel de compras".

Pasos (cada uno deja su resultado en `tmp/`, que no se sube a GitHub):

1. `1-extraer-csv.ps1 -CsvAnual <export anual de Gescom> -CsvReciente <export más nuevo>`: ventas desde el 7/7 por
   artículo y la categoría (familia) de Gescom. Hace falta para completar los 3 meses mientras la base compartida no
   tenga 90 días de historia (arranca el 3/8; desde noviembre alcanza con la base sola, pero la categoría sale siempre de este export).
2. `2-extraer-base.ps1`: stock (depósito PRI), costos, venta desde el 3/8, compras y días de operación de la base compartida.
3. `3-unir-datos.ps1`: une las dos fuentes, convierte los productos que se pesan (la base los cuenta en gramos) y calcula
   la clase ABC por proveedor (A hasta 80% de la venta, B hasta 95%, C el resto).
4. `4-armar-excel.ps1`: genera el Excel con fórmulas. Los parámetros (plazos por proveedor, días de seguridad, umbrales)
   se cambian en la solapa Parámetros y todo se recalcula.

Notas:

- La ventana arranca fija en 2026-07-07 (`ventanaDesde` en el paso 3 y los textos del paso 4): al regenerar, mover esa fecha.
- Usa Excel por COM (local). Los scripts con acentos se guardan en UTF-8 con BOM para que PowerShell 5.1 los lea bien.
- Supuesto clave: el plazo de reposición (7 días por defecto, no es el real de cada proveedor). Mueve mucho el pedido sugerido.
