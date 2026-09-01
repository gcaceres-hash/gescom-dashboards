# Dashboards Gescom — Fusión Lago Puelo / Elebes

Scripts que bajan datos de Gescom y publican los dashboards en Netlify. Antes corrían
solo con el Programador de tareas de Windows (dependen de que la PC esté prendida);
ahora también corren en GitHub Actions, sin depender de ninguna PC.

## Dashboards

| Dashboard | Script | Horario (hora Argentina) |
|---|---|---|
| Pizarra de Rentabilidad | `scripts/pull-venta-rentabilidad.ps1` | cada 4 horas |
| Torre de Cobranzas | `scripts/pull-cuentas-corrientes.ps1` | 07:00 |
| Seguimiento Martín Di Giorno | `scripts/pull-vendedor-digiorno.ps1` | 07:15 |
| Cobertura General | `scripts/pull-cobertura-general.ps1` | 07:30 |
| Clientes Potenciales | `scripts/pull-potenciales.ps1` | 07:45 |

`scripts/stock-dashboard-refresh.ps1` (Torre de control de stock) sigue corriendo
**solo en Windows**, porque genera un archivo local (no se publica a Netlify).

## Configurar los secrets de GitHub (una sola vez)

Entrá a **Settings → Secrets and variables → Actions → New repository secret** en este
repositorio y creá:

- **`GESCOM_CONFIG`**: pegá el contenido completo de tu `config.json` real (usuario y
  clave de Gescom incluidos). Mirá `scripts/config.example.json` para el formato.
- **`NETLIFY_TOKEN`**: tu token personal de Netlify (el mismo que ya usa la PC local).

Sin estos dos secrets, los workflows van a fallar al correr.

## Correr un workflow manualmente

En la pestaña **Actions** de GitHub, elegí el workflow (ej. "Torre de Cobranzas") →
**Run workflow**. Sirve para probar sin esperar al horario programado.

## Uso local (Windows)

1. Copiá `scripts/config.example.json` a `scripts/config.json` y completá tus
   credenciales reales (ese archivo nunca se sube a GitHub, está en `.gitignore`).
2. La variable de entorno `NETLIFY_TOKEN` ya está configurada en esta PC a nivel de
   usuario — las tareas programadas de Windows la usan automáticamente.
3. Los instaladores de tareas programadas (`scripts/install-*.ps1`) registran las
   tareas apuntando a esta carpeta.

## Importante

- Si GitHub Actions y el Programador de tareas de Windows corren **al mismo tiempo**
  para el mismo dashboard, no pasa nada grave (el último deploy gana), pero es más
  prolijo tener uno solo activo por dashboard. Ver con Claude cuál desactivar una vez
  confirmado que Actions funciona bien.
- Los archivos `*-data.json` no se versionan (cambian en cada corrida) — cada
  ejecución los regenera desde cero.
