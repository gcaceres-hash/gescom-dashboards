# Dashboards Gescom — Fusión Lago Puelo / Elebes

Scripts que bajan datos de Gescom y publican los dashboards como sitio estático en
**GitHub Pages** (gratis). Corren solos vía **GitHub Actions**, sin depender de que
ninguna PC esté prendida.

## Dashboards

Una vez activado GitHub Pages (ver más abajo), todo se ve desde:
`https://gcaceres-hash.github.io/gescom-dashboards/`

| Dashboard | Script | Carpeta publicada | Horario (hora Argentina) |
|---|---|---|---|
| Pizarra de Rentabilidad | `scripts/pull-venta-rentabilidad.ps1` | `docs/rentabilidad/` | cada 4 horas |
| Torre de Cobranzas | `scripts/pull-cuentas-corrientes.ps1` | `docs/cobranzas/` | 07:00 |
| Seguimiento Martín Di Giorno | `scripts/pull-vendedor-digiorno.ps1` | `docs/digiorno/` | 07:15 |
| Cobertura General | `scripts/pull-cobertura-general.ps1` | `docs/cobertura/` | 07:30 |
| Clientes Potenciales | `scripts/pull-potenciales.ps1` | `docs/potenciales/` | 07:45 |

`scripts/stock-dashboard-refresh.ps1` (Torre de control de stock) sigue corriendo
**solo en Windows**, porque genera un archivo local para ver en la PC, no un sitio.

## Activar GitHub Pages (una sola vez)

**Settings → Pages → Build and deployment → Source**: elegí **"Deploy from a branch"**,
branch **`main`**, carpeta **`/docs`** → Save.

Tarda 1-2 minutos en quedar activo la primera vez.

## Configurar el secret de GitHub (una sola vez)

**Settings → Secrets and variables → Actions → New repository secret**:

- **`GESCOM_CONFIG`**: pegá el contenido completo de tu `config.json` real (usuario y
  clave de Gescom incluidos). Mirá `scripts/config.example.json` para el formato.

Ya no hace falta ningún secret de Netlify — se dejó de usar.

## Correr un workflow manualmente

En la pestaña **Actions**, elegí el workflow (ej. "Torre de Cobranzas") →
**Run workflow**. Sirve para probar sin esperar al horario programado.

## Uso local (Windows)

1. Copiá `scripts/config.example.json` a `scripts/config.json` y completá tus
   credenciales reales (ese archivo nunca se sube a GitHub, está en `.gitignore`).
2. Los instaladores de tareas programadas (`scripts/install-*.ps1`) registran las
   tareas apuntando a esta carpeta. Si un dashboard ya corre por GitHub Actions,
   no hace falta tener también la tarea de Windows activa para ese mismo dashboard.
3. Una corrida local escribe los archivos en `docs/<dashboard>/` igual que Actions,
   pero **no los sube solos** — hay que hacer `git add`, `git commit` y `git push`
   a mano si corrés algo local y querés publicarlo.

## Importante

- Los archivos `docs/**/data.json` y `docs/**/index.html` SÍ se versionan (a
  diferencia de los viejos `scripts/*-data.json`, que siguen ignorados) — son
  justamente lo que GitHub Pages sirve como sitio.
- Si un dashboard corre tanto por Windows como por GitHub Actions al mismo tiempo,
  no rompe nada (gana el último push), pero conviene dejar uno solo activo por
  dashboard para no generar commits duplicados.
