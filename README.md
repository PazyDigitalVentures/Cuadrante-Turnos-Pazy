# Turnos Pazy (HTML + Google Sheets)

## Plantilla de Google Sheets

He generado un Excel listo para subir a Google Sheets:

- `turnos-pazy/plantilla-turnos-pazy.xlsx`

Al subirlo, se crearán estas pestañas (tal cual las espera la web y el Apps Script):

- `Comerciales`: `id`, `nombre`, `activo`
- `Vacaciones`: `nombre`, `desde`, `hasta`, `motivo`
- `Turnos`: `weekStart`, `fecha`, `franja`, `tipo`, `modo`, `asignadoA`, `nota`
- `Cambios`: `timestamp`, `weekStart`, `slotId`, `antes`, `despues`, `motivo`, `autor`
- `Config`: `clave`, `valor` (pon aquí `carpetaDriveId`)

### Notas rápidas

- `weekStart` siempre es el **jueves** (formato `YYYY-MM-DD`).
- `franja` es una de: `MANANA`, `TARDE`, `NOCHE`.
- `tipo` es una de: `FIJO`, `BACKUP`.
- `modo` es: `NORMAL` o `TODOS` (si está `TODOS`, la casilla no asigna personas).

## Estructura del proyecto

- Frontend:
  - `turnos-pazy/index.html`
  - `turnos-pazy/styles.css`
  - `turnos-pazy/app.js`
- Apps Script (backend):
  - `turnos-pazy/apps-script/Code.gs`
  - `turnos-pazy/apps-script/appsscript.json`

## Deploy de Apps Script (Web App)

1. Crea un proyecto nuevo en [Google Apps Script](https://script.google.com/).
2. Copia el contenido de `Code.gs` y `appsscript.json`.
3. Guarda el proyecto.
4. Deploy:
   - `Deploy` -> `New deployment` -> tipo `Web app`.
   - `Execute as`: `User accessing the web app` o `Me` (recomendado para empezar: `Me`).
   - `Who has access`: `Anyone` (si usarás HTML local sin login Google).
5. Copia la URL del despliegue y pégala en el campo **Web App URL** del frontend.

## Cómo usar la app web

1. Abre `turnos-pazy/index.html` en el navegador.
2. Rellena:
   - `Web App URL`
   - `Google Sheet ID` (el ID del Sheet subido desde la plantilla).
3. Pulsa `Cargar`.
4. Pulsa `Generar turnos` para rellenar automáticamente de forma equitativa.
5. Ajusta manualmente si hace falta o usa `Posibles cambios`.
6. Pulsa `Guardar en Sheets`.
7. Pulsa `Guardar imagen` para generar PNG y subirlo a Drive.

## Configuración de Drive para PNG

- En la hoja `Config`:
  - `clave = carpetaDriveId`
  - `valor = ID de la carpeta de Drive` donde guardar imágenes.
- Si no pones `carpetaDriveId`, el script guarda en tu carpeta raíz de Drive.
- Cada imagen subida también se registra en la hoja `Imagenes`.

## Vacaciones

- Añade filas en `Vacaciones` con:
  - `nombre`
  - `desde` (`YYYY-MM-DD`)
  - `hasta` (`YYYY-MM-DD`)
- La app excluye automáticamente a esa persona durante esas fechas.

## Branding Pazy

- El CSS ya tiene variables de marca en `styles.css`:
  - `--pazy`, `--pazy2`, `--pazySoft`.
- Cuando compartas el manual de marca (PDF), ajustamos los hex exactos y la tipografía corporativa.

