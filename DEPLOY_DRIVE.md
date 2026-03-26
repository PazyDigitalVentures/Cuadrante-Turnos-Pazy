# Publicar en Drive (sin local)

Objetivo: que cualquier persona abra el sistema desde Google Drive y lo use sin instalar nada.

## 1) Subir archivos a tu carpeta de Drive

Sube estos archivos a tu carpeta compartida:

- `turnos-pazy/index.html`
- `turnos-pazy/styles.css`
- `turnos-pazy/app.js`
- `turnos-pazy/plantilla-turnos-pazy.xlsx`
- `turnos-pazy/apps-script/Code.gs`
- `turnos-pazy/apps-script/appsscript.json`

Carpeta objetivo:

- `1QqodtcRBgpMYJEyJpM454svceGUWpgK3`

## 2) Crear Google Sheet desde la plantilla

1. En Drive, abre `plantilla-turnos-pazy.xlsx` con Google Sheets.
2. Haz `Archivo -> Guardar como Hojas de cálculo de Google`.
3. Copia el `Sheet ID` (de la URL).

La plantilla ya incluye en `Config`:

- `carpetaDriveId = 1QqodtcRBgpMYJEyJpM454svceGUWpgK3`

## 3) Publicar Apps Script Web App

1. Abre [script.google.com](https://script.google.com).
2. Crea un proyecto y pega:
   - `Code.gs`
   - `appsscript.json`
3. Deploy:
   - `Deploy -> New deployment -> Web app`
   - `Execute as`: `Me`
   - `Who has access`: `Anyone with the link` (o “Anyone”)
4. Copia la URL del Web App.

## 4) Abrir el frontend desde Drive

Tienes 2 opciones:

- Opción A (rápida): abrir `index.html` desde Drive con una app de vista HTML (si tu dominio lo permite).
- Opción B (recomendada): hostear `index.html`, `styles.css`, `app.js` en GitHub Pages / Netlify / Cloudflare Pages y usar el mismo Web App de Apps Script.

## 5) Configuración final en la página

En la web, rellena:

- `Web App URL` = URL del deployment de Apps Script
- `Google Sheet ID` = ID del sheet creado en el paso 2

Luego:

1. `Cargar`
2. `Generar`
3. Ajustar cambios si hace falta
4. `Guardar` (Sheets)
5. `Guardar imagen` (se sube a la carpeta de Drive configurada)

## 6) Compartir para que cualquiera lo use

Comparte:

- El enlace de la página (donde esté publicado el HTML)
- El Google Sheet con permiso de edición (si deben guardar cambios)
- Tu carpeta de Drive (si deben ver/descargar PNG)

