# Roadmap desde Excel (TypeScript + React + Vite)

Web sencilla que genera un roadmap visual estilo Gantt a partir de un Excel. Los datos que ve todo el mundo provienen de un archivo versionado en el repositorio; además cualquiera puede cargar otro `.xlsx`, `.xls` o `.csv` localmente para previsualizarlo.

## Ejecutar

```bash
npm install
npm run dev
```

Luego abrir la URL indicada por Vite.

## Datos compartidos (persistencia por repositorio)

La aplicación carga automáticamente al iniciar el archivo `public/matriz.xlsx`. Ese archivo se despliega junto con la web, así que **cualquiera que abra la página ve esos datos** sin tener que subir nada.

Para actualizar la información que ve todo el mundo:

1. Reemplaza `public/matriz.xlsx` por tu Excel actualizado (mismo nombre).
2. Haz commit y vuelve a desplegar (build/deploy).

El botón "Cargar Excel" sigue disponible, pero solo cambia la vista en tu navegador; no modifica el archivo del repositorio.

## Formato del Excel

La primera hoja debe tener una fila de encabezados. La aplicación reconoce automáticamente estas columnas (con sus variantes/acentos):

- `Proyecto` (código/ID)
- `Titulo` (nombre del proyecto)
- `Aplicación`
- `Desarrollo` (fecha o rango de la etapa de desarrollo)
- `Testing` (fecha o rango de la etapa de testing/certificación)
- `PaP` (fecha del pase a producción)
- `Responsable`
- `Componentes Impactados`
- `Direccion` (Negocio / Regulatorio / Tecnología)
- `Estado` (Producción, Testing, Desarrollo, Planificado, Refinamiento, Definición, Cancelado)

Las tres columnas de etapa (`Desarrollo`, `Testing`, `PaP`) alimentan la barra del roadmap y sus hitos. Cada celda de fecha admite varios formatos:

- Fecha real de Excel o ISO (`2026-07-15`)
- `dd/mm` o `dd-mm` (por ejemplo `22/09`), usa el año 2026 si no se indica
- Rango `dd/mm - dd/mm` (por ejemplo `15/09 - 21/09`), toma el inicio y el fin
- Nombre de mes suelto (`Julio`, `Agosto`) → se ubica al día 1 de ese mes
- `TBD` o vacío → sin fecha (el proyecto aparece sin barra)

La barra de cada proyecto va desde la primera fecha disponible hasta la última, y se colocan hitos para Desarrollo, Testing y PaP.

## Despliegue en Azure Static Web Apps

El proyecto está listo para desplegarse en Azure Static Web Apps vía GitHub:

- `output_location` = `dist` (lo genera `npm run build`).
- `app_location` = `/` (código en la raíz).
- `api_location` = vacío (no hay backend).
- `public/staticwebapp.config.json` define el fallback de SPA, excluye los assets y el `.xlsx` del rewrite, y fija el MIME correcto del Excel. Vite lo copia a `dist/` en el build.
- `.github/workflows/azure-static-web-apps.yml` contiene el workflow de CI/CD.

Pasos:

1. Sube el proyecto a un repositorio de GitHub (rama `main`).
2. En el portal de Azure: crear recurso **Static Web App** → origen **GitHub** → selecciona el repo y la rama `main`.
3. En "Build Details" usa preset **Custom** con: App location `/`, Api location vacío, Output location `dist`.
4. Azure crea el secreto `AZURE_STATIC_WEB_APPS_API_TOKEN` en el repo y dispara el primer deploy.

Para actualizar los datos: reemplaza `public/matriz.xlsx`, haz commit a `main` y el workflow redespliega solo.

## Notas

- El procesamiento se realiza en el navegador; el Excel no se sube a un backend.
- El roadmap está configurado inicialmente para junio-diciembre de 2026, siguiendo la referencia.
- Para producción se puede agregar selección dinámica de año/meses, mapeo manual de columnas y exportación a PNG/PDF.
