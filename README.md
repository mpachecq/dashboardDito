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

## Sincronización con Jira

El botón **"Sincronizar con Jira"** (arriba a la derecha) actualiza la columna **Estado Jira** de cada proyecto con el estado real tomado de Jira.

### ¿Por qué hay un backend (carpeta `api/`)?

El navegador **no puede** llamar a la API de Jira directamente por dos motivos:

1. **CORS**: Jira Cloud bloquea las llamadas hechas desde un navegador en otro dominio.
2. **Seguridad**: la autenticación con Jira usa un API token. Si estuviera en el código del frontend, cualquiera que abra la página podría verlo.

Por eso hay una **Azure Function** (`api/jira-status`) que actúa de intermediario: el navegador le pide los estados, y la función (con el token guardado de forma segura en variables de entorno del servidor) consulta Jira y devuelve solo los estados. Como vive en la misma Static Web App, se accede en `/api/jira-status` sin problemas de CORS.

La función usa el endpoint vigente de Jira Cloud `POST /rest/api/3/search/jql` (el antiguo `/rest/api/3/search` fue retirado por Atlassian y devuelve `410 Gone`).

### Cómo mapea los códigos

- Usa el **código padre** (`PROI-xxx`, primera línea de la columna Proyecto) como *issue key* de Jira.
- Si la fila tiene un **código hijo** (segunda línea), también lo consulta y su estado tiene prioridad para esa fila.
- Los códigos que no existan en Jira (o sin permiso de lectura) se reportan en el mensaje de resultado.

> Si tus códigos en el roadmap **no** coinciden con los issue keys reales de Jira (por ejemplo, en Jira son `ABC-123`), hay que mapearlos antes de consultar. Avísame y lo ajusto.

### Configuración (una sola vez)

1. **Crea un API token de Jira** en https://id.atlassian.com/manage-profile/security/api-tokens
2. **En el portal de Azure** → tu Static Web App → *Configuration* (o *Environment variables*), agrega estas tres variables:
   - `JIRA_BASE_URL` → `https://integratelperu.atlassian.net`
   - `JIRA_EMAIL` → el correo de la cuenta con permiso de lectura en Jira
   - `JIRA_API_TOKEN` → el token del paso 1
3. Guarda. La API ya puede autenticarse con Jira. **El token nunca se sube al repositorio.**

### Probar en local

Necesitas [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local) (`func`).

```bash
# 1. Instala dependencias de la API
cd api && npm install && cd ..

# 2. Copia el ejemplo de settings y pon tus credenciales reales
cp api/local.settings.json.example api/local.settings.json
#    edita api/local.settings.json con tu JIRA_BASE_URL / JIRA_EMAIL / JIRA_API_TOKEN

# 3. Levanta web + API juntas con el emulador de Static Web Apps
npm install -g @azure/static-web-apps-cli   # si no lo tienes
npm run build
swa start dist --api-location api
```

`local.settings.json` está en `.gitignore`: tus credenciales no se versionan.

## Despliegue en Azure Static Web Apps

El proyecto está listo para desplegarse en Azure Static Web Apps vía GitHub:

- `output_location` = `dist` (lo genera `npm run build`).
- `app_location` = `/` (código en la raíz).
- `api_location` = `api` (Azure Functions que consultan Jira).
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
