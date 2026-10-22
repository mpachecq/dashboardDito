import { app } from '@azure/functions'

// Azure Function que actúa de intermediario (proxy) entre el navegador y Jira.
// El navegador NO puede llamar a Jira directamente (CORS + el token quedaría
// expuesto). Esta función corre en el servidor, guarda el token en variables de
// entorno y devuelve solo los estados ya limpios.
//
// Variables de entorno necesarias (configurar en Azure, NUNCA en el código):
//   JIRA_BASE_URL  -> https://tuempresa.atlassian.net
//   JIRA_EMAIL     -> tu-correo@empresa.com   (cuenta con permiso de lectura)
//   JIRA_API_TOKEN -> token creado en https://id.atlassian.com/manage-profile/security/api-tokens
//
// Uso desde el frontend:
//   GET /api/jira-status?keys=PROI-392,PROI-424
//   -> { "statuses": { "PROI-392": "En progreso", "PROI-424": "Listo" }, "notFound": [...] }

const MAX_KEYS = 200 // límite de seguridad por petición

app.http('jira-status', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'jira-status',
  handler: async (request, context) => {
    const baseUrl = process.env.JIRA_BASE_URL
    const email = process.env.JIRA_EMAIL
    const token = process.env.JIRA_API_TOKEN

    if (!baseUrl || !email || !token) {
      context.error('Faltan variables de entorno de Jira (JIRA_BASE_URL/JIRA_EMAIL/JIRA_API_TOKEN).')
      return json(500, {
        error: 'El servidor no está configurado para Jira. Revisa las variables de entorno JIRA_BASE_URL, JIRA_EMAIL y JIRA_API_TOKEN.'
      })
    }

    // Lista de issue keys separada por comas: ?keys=PROI-392,PROI-424
    const raw = request.query.get('keys') || ''
    const keys = Array.from(
      new Set(
        raw
          .split(',')
          .map(k => k.trim())
          .filter(Boolean)
      )
    )

    if (keys.length === 0) {
      return json(400, { error: 'Debes enviar al menos un issue key en el parámetro ?keys=' })
    }
    if (keys.length > MAX_KEYS) {
      return json(400, { error: `Demasiados issue keys (máximo ${MAX_KEYS} por petición).` })
    }

    // Autenticación Basic de Jira Cloud: base64("email:token").
    const auth = Buffer.from(`${email}:${token}`).toString('base64')

    // Consultamos con JQL (key in (...)) y pedimos solo el campo status.
    // Es una sola llamada aunque sean muchos issues.
    const jql = `key in (${keys.map(k => `"${k}"`).join(',')})`
    const searchUrl = `${baseUrl.replace(/\/$/, '')}/rest/api/3/search`

    try {
      const res = await fetch(searchUrl, {
        method: 'POST',
        headers: {
          Authorization: `Basic ${auth}`,
          Accept: 'application/json',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          jql,
          fields: ['status'],
          maxResults: MAX_KEYS
        })
      })

      if (!res.ok) {
        const detail = await safeText(res)
        context.error(`Jira respondió ${res.status}: ${detail}`)
        return json(res.status === 401 || res.status === 403 ? 502 : 502, {
          error: `Jira respondió con error ${res.status}. Revisa las credenciales y la URL base.`
        })
      }

      const data = await res.json()
      const statuses = {}
      for (const issue of data.issues ?? []) {
        statuses[issue.key] = issue.fields?.status?.name ?? ''
      }

      // Qué keys pedidos no volvieron (no existen o sin permiso de lectura).
      const notFound = keys.filter(k => !(k in statuses))

      return json(200, { statuses, notFound })
    } catch (err) {
      context.error('Error llamando a Jira:', err)
      return json(502, { error: 'No se pudo contactar con Jira. Inténtalo de nuevo más tarde.' })
    }
  }
})

function json(status, body) {
  return {
    status,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  }
}

async function safeText(res) {
  try {
    return (await res.text()).slice(0, 500)
  } catch {
    return '(sin cuerpo)'
  }
}
