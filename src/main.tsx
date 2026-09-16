import React, { useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import * as XLSX from 'xlsx'
import './styles.css'

type RawRow = Record<string, unknown>

type Status =
  | 'produccion'
  | 'testing'
  | 'desarrollo'
  | 'planificado'
  | 'refinamiento'
  | 'definicion'
  | 'cancelado'
  | 'sin-estado'

type PhaseKey = 'refinamiento' | 'desarrollo' | 'testing' | 'pap'

type Phase = {
  key: PhaseKey
  label: string
  start: Date
  end: Date
  // true cuando la celda tenía un rango real (inicio distinto de fin)
  isRange: boolean
}

type Project = {
  key: string
  id: string
  name: string
  application: string
  status: Status
  statusLabel: string
  stage: string
  responsible: string
  devTeam: string
  components: string
  direction: string
  start: Date | null
  end: Date | null
  pap: Date | null
  phases: Phase[]
}

const PHASE_LABELS: Record<PhaseKey, string> = {
  refinamiento: 'Refinamiento',
  desarrollo: 'Desarrollo',
  testing: 'Testing',
  pap: 'PaP'
}

const MONTHS = ['Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
// Meses cubiertos por el timeline (índice de mes JS): Junio(5) .. Diciembre(11)
const FIRST_MONTH = 5
const MONTH_COUNT = MONTHS.length

const HEADER_ALIASES: Record<string, string[]> = {
  id: ['proyecto', 'id', 'codigo', 'código', 'project id', 'project'],
  name: ['titulo', 'título', 'nombre', 'project name', 'descripcion', 'descripción'],
  application: ['aplicacion', 'aplicación', 'app', 'canal', 'application'],
  status: ['estado', 'status'],
  responsible: ['responsable', 'owner', 'líder', 'lider'],
  devTeam: ['resp. nttdata', 'resp nttdata', 'nttdata', 'desarrollador', 'desarrolladores', 'equipo'],
  components: ['componentes impactados', 'componentes', 'components'],
  direction: ['direccion', 'dirección', 'area', 'área'],
  refinement: ['refinamiento', 'refinement'],
  development: ['desarrollo', 'development', 'dev'],
  testing: ['testing', 'certificacion', 'certificación', 'qa'],
  pap: ['pap', 'pase a produccion', 'pase a producción', 'produccion', 'producción', 'fin']
}

const MONTH_NAMES: Record<string, number> = {
  enero: 0, febrero: 1, marzo: 2, abril: 3, mayo: 4, junio: 5,
  julio: 6, agosto: 7, septiembre: 8, setiembre: 8, octubre: 9,
  noviembre: 10, diciembre: 11
}
const CURRENT_YEAR = 2026

function normalize(s: string) {
  return s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim()
}

function findColumn(headers: string[], aliases: string[]) {
  const normalized = headers.map(normalize)
  return aliases.map(normalize).reduce<string | undefined>((found, alias) => {
    if (found) return found
    const i = normalized.findIndex(h => h === alias || h.includes(alias) || alias.includes(h))
    return i >= 0 ? headers[i] : undefined
  }, undefined)
}

// Interpreta un valor de celda como una única fecha. Soporta:
// - Fechas reales de Excel (Date) o números seriales
// - "dd/mm" o "dd/mm/aaaa" (usa año actual si falta)
// - "dd-mm"
// - Nombres de mes ("Julio", "Agosto") -> día 1 de ese mes
// Devuelve null para "TBD", vacío o no reconocible.
function parseSingleDate(value: unknown): Date | null {
  if (value instanceof Date && !isNaN(value.getTime())) {
    // Excel con cellDates puede introducir desfase de zona horaria; normalizamos
    // usando componentes UTC para quedarnos con el día calendario correcto.
    return new Date(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate())
  }
  if (typeof value === 'number' && isFinite(value)) {
    const d = XLSX.SSF.parse_date_code(value)
    if (d) return new Date(d.y, d.m - 1, d.d)
  }
  if (typeof value !== 'string') return null

  const s = value.trim()
  if (!s || /^(tbd|na|n\/a)$/i.test(s) || s === '-') return null

  // Nombre de mes suelto -> día 1
  const monthOnly = MONTH_NAMES[normalize(s)]
  if (monthOnly !== undefined) return new Date(CURRENT_YEAR, monthOnly, 1)

  // ISO aaaa-mm-dd
  const isoMatch = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/)
  if (isoMatch) {
    return new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]))
  }

  // dd/mm o dd-mm (con año opcional)
  const dm = s.match(/^(\d{1,2})[-/](\d{1,2})(?:[-/](\d{2,4}))?$/)
  if (dm) {
    const day = Number(dm[1])
    const month = Number(dm[2]) - 1
    const year = dm[3] ? (dm[3].length === 2 ? 2000 + Number(dm[3]) : Number(dm[3])) : CURRENT_YEAR
    const d = new Date(year, month, day)
    if (!isNaN(d.getTime())) return d
  }

  const fallback = new Date(s)
  return isNaN(fallback.getTime()) ? null : fallback
}

// Interpreta una celda que puede ser una fecha única o un rango
// "dd/mm - dd/mm" / "16-10 - 01/11". Devuelve inicio y fin.
function parseDateCell(value: unknown): { start: Date | null; end: Date | null } {
  if (typeof value === 'string') {
    const parts = value.split(/\s*[–-]\s*(?=\d{1,2}[\/-]\d{1,2})/)
    if (parts.length >= 2) {
      const start = parseSingleDate(parts[0])
      const end = parseSingleDate(parts[parts.length - 1])
      if (start || end) return { start: start ?? end, end: end ?? start }
    }
  }
  const single = parseSingleDate(value)
  return { start: single, end: single }
}

// Compatibilidad: fecha única (usada para hitos)
function parseDate(value: unknown): Date | null {
  return parseSingleDate(value)
}

function formatDate(d: Date | null) {
  return d ? d.toLocaleDateString('es-PE', { day: '2-digit', month: '2-digit' }) : ''
}

const STATUS_LABELS: Record<Status, string> = {
  produccion: 'Producción',
  testing: 'Testing',
  desarrollo: 'Desarrollo',
  planificado: 'Planificado',
  refinamiento: 'Refinamiento',
  definicion: 'Definición',
  cancelado: 'Cancelado',
  'sin-estado': 'Sin estado'
}

function statusFrom(value: unknown): Status {
  const s = normalize(String(value ?? ''))
  if (!s) return 'sin-estado'
  if (s.includes('produccion') || s.includes('entregado') || s.includes('pap')) return 'produccion'
  if (s.includes('testing') || s.includes('certificac') || s.includes('qa')) return 'testing'
  if (s.includes('desarrollo')) return 'desarrollo'
  if (s.includes('planificado') || s.includes('planeado')) return 'planificado'
  if (s.includes('refinamiento')) return 'refinamiento'
  if (s.includes('definicion') || s.includes('definición')) return 'definicion'
  if (s.includes('cancelado') || s.includes('cancel')) return 'cancelado'
  return 'sin-estado'
}

function cell(row: RawRow, col: string | undefined): unknown {
  return col ? row[col] : undefined
}

function text(row: RawRow, col: string | undefined, fallback = ''): string {
  const v = cell(row, col)
  const s = v == null ? '' : String(v).replace(/\r\n|\r|\n/g, ' ').trim()
  return s || fallback
}

// Como text(), pero descarta marcadores de "sin valor" como "NA" o "N/A".
function optionalText(row: RawRow, col: string | undefined): string {
  const s = text(row, col)
  return /^(na|n\/a)$/i.test(s) ? '' : s
}

function excelToProjects(rows: RawRow[]): Project[] {
  if (!rows.length) return []
  const headers = Object.keys(rows[0])
  const cols = Object.fromEntries(
    Object.entries(HEADER_ALIASES).map(([key, aliases]) => [key, findColumn(headers, aliases)])
  ) as Record<string, string | undefined>

  return rows
    .filter(row => text(row, cols.id) || text(row, cols.name))
    .map((row, index) => {
      const cells: { key: PhaseKey; range: { start: Date | null; end: Date | null } }[] = [
        { key: 'refinamiento', range: parseDateCell(cell(row, cols.refinement)) },
        { key: 'desarrollo', range: parseDateCell(cell(row, cols.development)) },
        { key: 'testing', range: parseDateCell(cell(row, cols.testing)) },
        { key: 'pap', range: parseDateCell(cell(row, cols.pap)) }
      ]

      // Cada etapa con fecha se convierte en una fase con inicio y fin.
      // Si la celda trae un rango real (inicio distinto de fin) se marca isRange.
      const phases: Phase[] = cells
        .filter(c => c.range.start || c.range.end)
        .map(c => {
          const s = c.range.start ?? c.range.end!
          const e = c.range.end ?? c.range.start!
          // PaP siempre es una fecha única: se muestra como punto, nunca como rango.
          if (c.key === 'pap') {
            return { key: c.key, label: PHASE_LABELS[c.key], start: s, end: s, isRange: false }
          }
          const [start, end] = s <= e ? [s, e] : [e, s]
          return { key: c.key, label: PHASE_LABELS[c.key], start, end, isRange: start.getTime() !== end.getTime() }
        })

      const pap = cells[3].range
      const papDate = pap.start ?? pap.end ?? null

      // Rango total del proyecto: primera y última fecha entre todas las fases.
      const allDates = phases.flatMap(ph => [ph.start, ph.end]).sort((a, b) => a.getTime() - b.getTime())
      const start = allDates[0] ?? null
      const end = allDates[allDates.length - 1] ?? null

      const statusRaw = text(row, cols.status)
      const status = statusFrom(statusRaw)

      return {
        key: `row-${index}-${text(row, cols.id, '')}`,
        id: text(row, cols.id, `PROY-${index + 1}`),
        name: text(row, cols.name, `Proyecto ${index + 1}`),
        application: text(row, cols.application, '—'),
        status,
        statusLabel: statusRaw || STATUS_LABELS[status],
        stage: statusRaw || STATUS_LABELS[status],
        responsible: text(row, cols.responsible, '—'),
        devTeam: optionalText(row, cols.devTeam),
        components: text(row, cols.components, '—'),
        direction: text(row, cols.direction, '—'),
        start,
        end,
        pap: papDate,
        phases
      }
    })
}

// Parsea un archivo Excel (como ArrayBuffer) y devuelve los proyectos.
// Reutilizado tanto por la carga automática desde el repo como por la carga manual.
function projectsFromBuffer(data: ArrayBuffer): Project[] {
  const workbook = XLSX.read(data, { type: 'array', cellDates: true })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<RawRow>(sheet, { defval: '' })
  const parsed = excelToProjects(rows)
  if (!parsed.length) throw new Error('No se encontraron filas de datos.')
  return parsed
}

function sampleProjects(): Project[] {
  const d = (s: string) => parseDate(s)!
  // Una etapa puede ser una fecha "2026-07-01" (punto) o un rango "inicio|fin".
  const phase = (key: PhaseKey, spec: string | null): Phase | null => {
    if (!spec) return null
    const [a, b] = spec.split('|')
    // PaP siempre es una fecha única (punto), aunque se pase un rango.
    const s = d(a), e = key === 'pap' ? s : (b ? d(b) : s)
    return { key, label: PHASE_LABELS[key], start: s, end: e, isRange: s.getTime() !== e.getTime() }
  }
  const mk = (
    id: string, name: string, application: string, status: Status, responsible: string, devTeam: string, direction: string,
    refine: string | null, dev: string | null, test: string | null, pap: string | null
  ): Project => {
    const phases = [
      phase('refinamiento', refine),
      phase('desarrollo', dev),
      phase('testing', test),
      phase('pap', pap)
    ].filter(Boolean) as Phase[]
    const dates = phases.flatMap(ph => [ph.start, ph.end]).sort((a, b) => a.getTime() - b.getTime())
    return {
      key: id, id, name, application, status, statusLabel: STATUS_LABELS[status],
      stage: STATUS_LABELS[status], responsible, devTeam, components: '—', direction,
      start: dates[0] ?? null, end: dates[dates.length - 1] ?? null,
      pap: pap ? d(pap.split('|')[0]) : null, phases
    }
  }
  return [
    mk('PROI-392', 'Subsanación de huelas Fase 3 - Regeneración de contratos', 'Web', 'testing', 'Diego H.', 'Roberto V.', 'Regulatorio', null, '2026-06-02|2026-06-30', '2026-07-03|2026-07-20', '2026-10-15'),
    mk('PROI-424', 'Dito - Proceso de subsanación Inconsistencia de Materiales', 'Web', 'testing', 'Diego H.', '', 'Regulatorio', null, '2026-07-06|2026-08-10', '2026-08-21|2026-09-05', '2026-10-15'),
    mk('PROI-703', 'Transformación del Modelo de Apellidos del Contacto', 'App', 'desarrollo', 'Juan Pablo', 'Roberto V / Bryan', 'Regulatorio', '2026-06-20|2026-06-30', '2026-07-01|2026-07-14', '2026-07-15|2026-08-05', '2026-10-08'),
    mk('PROI-511', 'Segmentación B2C única para herramientas - Parte2', 'App', 'testing', 'Miguel Pacheco', '', 'Negocio', null, '2026-07-01|2026-08-31', '2026-09-15|2026-09-21', '2026-09-24'),
    mk('PROI-723', 'Venta de eSIM Retail Dito', 'Web', 'testing', 'Joel T.', 'Roberto V.', 'Negocio', null, '2026-07-01|2026-08-31', '2026-09-15|2026-09-21', '2026-09-24'),
    mk('PROI-800', 'Proyecto de definición pendiente', 'Backend', 'definicion', 'Por asignar', '', 'Negocio', null, null, null, '2026-09-24'),
    mk('PROI-801', 'Iniciativa en refinamiento', 'Backend', 'refinamiento', 'Por asignar', '', 'Tecnologia', '2026-10-09|2026-10-16', null, null, null)
  ]
}

// Posición horizontal 0..1 dentro del timeline (Junio..Diciembre) para una fecha.
function timelinePosition(date: Date): number {
  const monthsFromStart = (date.getFullYear() - CURRENT_YEAR) * 12 + (date.getMonth() - FIRST_MONTH)
  const dayFraction = (date.getDate() - 1) / 31
  const pos = (monthsFromStart + dayFraction) / MONTH_COUNT
  return Math.max(0, Math.min(1, pos))
}

type SortMode = 'none' | 'pap-asc' | 'pap-desc'
const ALL_APPS = '__all__'
// Excel versionado en el repo (carpeta public/). Para actualizar los datos que
// ve todo el mundo, reemplaza este archivo y vuelve a desplegar.
const DEFAULT_FILE = 'matriz_backlog.xlsx'

function App() {
  const [projects, setProjects] = useState<Project[]>(sampleProjects())
  const [fileName, setFileName] = useState('Ejemplo')
  const [error, setError] = useState('')
  const [appFilter, setAppFilter] = useState<string>(ALL_APPS)
  const [sortMode, setSortMode] = useState<SortMode>('none')
  const [monthWidth, setMonthWidth] = useState(200) // px por mes (zoom del timeline)
  const inputRef = useRef<HTMLInputElement>(null)

  const MIN_ZOOM = 120, MAX_ZOOM = 420, ZOOM_STEP = 60
  const zoomIn = () => setMonthWidth(w => Math.min(MAX_ZOOM, w + ZOOM_STEP))
  const zoomOut = () => setMonthWidth(w => Math.max(MIN_ZOOM, w - ZOOM_STEP))

  // Aplicaciones únicas presentes en los datos, para poblar el selector.
  const applications = useMemo(
    () => Array.from(new Set(projects.map(p => p.application).filter(a => a && a !== '—'))).sort(),
    [projects]
  )

  // Proyectos visibles tras aplicar filtro por aplicación y orden por PaP.
  const visibleProjects = useMemo(() => {
    let list = appFilter === ALL_APPS ? projects : projects.filter(p => p.application === appFilter)
    if (sortMode !== 'none') {
      const dir = sortMode === 'pap-asc' ? 1 : -1
      list = [...list].sort((a, b) => {
        // Los proyectos sin fecha de PaP van siempre al final.
        if (!a.pap && !b.pap) return 0
        if (!a.pap) return 1
        if (!b.pap) return -1
        return (a.pap.getTime() - b.pap.getTime()) * dir
      })
    }
    return list
  }, [projects, appFilter, sortMode])

  const isDirty = appFilter !== ALL_APPS || sortMode !== 'none' || monthWidth !== 200
  function resetView() {
    setAppFilter(ALL_APPS)
    setSortMode('none')
    setMonthWidth(200)
  }

  const stats = useMemo(() => ({
    total: visibleProjects.length,
    produccion: visibleProjects.filter(p => p.status === 'produccion').length,
    testing: visibleProjects.filter(p => p.status === 'testing').length,
    desarrollo: visibleProjects.filter(p => p.status === 'desarrollo').length,
    planificado: visibleProjects.filter(p => p.status === 'planificado' || p.status === 'refinamiento' || p.status === 'definicion').length,
    cancelado: visibleProjects.filter(p => p.status === 'cancelado').length
  }), [visibleProjects])

  async function handleFile(file: File) {
    setError('')
    try {
      const parsed = projectsFromBuffer(await file.arrayBuffer())
      setProjects(parsed)
      setFileName(file.name)
      resetView()
    } catch (e) {
      setError(e instanceof Error ? e.message : 'No se pudo procesar el Excel.')
    }
  }

  // Al iniciar, carga automáticamente el Excel incluido en el repositorio
  // (public/matriz.xlsx) para que todos vean los mismos datos sin subir nada.
  useEffect(() => {
    let cancelled = false
    async function loadDefault() {
      try {
        const url = `${import.meta.env.BASE_URL}${DEFAULT_FILE}`
        const res = await fetch(url)
        if (!res.ok) throw new Error(`No se pudo leer ${DEFAULT_FILE} (HTTP ${res.status}).`)
        const parsed = projectsFromBuffer(await res.arrayBuffer())
        if (!cancelled) {
          setProjects(parsed)
          setFileName(DEFAULT_FILE)
        }
      } catch {
        // Si no existe el archivo del repo, se mantienen los datos de ejemplo.
        if (!cancelled) setFileName('Ejemplo')
      }
    }
    loadDefault()
    return () => { cancelled = true }
  }, [])

  const statusIcon: Record<Status, string> = {
    produccion: '✓', testing: '⚑', desarrollo: '▣', planificado: '◷',
    refinamiento: '◔', definicion: '○', cancelado: '×', 'sin-estado': '—'
  }

  return (
    <div className="app">
      <header className="topbar">
        <div>
          <h1>ROADMAP PROYECTOS Y REQUERIMIENTOS</h1>
          <p>Generador automático desde Excel</p>
        </div>
        <button className="upload" onClick={() => inputRef.current?.click()}>Cargar Excel</button>
        <input ref={inputRef} hidden type="file" accept=".xlsx,.xls,.csv" onChange={e => e.target.files?.[0] && handleFile(e.target.files[0])} />
      </header>

      <section className="toolbar">
        <div className="file-pill">Datos: <strong>{fileName}</strong></div>

        <label className="control">
          <span>Aplicación</span>
          <select value={appFilter} onChange={e => setAppFilter(e.target.value)}>
            <option value={ALL_APPS}>Todas</option>
            {applications.map(a => <option key={a} value={a}>{a}</option>)}
          </select>
        </label>

        <label className="control">
          <span>Ordenar por PaP</span>
          <select value={sortMode} onChange={e => setSortMode(e.target.value as SortMode)}>
            <option value="none">Sin ordenar</option>
            <option value="pap-asc">Más próximo primero</option>
            <option value="pap-desc">Más lejano primero</option>
          </select>
        </label>

        <div className="control zoom">
          <span>Zoom del timeline</span>
          <div className="zoom-buttons">
            <button onClick={zoomOut} disabled={monthWidth <= MIN_ZOOM} aria-label="Alejar">−</button>
            <button onClick={zoomIn} disabled={monthWidth >= MAX_ZOOM} aria-label="Acercar">+</button>
          </div>
        </div>

        <button className="reset" onClick={resetView} disabled={!isDirty}>Volver al inicio</button>

        <div className="count">{visibleProjects.length} de {projects.length}</div>
      </section>

      {error && <div className="error">{error}</div>}

      <section className="stats">
        <Stat value={stats.total} label="Proyectos totales" icon="▥" />
        <Stat value={stats.produccion} label="Producción" icon="✓" />
        <Stat value={stats.testing} label="Testing" icon="⚑" />
        <Stat value={stats.desarrollo} label="Desarrollo" icon="▣" />
        <Stat value={stats.planificado} label="Planificado / Backlog" icon="◷" />
        <Stat value={stats.cancelado} label="Cancelado" icon="×" />
      </section>

      <main className="roadmap-card" style={{ '--month-w': `${monthWidth}px` } as React.CSSProperties}>
        <div className="roadmap-scroll">
          <div className="roadmap-inner">
            <div className="grid header-row">
              <div>Proyecto</div><div>Aplicación</div><div className="timeline-head"><div className="year">2026</div><div className="months">{MONTHS.map(m => <span key={m}>{m}</span>)}</div></div>
              <div>Estado</div><div>Responsable</div><div>Componentes</div><div>Dirección</div>
            </div>

            {visibleProjects.length
              ? visibleProjects.map(p => <RoadmapRow key={p.key} p={p} statusIcon={statusIcon[p.status]} />)
              : <div className="empty">No hay proyectos para esta aplicación.</div>}

            <div className="legend">
              <strong>ETAPAS</strong>
              <span className="legend-item"><i className="seg phase-refinamiento" /> Refinamiento</span>
              <span className="legend-item"><i className="seg phase-desarrollo" /> Desarrollo</span>
              <span className="legend-item"><i className="seg phase-testing" /> Testing</span>
              <span className="legend-item"><i className="seg phase-pap" /> PaP</span>
              <span className="legend-sep" />
              <strong>ESTADO</strong>
              <span className="legend-item"><i className="dot produccion">✓</i> Producción</span>
              <span className="legend-item"><i className="dot testing">⚑</i> Testing</span>
              <span className="legend-item"><i className="dot desarrollo">▣</i> Desarrollo</span>
              <span className="legend-item"><i className="dot planificado">◷</i> Planificado</span>
              <span className="legend-item"><i className="dot cancelado">×</i> Cancelado</span>
            </div>
          </div>
        </div>
      </main>
    </div>
  )
}

function Stat({ value, label, icon }: { value:number, label:string, icon:string }) {
  return <div className="stat"><i>{icon}</i><div><b>{value}</b><span>{label}</span></div></div>
}

// Abreviatura corta de etapa para la etiqueta del segmento.
const PHASE_SHORT: Record<PhaseKey, string> = {
  refinamiento: 'Refin.', desarrollo: 'Desarrollo', testing: 'Testing', pap: 'PaP'
}

function RoadmapRow({ p, statusIcon }: { p: Project, statusIcon:string }) {
  const left = p.start ? timelinePosition(p.start) * 100 : 0
  const right = p.end ? timelinePosition(p.end) * 100 : 0
  const width = p.start && p.end ? Math.max(0.6, right - left) : 0

  // Cada fase se ubica en el timeline y se reparte en dos carriles (arriba/abajo)
  // alternando por orden, para que las etiquetas no se amontonen.
  const placed = p.phases
    .map((ph, i) => {
      const x1 = timelinePosition(ph.start) * 100
      const x2 = timelinePosition(ph.end) * 100
      return { ...ph, i, x1, x2, lane: i % 2 }
    })

  return <div className="grid data-row">
    <div className="project-cell">
      <b>{p.id}</b>
      <span>{p.name}</span>
      {p.devTeam && <small className="dev-team">Dev: {p.devTeam}</small>}
    </div>
    <div className="app-cell">{p.application}</div>
    <div className="timeline">
      {p.start && p.end
        ? <div className="bar-base" style={{left:`${left}%`, width:`${width}%`}} />
        : <div className="dash" />}
      {placed.map(ph => {
        // Rango: segmento con ancho real. Punto (p.ej. PaP): ancho 0 -> un solo círculo.
        const w = ph.isRange ? Math.max(0.6, ph.x2 - ph.x1) : 0
        return (
          <div className={`phase phase-${ph.key} lane-${ph.lane} ${ph.isRange ? 'is-range' : 'is-point'}`} key={ph.i} style={{left:`${ph.x1}%`, width:`${w}%`}}>
            <em className="phase-start">{formatDate(ph.start)}</em>
            <span className="phase-seg" />
            {ph.isRange && <em className="phase-end">{formatDate(ph.end)}</em>}
            <small className="phase-label">{PHASE_SHORT[ph.key]}</small>
          </div>
        )
      })}
    </div>
    <div className={`status ${p.status}`} title={p.statusLabel}><i>{statusIcon}</i><span>{p.statusLabel}</span></div>
    <div className="responsible">{p.responsible}</div>
    <div className="components">{p.components}</div>
    <div className="direction">{p.direction}</div>
  </div>
}

createRoot(document.getElementById('root')!).render(<App />)
