const state = {
  sources: {},
  rows: [],
  filteredRows: [],
  xlsxPromise: null,
  profile: 'supervisor',
  processStarted: true,
  supervisorSubmitted: false,
};

const PROFILE_META = {
  supervisor: { label: 'Supervisor', initials: 'S', banner: 'Vista Supervisor · Juan Alzualde', detail: 'Planilla del personal asignado al centro de costo del piloto.', supervisor: 'Juan Alzualde', phase: 1 },
  rrhh: { label: 'RRHH', initials: 'R', banner: 'Vista RRHH', detail: 'Consulta todos los reportes, fuentes e historial del período.' },
  gerente: { label: 'Gerente', initials: 'G', banner: 'Vista Gerente · Pedro Albornoz', detail: 'Aprueba, edita u observa lo reportado por supervisión.', supervisor: 'Juan Alzualde', phase: 1 },
  admin: { label: 'Admin', initials: 'A', banner: 'Vista Admin · artBPO', detail: 'Carga fuentes, inicia y procesa el período mensual.' },
  superadmin: { label: 'SuperAdmin', initials: 'SA', banner: 'Vista SuperAdmin', detail: 'Acceso total: mantenedores, operación y suplantación de perfiles.' },
};

const SOURCE_LABELS = {
  payroll: 'Libro provisional REX+',
  master: 'Maestro de personal REX+',
  norte: 'Centralizado Norte',
  metropolitana: 'Centralizado Metropolitana',
};

const SOURCE_IDS = ['payroll', 'master', 'norte', 'metropolitana'];
const CENTRAL_SOURCE_IDS = ['norte', 'metropolitana'];
const PILOT_CENTERS = new Set(['TECNICO MULTISKILL', 'TERRENO']);
const PHASE_TWO_CENTERS = new Set(['RECUPERA', 'MONTAJE']);

const CENTRAL_COLUMNS = {
  tag: ['Suma de TAG', 'TAG'],
  multa: ['Suma de Multa', 'Multa'],
  fuel: ['Suma de Combustible - Consumo', 'Combustible - Consumo', 'Consumo combustible'],
  assignment: ['Suma de Asignación Combustible', 'Asignación Combustible', 'Asignacion Combustible'],
  total: ['Suma de TOTAL CONSUMO', 'TOTAL CONSUMO', 'Total Consumo'],
};

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

function normalize(value) {
  return String(value ?? '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');
}

function normalizeRut(value) {
  return String(value ?? '').replace(/[.\s]/g, '').toUpperCase();
}

function formatNumber(value) {
  return new Intl.NumberFormat('es-CL').format(Number(value) || 0);
}

function formatCurrency(value) {
  return new Intl.NumberFormat('es-CL', { style: 'currency', currency: 'CLP', maximumFractionDigits: 0 }).format(Math.round(Number(value) || 0));
}

function parseAmount(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  let text = String(value ?? '').trim().replace(/[$\s%]/g, '');
  if (!text) return 0;
  const comma = text.includes(',');
  const dot = text.includes('.');
  if (comma && dot) text = text.lastIndexOf(',') > text.lastIndexOf('.') ? text.replace(/\./g, '').replace(',', '.') : text.replace(/,/g, '');
  else if (comma) text = text.split(',').slice(1).every((part) => part.length === 3) ? text.replace(/,/g, '') : text.replace(',', '.');
  else if (dot && text.split('.').slice(1).every((part) => part.length === 3)) text = text.replace(/\./g, '');
  const amount = Number(text);
  return Number.isFinite(amount) ? amount : 0;
}

function round(value) {
  return Math.round((Number(value || 0) + Number.EPSILON) * 100) / 100;
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' })[character]);
}

function pick(record, names) {
  const values = new Map(Object.entries(record).map(([key, value]) => [normalize(key), value]));
  for (const name of names) {
    const value = values.get(normalize(name));
    if (value !== undefined && value !== null && String(value).trim() !== '') return value;
  }
  return '';
}

function fileSize(bytes) {
  if (!bytes) return '';
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function loadXlsx() {
  if (window.XLSX) return Promise.resolve();
  if (state.xlsxPromise) return state.xlsxPromise;
  state.xlsxPromise = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'vendor/xlsx.full.min.js';
    script.onload = resolve;
    script.onerror = () => reject(new Error('No se pudo cargar el lector de Excel.'));
    document.head.appendChild(script);
  });
  return state.xlsxPromise;
}

function findHeaderIndex(matrix, sourceId) {
  if (sourceId === 'master') {
    const masterIndex = matrix.findIndex((row) => {
      const headers = row.map(normalize);
      return headers.includes('EMPLEADO') && headers.includes('NOMBRE COMPLETO') && headers.includes('NOMBRE CENTRO COSTO');
    });
    if (masterIndex >= 0) return masterIndex;
  }
  const index = matrix.findIndex((row) => {
    const headers = row.map(normalize);
    const rut = headers.includes('RUT') || headers.includes('RUT TAC');
    const name = headers.includes('NOMBRE') || headers.includes('NOMBRE TAC');
    return rut && name;
  });
  if (index >= 0) return index;
  return matrix.findIndex((row) => row.some((cell) => String(cell).trim() !== ''));
}

function recordsFromMatrix(matrix, sourceId) {
  const headerIndex = findHeaderIndex(matrix, sourceId);
  if (headerIndex < 0) return { header: [], records: [] };
  const rawHeader = matrix[headerIndex].map((cell) => String(cell ?? '').trim());
  const header = rawHeader.filter(Boolean);
  const records = matrix.slice(headerIndex + 1)
    .filter((row) => row.some((cell) => String(cell ?? '').trim() !== ''))
    .map((row) => {
      const seen = new Map();
      return rawHeader.reduce((record, key, index) => {
        if (!key) return record;
        const count = (seen.get(key) || 0) + 1;
        seen.set(key, count);
        record[count === 1 ? key : `${key}_${count}`] = row[index] ?? '';
        return record;
      }, {});
    });
  return { header, records };
}

async function inspectFile(file, sourceId) {
  await loadXlsx();
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true, cellText: true, raw: false });
  const sheets = workbook.SheetNames.map((name) => {
    const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: false });
    const parsed = recordsFromMatrix(matrix, sourceId);
    return { name, ...parsed };
  });
  const warnings = [];
  if (sourceId === 'payroll' && !sheets.some((sheet) => normalize(sheet.name) === 'DETALLE')) warnings.push('No se encontró la hoja Detalle; se revisará la primera hoja con datos.');
  if (sourceId === 'master' && !sheets.some((sheet) => sheet.header.some((cell) => normalize(cell) === 'EMPLEADO'))) warnings.push('No se encontró la estructura de maestro REX+ esperada.');
  if (CENTRAL_SOURCE_IDS.includes(sourceId) && !sheets.some((sheet) => normalize(sheet.name) === 'TECNICOS')) warnings.push('No se encontró la hoja TECNICOS; verifica el template regional.');
  const rows = sheets.reduce((total, sheet) => total + sheet.records.length, 0);
  const duplicateRuts = findDuplicateRuts(sheets.flatMap((sheet) => sheet.records));
  if (duplicateRuts.length) warnings.push(`${duplicateRuts.length} RUT repetido${duplicateRuts.length === 1 ? '' : 's'}; se consolidará por RUT.`);
  return { sheets, rows, warnings, level: warnings.length ? 'warning' : 'valid', filename: file.name };
}

function findDuplicateRuts(records) {
  const seen = new Set();
  const duplicates = new Set();
  records.forEach((record) => {
    const rut = normalizeRut(pick(record, ['RUT', 'RUT TAC', 'empleado', 'Empleado']));
    if (!rut) return;
    if (seen.has(rut)) duplicates.add(rut);
    seen.add(rut);
  });
  return [...duplicates];
}

function setSourceStatus(sourceId, level, label, detail) {
  const status = $(`[data-source-status="${sourceId}"]`);
  if (status) {
    status.className = `status-pill ${level}`;
    status.textContent = label;
  }
  const detailNode = $(`[data-source-detail="${sourceId}"]`);
  if (detailNode) detailNode.textContent = detail;
  const fileNode = $(`[data-file-name="${sourceId}"]`);
  if (fileNode && state.sources[sourceId]) fileNode.textContent = `${state.sources[sourceId].file.name} · ${fileSize(state.sources[sourceId].file.size)}`;
}

async function handleFile(sourceId, file) {
  state.processStarted = false;
  state.rows = [];
  setSourceStatus(sourceId, 'neutral', 'Revisando', 'Validando estructura…');
  const fileNode = $(`[data-file-name="${sourceId}"]`);
  if (fileNode) fileNode.textContent = `${file.name} · ${fileSize(file.size)}`;
  try {
    const report = await inspectFile(file, sourceId);
    state.sources[sourceId] = { file, report };
    const level = report.level === 'warning' ? 'warning' : 'valid';
    setSourceStatus(sourceId, level, report.level === 'warning' ? 'Revisar' : 'Estructura OK', `${formatNumber(report.rows)} registros · ${report.warnings.length ? report.warnings.join(' ') : 'Lectura correcta'}`);
  } catch (error) {
    state.sources[sourceId] = { file, report: { rows: 0, sheets: [], warnings: [error.message], level: 'error' } };
    setSourceStatus(sourceId, 'error', 'No leído', error.message);
  }
  renderSourceSummary();
  updateCalculateState();
  applyProfileView();
}

function masterApiUrl() {
  const query = '?empresa=NOW';
  return location.protocol === 'file:' ? `http://127.0.0.1:8063/api/rex/master${query}` : `/api/rex/master${query}`;
}

async function loadMasterFromApi() {
  const button = $('#load-master-api');
  if (button) button.disabled = true;
  setSourceStatus('master', 'neutral', 'Consultando', 'Conectando con REX+…');
  try {
    const response = await fetch(masterApiUrl(), { headers: { Accept: 'application/json' } });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.error || `REX+ respondió ${response.status}`);
    const records = Array.isArray(payload.rows) ? payload.rows : [];
    if (!records.length) throw new Error('REX+ respondió sin registros para la empresa NOW.');
    const file = { name: 'Mestro NOW · API REX+', size: 0 };
    const report = {
      sheets: [{ name: 'API REX+', header: Object.keys(records[0]), records }],
      rows: records.length,
      warnings: [],
      level: 'valid',
      filename: file.name,
    };
    state.sources.master = { file, report };
    state.rows = buildRows();
    $('#kpi-workers').textContent = formatNumber(state.rows.length);
    setSourceStatus('master', 'valid', 'API conectada', `${formatNumber(records.length)} registros recibidos desde REX+`);
    renderSourceSummary();
    updateCalculateState();
    applyProfileView();
    return true;
  } catch (error) {
    setSourceStatus('master', 'warning', 'API pendiente', error.message);
    const detail = $('[data-source-detail="master"]');
    if (detail) detail.textContent = `${error.message} Puedes cargar el Excel sólo como respaldo del piloto.`;
    return false;
  } finally {
    if (button) button.disabled = !['rrhh', 'admin', 'superadmin'].includes(state.profile);
  }
}

function pilotSourceUrl(sourceId) {
  return location.protocol === 'file:' ? `http://127.0.0.1:8063/api/pilot/source/${sourceId}` : `/api/pilot/source/${sourceId}`;
}

async function fetchPilotFile(sourceId) {
  const response = await fetch(pilotSourceUrl(sourceId), { headers: { Accept: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' } });
  if (!response.ok) throw new Error(`No se pudo cargar la fuente piloto ${sourceId}.`);
  const blob = await response.blob();
  const names = { payroll: 'LIBRO NOW AGOSTO 26.xlsx', master: 'Mestro NOW (1).xlsx', norte: 'NORTE-AGOSTO-2026.xlsx', metropolitana: 'METROPOLITANA-AGOSTO-2026.xlsx' };
  return new File([blob], names[sourceId], { type: blob.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

async function loadAugustPilotSources() {
  if ($('#period').value !== '2026-08') return;
  try {
    const apiLoaded = await loadMasterFromApi();
    if (!apiLoaded) await handleFile('master', await fetchPilotFile('master'));
    for (const sourceId of ['payroll', 'norte', 'metropolitana']) {
      await handleFile(sourceId, await fetchPilotFile(sourceId));
    }
    if (SOURCE_IDS.every(sourceReady)) {
      calculate();
      // The pilot loader completes the same transition as the manual button.
      state.processStarted = true;
      applyProfileView();
    }
  } catch {
    // The local pilot loader is optional; regular uploads remain available.
  }
}

function wireFileInput(sourceId) {
  const input = $(`[data-file-input="${sourceId}"]`);
  if (!input) return;
  input.addEventListener('change', () => { if (input.files[0]) handleFile(sourceId, input.files[0]); });
  const dropzone = $(`[data-drop-zone="${sourceId}"]`);
  if (!dropzone) return;
  ['dragenter', 'dragover'].forEach((eventName) => dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.add('dragging'); }));
  ['dragleave', 'drop'].forEach((eventName) => dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.remove('dragging'); }));
  dropzone.addEventListener('drop', (event) => { if (event.dataTransfer.files[0]) handleFile(sourceId, event.dataTransfer.files[0]); });
}

function wireUploadButtons() {
  $$('[data-browse]').forEach((button) => button.addEventListener('click', () => $(`[data-file-input="${button.dataset.browse}"]`).click()));
  [...SOURCE_IDS].forEach(wireFileInput);
  const supervisorInput = $('#file-supervisor-variables');
  if (supervisorInput) supervisorInput.addEventListener('change', () => { if (supervisorInput.files[0]) handleSupervisorFile(supervisorInput.files[0]); });
}

function sourceLoaded(sourceId) {
  return Boolean(state.sources[sourceId]?.report?.rows >= 0);
}

function sourceReady(sourceId) {
  return Boolean(state.sources[sourceId]?.report && state.sources[sourceId].report.level !== 'error');
}

function renderSourceSummary() {
  const loadedIds = SOURCE_IDS.filter(sourceLoaded);
  const centralLoaded = CENTRAL_SOURCE_IDS.filter(sourceLoaded);
  const rows = loadedIds.reduce((total, id) => total + Number(state.sources[id].report.rows || 0), 0);
  const alerts = loadedIds.reduce((total, id) => total + state.sources[id].report.warnings.length, 0);
  $('#validation-files').textContent = `${loadedIds.length} / ${SOURCE_IDS.length}`;
  $('#validation-rows').textContent = formatNumber(rows);
  $('#validation-alerts').textContent = formatNumber(alerts);
  $('#kpi-sources').textContent = `${loadedIds.length} / ${SOURCE_IDS.length}`;
  $('#validation-summary').textContent = loadedIds.length === 0 ? 'Esperando las fuentes del período' : `${loadedIds.length} fuente${loadedIds.length === 1 ? '' : 's'} recibida${loadedIds.length === 1 ? '' : 's'} · ${alerts} alerta${alerts === 1 ? '' : 's'}`;
  const centralStatus = $('#central-status');
  const centralDetail = $('#central-detail');
  if (centralLoaded.length === 0) {
    centralStatus.className = 'status-pill neutral';
    centralStatus.textContent = 'Pendiente';
    centralDetail.textContent = 'Carga ambos archivos regionales';
  } else {
    const centralAlerts = centralLoaded.reduce((total, id) => total + state.sources[id].report.warnings.length, 0);
    centralStatus.className = `status-pill ${centralAlerts ? 'warning' : 'valid'}`;
    centralStatus.textContent = centralAlerts ? 'Revisar' : 'Fuentes OK';
    centralDetail.textContent = `${centralLoaded.length} región${centralLoaded.length === 1 ? '' : 'es'} cargada${centralLoaded.length === 1 ? '' : 's'} · ${formatNumber(centralLoaded.reduce((total, id) => total + state.sources[id].report.rows, 0))} registros`;
  }
  const alertList = $('#alert-list');
  const alertItems = loadedIds.flatMap((id) => state.sources[id].report.warnings.map((warning) => `<li><strong>${SOURCE_LABELS[id]}:</strong> ${escapeHtml(warning)}</li>`));
  alertList.hidden = alertItems.length === 0;
  alertList.innerHTML = alertItems.length ? `<strong>Alertas para revisar</strong><ul>${alertItems.join('')}</ul>` : '';
  $('#kpi-workers').textContent = state.rows.length ? formatNumber(state.rows.length) : '0';
}

function updateCalculateState() {
  const ready = SOURCE_IDS.every(sourceReady);
  $('#calculate-button').disabled = !ready;
}

function findSheet(report, expected) {
  return report?.sheets.find((sheet) => normalize(sheet.name) === normalize(expected)) || report?.sheets.find((sheet) => sheet.records.length) || null;
}

function isNowCompany(record) {
  const namedCompany = String(pick(record, ['Nombre empresa', 'Compañía', 'Sociedad'])).trim();
  if (namedCompany) return normalize(namedCompany).includes('NOW');
  const companyCode = String(pick(record, ['Empresa'])).trim();
  return !companyCode || /^\d+$/.test(companyCode) || normalize(companyCode).includes('NOW');
}

function assignmentForCenter(value) {
  const centroCosto = normalize(value);
  if (PILOT_CENTERS.has(centroCosto)) return { supervisor: 'Juan Alzualde', phase: 1, phaseLabel: 'Piloto actual' };
  if (PHASE_TWO_CENTERS.has(centroCosto)) return { supervisor: 'Mauricio Sanhueza', phase: 2, phaseLabel: 'Fase 2' };
  return { supervisor: 'Sin asignar', phase: 2, phaseLabel: 'Fase 2' };
}

function buildMasterRows() {
  const sheet = findSheet(state.sources.master?.report, 'Hoja 1');
  return (sheet?.records || []).map((record) => {
    const rut = String(pick(record, ['empleado', 'RUT', 'Rut', 'RUT TAC'])).trim();
    if (!rut || !isNowCompany(record)) return null;
    const centroCosto = String(pick(record, ['nombre_centro_costo', 'Centro de costo', 'Centro Costo', 'CC'])).trim();
    return {
      rut,
      key: normalizeRut(rut),
      nombre: String(pick(record, ['nombre_completo', 'Nombre', 'NOMBRE'])).trim(),
      empresa: String(pick(record, ['nombre_empresa', 'Empresa', 'Nombre empresa'])).trim(),
      estado: String(pick(record, ['estado', 'Estado'])).trim(),
      contrato: String(pick(record, ['contratoActi', 'Contrato', 'Tipo Contrato'])).trim() || '1',
      cargo: String(pick(record, ['nombre_cargo', 'Cargo'])).trim(),
      sede: String(pick(record, ['nombre_sede', 'Sede'])).trim(),
      fechaInicio: String(pick(record, ['fechaInic', 'Fecha inicio', 'Fecha Inicio', 'fecha_inicio'])).trim(),
      centroCosto,
      ...assignmentForCenter(centroCosto),
    };
  }).filter(Boolean);
}

function buildPayrollRows() {
  const sheet = findSheet(state.sources.payroll?.report, 'Detalle');
  return (sheet?.records || []).map((record) => {
    const rut = String(pick(record, ['RUT', 'Rut', 'RUT TAC'])).trim();
    if (!rut || !isNowCompany(record)) return null;
    return {
      rut,
      key: normalizeRut(rut),
      nombre: String(pick(record, ['Nombre', 'NOMBRE', 'Nombre TAC'])).trim(),
      empresa: String(pick(record, ['Nombre empresa', 'Empresa'])).trim(),
      contrato: String(pick(record, ['Contrato', 'Tipo Contrato'])).trim() || '1',
      centroCosto: String(pick(record, ['Centro de costo', 'Centro Costo', 'CC'])).trim(),
      supervisor: String(pick(record, ['Supervisor', 'NOMBRE SUPERVISOR'])).trim(),
      zona: String(pick(record, ['Zona', 'Región', 'Region', 'Sede'])).trim(),
      fechaInicio: String(pick(record, ['Fecha inicio', 'Fecha Inicio', 'fechaInic', 'fecha_inicio'])).trim(),
      sueldoBase: parseAmount(pick(record, ['Sueldo base', 'Sueldo Base', 'VALOR SUELDO BASE'])),
      diasTrabajados: parseAmount(pick(record, ['Días trabajados', 'Dias trabajados', 'Días Trabajados'])),
      hheeHoras: parseAmount(pick(record, ['Cantidad HHEE', 'Q HHEE', 'Horas Extra', 'HHEE', 'Horas extraordinarias'])),
      hheeMontoInformado: parseAmount(pick(record, ['Monto HHEE', 'Total HHEE', 'Horas Extra Monto'])),
      desgasteContractual: parseAmount(pick(record, ['Desgaste contractual', 'Desgaste de herramientas', 'Desgaste Herramientas'])),
    };
  }).filter(Boolean);
}

function centralRows() {
  const grouped = new Map();
  ['norte', 'metropolitana'].forEach((sourceId) => {
    const report = state.sources[sourceId]?.report;
    const technical = findSheet(report, 'TECNICOS');
    (technical?.records || []).forEach((record) => {
      const rut = String(pick(record, ['RUT TAC', 'RUT'])).trim();
      const key = normalizeRut(rut);
      if (!key) return;
      const current = grouped.get(key) || { rut, region: sourceId === 'norte' ? 'Norte' : 'Metropolitana', tag: 0, multa: 0, fuel: 0, assignment: 0, total: 0 };
      current.tag += parseAmount(pick(record, CENTRAL_COLUMNS.tag));
      current.multa += parseAmount(pick(record, CENTRAL_COLUMNS.multa));
      current.fuel += parseAmount(pick(record, CENTRAL_COLUMNS.fuel));
      current.assignment += parseAmount(pick(record, CENTRAL_COLUMNS.assignment));
      current.total += parseAmount(pick(record, CENTRAL_COLUMNS.total));
      grouped.set(key, current);
    });
  });
  return grouped;
}

function buildRows() {
  const central = centralRows();
  const payroll = buildPayrollRows();
  const masterRows = buildMasterRows();
  const master = new Map(masterRows.map((person) => [person.key, person]));
  const payrollKeys = new Set(payroll.map((person) => person.key));
  const rowsByRut = new Map();
  payroll.forEach((person) => {
    if (!rowsByRut.has(person.key)) rowsByRut.set(person.key, person);
  });
  masterRows.forEach((person) => {
    if (!rowsByRut.has(person.key)) rowsByRut.set(person.key, {
      ...person,
      sueldoBase: null,
      diasTrabajados: null,
      hheeHoras: null,
      hheeMontoInformado: null,
      desgasteContractual: null,
      zona: person.sede || '',
    });
  });
  return [...rowsByRut.values()].map((person) => {
    const masterPerson = master.get(person.key);
    const assignment = masterPerson || assignmentForCenter(person.centroCosto);
    const supervisor = masterPerson?.supervisor || (assignment.supervisor !== 'Sin asignar' ? assignment.supervisor : person.supervisor || 'Sin asignar');
    const source = central.get(person.key) || { tag: 0, multa: 0, fuel: 0, assignment: 0, total: 0, region: person.zona || '' };
    const saldo = round(source.assignment - (source.tag + source.fuel));
    const movilizacion = saldo > 0 ? saldo : 0;
    const eficiencia = saldo < 0 ? saldo : 0;
    const hheeRate = person.sueldoBase > 0 ? person.sueldoBase / 180 * 1.5 : 0;
    const hheeMonto = person.hheeMontoInformado ?? (person.hheeHoras === null ? null : round(person.hheeHoras * hheeRate));
    return {
      ...person,
      nombre: masterPerson?.nombre || person.nombre,
      empresa: masterPerson?.empresa || person.empresa,
      contrato: masterPerson?.contrato || person.contrato,
      cargo: masterPerson?.cargo || person.cargo || '',
      sede: masterPerson?.sede || person.sede || '',
      fechaInicio: masterPerson?.fechaInicio || person.fechaInicio || '',
      centroCosto: masterPerson?.centroCosto || person.centroCosto,
      supervisor,
      payrollFound: payrollKeys.has(person.key),
      phase: assignment.phase,
      phaseLabel: assignment.phaseLabel,
      region: source.region || person.zona,
      centralFound: central.has(person.key),
      tag: round(source.tag),
      multa: round(source.multa),
      fuel: round(source.fuel),
      assignment: round(source.assignment),
      totalConsumption: round(source.total),
      saldo,
      movilizacion,
      eficiencia,
      hheeRate: round(hheeRate),
      hheeMonto: round(hheeMonto),
      desgaste: person.desgasteContractual === null ? null : round(person.desgasteContractual),
      desgasteAdicional: null,
      concurso: null,
      maestroGuia: null,
      compensacion: null,
      observaciones: '',
      approved: false,
      rejected: false,
      status: 'PENDIENTE',
    };
  }).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
}

function resultStatus(row) {
  const manualComplete = [row.concurso, row.desgasteAdicional, row.maestroGuia, row.compensacion].every((value) => value !== null && value !== undefined);
  return row.approved ? 'APROBADO' : (row.rejected ? 'OBSERVADO' : (manualComplete ? 'LISTO' : 'PENDIENTE'));
}

function visibleRows() {
  const metadata = PROFILE_META[state.profile];
  if (!metadata?.supervisor) return state.rows;
  return state.rows.filter((row) => row.supervisor === metadata.supervisor && row.phase === metadata.phase);
}

function conceptValue(value, pending = false) {
  return pending || value === null ? '<span class="pending-value">Pendiente</span>' : `<span class="money-value">${formatCurrency(value)}</span>`;
}

function desgasteValue(row) {
  const extra = row.desgasteAdicional === null ? '<span class="pending-value">Adicional pendiente</span>' : `<span class="manual-subvalue">+ ${formatCurrency(row.desgasteAdicional)}</span>`;
  return `<div class="amount-with-detail"><span class="money-value">${formatCurrency(row.desgaste)}</span><small>Base REX+ · ${extra}</small></div>`;
}

function renderResults() {
  const query = normalize($('#search-results').value);
  const supervisor = $('#supervisor-filter').value;
  const status = $('#result-status-filter').value;
  state.filteredRows = visibleRows().filter((row) => {
    const matchesQuery = !query || normalize(`${row.rut} ${row.nombre}`).includes(query);
    const matchesSupervisor = !supervisor || row.supervisor === supervisor;
    const matchesStatus = !status || resultStatus(row) === status;
    return matchesQuery && matchesSupervisor && matchesStatus;
  });
  const body = $('#results-body');
  if (!state.filteredRows.length) {
    body.innerHTML = `<tr><td colspan="11" class="empty-table"><span>⌕</span><strong>No hay trabajadores que coincidan</strong><small>Prueba con otro RUT, nombre, supervisor o estado.</small></td></tr>`;
  } else {
    body.innerHTML = state.filteredRows.map((row) => `<tr class="${row.approved ? 'row-approved' : ''}">
      <td><div class="worker-cell"><strong>${escapeHtml(row.nombre || 'Sin nombre')}</strong><small>${escapeHtml(row.rut)} · ${escapeHtml(row.region || 'Sin zona')}</small></div></td>
      <td>${escapeHtml(row.supervisor || 'Sin supervisor')}</td>
      <td><span class="phase-badge ${row.phase === 1 ? 'pilot' : 'future'}">${escapeHtml(row.phaseLabel || 'Fase 2')}</span></td>
      <td>${conceptValue(row.movilizacion)}</td>
      <td>${conceptValue(row.eficiencia)}</td>
      <td>${conceptValue(row.concurso, true)}</td>
      <td><div class="amount-with-detail">${conceptValue(row.hheeMonto)}<small>${row.hheeHoras ? `${row.hheeHoras} h` : 'Sin horas'}</small></div></td>
      <td>${desgasteValue(row)}</td>
      <td>${conceptValue(row.maestroGuia, true)}</td>
      <td>${conceptValue(row.compensacion, true)}</td>
      <td><span class="table-status ${resultStatus(row).toLowerCase()}">${resultStatus(row)}</span></td>
    </tr>`).join('');
  }
  $('#result-count').textContent = `${formatNumber(state.filteredRows.length)} trabajador${state.filteredRows.length === 1 ? '' : 'es'}`;
  $('#table-footer-info').textContent = state.rows.length ? `${formatNumber(state.filteredRows.length)} de ${formatNumber(state.rows.length)} trabajadores · valores automáticos y manuales por completar` : 'Sin datos calculados';
  updateApprovalSummary();
  renderSupervisorSelectors();
  renderManualSummary();
}

function formatStartDate(value) {
  const text = String(value ?? '').trim();
  if (!text) return 'Sin fecha';
  const iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (iso) return `${iso[3].padStart(2, '0')}/${iso[2].padStart(2, '0')}/${iso[1]}`;
  const local = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (local) return `${local[1].padStart(2, '0')}/${local[2].padStart(2, '0')}/${local[3]}`;
  return text;
}

function worksheetAmount(value) {
  return value === null || value === undefined ? '' : String(value);
}

function worksheetAutomatic(value, ready, currency = false) {
  if (!ready || value === null || value === undefined) return '<span class="pending-value">Pendiente</span>';
  return currency ? formatCurrency(value) : formatNumber(value);
}

function renderSupervisorWorksheet() {
  const body = $('#supervisor-worksheet-body');
  if (!body) return;
  const rows = visibleRows();
  const centralSourcesReady = CENTRAL_SOURCE_IDS.every(sourceReady);
  $('#supervisor-worksheet-count').textContent = `${formatNumber(rows.length)} trabajador${rows.length === 1 ? '' : 'es'}`;
  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="18" class="empty-table"><span>◌</span><strong>${state.processStarted ? 'No hay trabajadores asignados a este piloto' : 'La planilla está lista para recibir los datos del equipo'}</strong><small>${state.processStarted ? 'Revisa el maestro REX+ y el centro de costo asignado a Juan Alzualde.' : 'RRHH o Admin debe cargar el libro REX+, el maestro y los archivos centralizados para completar las filas.'}</small></td></tr>`;
  } else {
    body.innerHTML = rows.map((row) => `<tr data-worker-key="${escapeHtml(row.key)}">
      <td class="worksheet-fixed worksheet-rut">${escapeHtml(row.rut)}</td>
      <td class="worksheet-fixed worksheet-name"><strong>${escapeHtml(row.nombre || 'Sin nombre')}</strong></td>
      <td>${escapeHtml(row.cargo || 'Sin cargo')}</td>
      <td>${escapeHtml(row.centroCosto || 'Sin centro')}</td>
      <td>${escapeHtml(formatStartDate(row.fechaInicio))}</td>
      <td class="worksheet-number">${worksheetAutomatic(row.diasTrabajados, row.payrollFound)}</td>
      <td class="worksheet-number">${worksheetAutomatic(row.sueldoBase, row.payrollFound, true)}</td>
      <td class="worksheet-number">${worksheetAutomatic(row.hheeHoras, row.payrollFound)}</td>
      <td class="worksheet-number">${worksheetAutomatic(row.hheeMonto, row.payrollFound, true)}</td>
      <td class="worksheet-number">${worksheetAutomatic(row.fuel, centralSourcesReady, true)}</td>
      <td class="worksheet-number">${worksheetAutomatic(row.tag, centralSourcesReady, true)}</td>
      <td class="worksheet-number">${worksheetAutomatic(row.eficiencia, centralSourcesReady, true)}</td>
      <td class="worksheet-number">${worksheetAutomatic(row.desgasteContractual, row.payrollFound, true)}</td>
      <td><input class="worksheet-input" type="number" min="0" step="1" data-worksheet-field="concurso" data-worker-key="${escapeHtml(row.key)}" value="${escapeHtml(worksheetAmount(row.concurso))}" placeholder="$0" aria-label="Concurso para ${escapeHtml(row.nombre || row.rut)}"></td>
      <td><input class="worksheet-input" type="number" min="0" step="1" data-worksheet-field="desgasteAdicional" data-worker-key="${escapeHtml(row.key)}" value="${escapeHtml(worksheetAmount(row.desgasteAdicional))}" placeholder="$0" aria-label="Desgaste adicional para ${escapeHtml(row.nombre || row.rut)}"></td>
      <td><input class="worksheet-input" type="number" min="0" step="1" data-worksheet-field="maestroGuia" data-worker-key="${escapeHtml(row.key)}" value="${escapeHtml(worksheetAmount(row.maestroGuia))}" placeholder="$0" aria-label="Maestro guía para ${escapeHtml(row.nombre || row.rut)}"></td>
      <td><input class="worksheet-input" type="number" min="0" step="1" data-worksheet-field="compensacion" data-worker-key="${escapeHtml(row.key)}" value="${escapeHtml(worksheetAmount(row.compensacion))}" placeholder="$0" aria-label="Compensación para ${escapeHtml(row.nombre || row.rut)}"></td>
      <td><textarea class="worksheet-observation" rows="1" data-worksheet-field="observaciones" data-worker-key="${escapeHtml(row.key)}" aria-label="Observaciones para ${escapeHtml(row.nombre || row.rut)}">${escapeHtml(row.observaciones || '')}</textarea></td>
    </tr>`).join('');
  }
  $('#supervisor-send-button').disabled = !rows.length || state.supervisorSubmitted;
  const supervisorFile = $('#file-supervisor-variables');
  if (supervisorFile) supervisorFile.disabled = state.supervisorSubmitted || !['supervisor', 'superadmin'].includes(state.profile);
  const supervisorTemplate = $('#download-supervisor-template');
  if (supervisorTemplate) supervisorTemplate.disabled = state.supervisorSubmitted || !['supervisor', 'gerente', 'superadmin'].includes(state.profile);
  body.querySelectorAll('[data-worksheet-field]').forEach((field) => { field.disabled = state.supervisorSubmitted; });
}

function handleSupervisorWorksheetChange(event) {
  const field = event.target.closest('[data-worksheet-field]');
  if (!field || state.supervisorSubmitted) return;
  const row = state.rows.find((item) => item.key === field.dataset.workerKey);
  if (!row) return;
  if (field.dataset.worksheetField === 'observaciones') {
    row.observaciones = field.value.trim();
    return;
  }
  row[field.dataset.worksheetField] = field.value.trim() === '' ? null : round(parseAmount(field.value));
  if (field.dataset.worksheetField === 'desgasteAdicional') {
    row.desgaste = round(row.desgasteContractual + Number(row.desgasteAdicional || 0));
  }
}

function sendSupervisorVariables() {
  const rows = visibleRows();
  if (!rows.length) return;
  state.supervisorSubmitted = true;
  $('#supervisor-submit-feedback').className = 'supervisor-submit-feedback success';
  $('#supervisor-submit-feedback').textContent = `Variables enviadas para revisión de Pedro · ${formatNumber(rows.length)} trabajadores del equipo de Juan Alzualde.`;
  renderSupervisorWorksheet();
}

function updateApprovalSummary() {
  const rows = visibleRows();
  const pending = rows.filter((row) => resultStatus(row) === 'PENDIENTE').length;
  const ready = rows.filter((row) => resultStatus(row) === 'LISTO').length;
  const approved = rows.filter((row) => resultStatus(row) === 'APROBADO').length;
  $('#kpi-pending').textContent = formatNumber(pending);
  $('#pending-count').textContent = formatNumber(pending);
  $('#approval-pending').textContent = formatNumber(pending);
  $('#approval-ready').textContent = formatNumber(ready);
  $('#approval-approved').textContent = formatNumber(approved);
  $('#approve-visible').disabled = !state.filteredRows.length || state.filteredRows.some((row) => resultStatus(row) === 'PENDIENTE');
  $('#reject-visible').disabled = !state.filteredRows.length;
}

function populateSupervisorFilter() {
  const supervisors = [...new Set(visibleRows().map((row) => row.supervisor).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
  $('#supervisor-filter').innerHTML = '<option value="">Todos los supervisores</option>' + supervisors.map((supervisor) => `<option value="${escapeHtml(supervisor)}">${escapeHtml(supervisor)}</option>`).join('');
  $('#supervisor-filter').hidden = state.profile === 'supervisor';
}

const manualConceptMeta = {
  concurso: { label: 'Concurso', valueLabel: 'Monto a informar', suffix: '$', step: '1', help: 'Informa el valor definido por el mantenedor para este técnico. Si no corresponde, ingresa 0.' },
  desgasteAdicional: { label: 'Desgaste adicional', valueLabel: 'Monto adicional', suffix: '$', step: '1', help: 'La base contractual viene desde REX+. Informa sólo el adicional; el total se recalcula automáticamente.' },
  maestroGuia: { label: 'Maestro guía', valueLabel: 'Monto estándar', suffix: '$', step: '1', help: 'Selecciona al técnico que corresponde y registra el monto vigente del mantenedor. Si no corresponde, ingresa 0.' },
  compensacion: { label: 'Compensación discrecional', valueLabel: 'Monto a informar', suffix: '$', step: '1', help: 'Informa el monto acordado para este técnico. El Gerente podrá modificarlo individualmente.' },
  hhee: { label: 'HHEE · cantidad de horas', valueLabel: 'Cantidad de HHEE', suffix: 'horas', step: '0.01', help: 'La cantidad se informa aquí. El monto se recalcula con sueldo base ÷ 180 × 1,5.' },
};

function renderSupervisorSelectors() {
  const supervisorSelect = $('#supervisor-entry-filter');
  const workerSelect = $('#supervisor-worker-select');
  if (!supervisorSelect || !workerSelect) return;
  const previousSupervisor = supervisorSelect.value;
  const previousWorker = workerSelect.value;
  const supervisors = [...new Set(visibleRows().map((row) => row.supervisor).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
  supervisorSelect.innerHTML = '<option value="">Todos los supervisores</option>' + supervisors.map((supervisor) => `<option value="${escapeHtml(supervisor)}">${escapeHtml(supervisor)}</option>`).join('');
  if (state.profile === 'supervisor') {
    supervisorSelect.value = PROFILE_META.supervisor.supervisor;
    supervisorSelect.disabled = true;
  } else {
    supervisorSelect.disabled = false;
    supervisorSelect.value = supervisors.includes(previousSupervisor) ? previousSupervisor : '';
  }
  const workers = visibleRows().filter((row) => !supervisorSelect.value || row.supervisor === supervisorSelect.value);
  workerSelect.innerHTML = workers.length
    ? workers.map((row) => `<option value="${escapeHtml(row.key)}">${escapeHtml(row.nombre || 'Sin nombre')} · ${escapeHtml(row.rut)}</option>`).join('')
    : '<option value="">Carga las fuentes para comenzar</option>';
  workerSelect.disabled = !workers.length;
  workerSelect.value = workers.some((row) => row.key === previousWorker) ? previousWorker : (workers[0]?.key || '');
  $('#supervisor-value-input').disabled = !workerSelect.value;
  $('#save-supervisor-entry').disabled = !workerSelect.value;
  updateSupervisorConceptHelp();
}

function updateSupervisorConceptHelp() {
  const concept = $('#supervisor-concept-select')?.value || 'concurso';
  const metadata = manualConceptMeta[concept];
  if (!metadata) return;
  $('#supervisor-value-label').textContent = metadata.valueLabel;
  $('#supervisor-value-suffix').textContent = metadata.suffix;
  $('#supervisor-value-input').step = metadata.step;
  $('#supervisor-value-input').placeholder = concept === 'hhee' ? 'Ej. 4,5' : 'Ingresa un monto';
  $('#supervisor-entry-help').textContent = metadata.help;
}

function renderManualSummary() {
  const manualKeys = ['concurso', 'desgasteAdicional', 'maestroGuia', 'compensacion'];
  const rows = visibleRows();
  const completed = rows.reduce((total, row) => total + manualKeys.filter((key) => row[key] !== null && row[key] !== undefined).length, 0);
  const pending = rows.reduce((total, row) => total + manualKeys.filter((key) => row[key] === null || row[key] === undefined).length, 0);
  $('#manual-complete-count').textContent = formatNumber(completed);
  $('#manual-pending-count').textContent = formatNumber(pending);
  const list = $('#manual-entry-list');
  if (!rows.length) {
    list.innerHTML = '<div class="manual-empty"><span>✎</span><strong>Aún no hay información manual</strong><small>Cuando el supervisor guarde un valor, aparecerá aquí para su revisión.</small></div>';
    return;
  }
  list.innerHTML = rows.slice(0, 8).map((row) => `<div class="manual-entry-row"><div class="manual-worker"><strong>${escapeHtml(row.nombre || 'Sin nombre')}</strong><small>${escapeHtml(row.rut)} · ${escapeHtml(row.supervisor || 'Sin supervisor')}</small></div><div class="manual-chip-list">${manualKeys.map((key) => `<span class="manual-chip ${row[key] === null || row[key] === undefined ? 'pending' : 'complete'}"><i></i>${manualConceptMeta[key].label}${row[key] === null || row[key] === undefined ? '' : ` · ${formatCurrency(row[key])}`}</span>`).join('')}</div></div>`).join('');
}

function hasInputValue(value) {
  return String(value ?? '').trim() !== '';
}

async function handleSupervisorFile(file) {
  if (state.supervisorSubmitted) return;
  const feedback = $('#supervisor-feedback');
  try {
    await loadXlsx();
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true, cellText: true, raw: false });
    const sheets = workbook.SheetNames.map((name) => {
      const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '', raw: false });
      return { name, ...recordsFromMatrix(matrix, 'supervisor') };
    });
    const records = sheets.find((sheet) => sheet.records.length)?.records || [];
    let updated = 0;
    let skipped = 0;
    records.forEach((record) => {
      const key = normalizeRut(pick(record, ['RUT', 'Rut', 'Empleado']));
      const row = state.rows.find((item) => item.key === key && visibleRows().some((visible) => visible.key === key));
      if (!row) {
        skipped += 1;
        return;
      }
      const fields = [
        ['concurso', ['CONCURSO', 'Concurso']],
        ['desgasteAdicional', ['DESGASTE ADICIONAL', 'Desgaste adicional']],
        ['maestroGuia', ['MAESTRO GUIA', 'Maestro guía', 'Maestro guia']],
        ['compensacion', ['COMPENSACION', 'Compensación', 'Compensacion']],
      ];
      fields.forEach(([field, names]) => {
        const raw = pick(record, names);
        if (hasInputValue(raw)) row[field] = round(parseAmount(raw));
      });
      const hhee = pick(record, ['HHEE CANTIDAD', 'HHEE', 'Cantidad HHEE', 'Horas extra']);
      if (hasInputValue(hhee)) {
        row.hheeHoras = round(parseAmount(hhee));
        row.hheeMontoInformado = 0;
        row.hheeMonto = round(row.hheeHoras * row.hheeRate);
      }
      const observations = pick(record, ['OBSERVACIONES', 'Observaciones', 'OBSERVACION', 'Comentario']);
      if (hasInputValue(observations)) row.observaciones = String(observations).trim();
      row.desgaste = round(row.desgasteContractual + Number(row.desgasteAdicional || 0));
      row.rejected = false;
      updated += 1;
    });
    feedback.className = 'supervisor-feedback success';
    feedback.textContent = `${updated} registro${updated === 1 ? '' : 's'} cargado${updated === 1 ? '' : 's'} desde ${file.name}${skipped ? ` · ${skipped} omitido${skipped === 1 ? '' : 's'} por no pertenecer a tu equipo` : ''}.`;
    renderResults();
    renderSupervisorWorksheet();
  } catch (error) {
    feedback.className = 'supervisor-feedback error';
    feedback.textContent = `No se pudo leer ${file.name}: ${error.message}`;
  }
}

function saveSupervisorEntry() {
  const row = state.rows.find((item) => item.key === $('#supervisor-worker-select').value);
  const concept = $('#supervisor-concept-select').value;
  const rawValue = $('#supervisor-value-input').value.trim();
  const value = parseAmount(rawValue);
  if (!row || rawValue === '' || value < 0) {
    $('#supervisor-feedback').className = 'supervisor-feedback error';
    $('#supervisor-feedback').textContent = 'Selecciona un trabajador e informa un valor igual o mayor que cero.';
    return;
  }
  if (concept === 'hhee') {
    row.hheeHoras = round(value);
    row.hheeMontoInformado = 0;
    row.hheeMonto = round(row.hheeHoras * row.hheeRate);
  } else if (concept === 'desgasteAdicional') {
    row.desgasteAdicional = round(value);
    row.desgaste = round(row.desgasteContractual + row.desgasteAdicional);
  } else {
    row[concept] = round(value);
  }
  row.rejected = false;
  $('#supervisor-feedback').className = 'supervisor-feedback success';
  $('#supervisor-feedback').textContent = `${manualConceptMeta[concept].label} ${state.profile === 'gerente' ? 'editado' : 'guardado'} para ${row.nombre || row.rut}.`;
  $('#supervisor-value-input').value = '';
  renderResults();
}

function calculate() {
  state.rows = buildRows();
  state.processStarted = true;
  state.supervisorSubmitted = false;
  $('#kpi-workers').textContent = formatNumber(state.rows.length);
  $('#kpi-concepts').textContent = state.rows.length ? '4 / 6' : '0 / 6';
  $('#last-sync').textContent = new Intl.DateTimeFormat('es-CL', { hour: '2-digit', minute: '2-digit' }).format(new Date());
  $('#validation-summary').textContent = `${formatNumber(state.rows.length)} trabajadores listos para calcular`;
  populateSupervisorFilter();
  renderResults();
  renderSupervisorWorksheet();
  setWorkflow(2);
  $('#calculo').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function ensureViewerRows() {
  if (!state.rows.length && state.processStarted) {
    state.rows = buildRows();
    $('#kpi-workers').textContent = formatNumber(state.rows.length);
    $('#kpi-concepts').textContent = state.rows.length ? '4 / 6' : '0 / 6';
  }
}

function approveVisible() {
  if (state.filteredRows.some((row) => resultStatus(row) === 'PENDIENTE')) return;
  state.filteredRows.forEach((row) => { row.approved = true; row.rejected = false; });
  renderResults();
}

function rejectVisible() {
  if (!state.filteredRows.length) return;
  state.filteredRows.forEach((row) => { row.approved = false; row.rejected = true; });
  renderResults();
}

function setWorkflow(step) {
  $$('.workflow-step').forEach((item, index) => item.classList.toggle('active', index < step));
}

function periodLabel(value) {
  if (!value) return 'Sin período';
  const [year, month] = value.split('-');
  return new Intl.DateTimeFormat('es-CL', { month: 'long', year: 'numeric' }).format(new Date(Number(year), Number(month) - 1, 1)).replace(/^./, (char) => char.toUpperCase());
}

function updatePeriod() {
  const label = periodLabel($('#period').value);
  $('#period-label').textContent = label;
  $('#sidebar-period').textContent = label;
}

function applyProfileView() {
  const metadata = PROFILE_META[state.profile];
  ensureViewerRows();
  $('#profile-select').value = state.profile;
  $('#profile-avatar').textContent = metadata.initials;
  $('#page-subtitle').textContent = state.profile === 'supervisor'
    ? 'Completa y envía las variables del personal a tu cargo.'
    : 'Calcula, revisa y prepara los conceptos que complementan el bono de producción.';
  $('#profile-banner').querySelector('.profile-banner-icon').textContent = metadata.initials;
  $('#profile-banner').querySelector('strong').textContent = state.profile === 'supervisor' && !state.processStarted ? 'Vista Supervisor · Proceso pendiente' : metadata.banner;
  $('#profile-banner').querySelector('small').textContent = state.profile === 'supervisor' && !state.processStarted ? 'RRHH debe cargar el libro, el maestro y los libros centralizados antes de habilitar la planilla.' : metadata.detail;
  const sourceNote = $('#source-access-note');
  sourceNote.textContent = state.profile === 'rrhh' ? 'RRHH puede cargar y reemplazar las fuentes del período.' : state.profile === 'admin' ? 'Admin artBPO puede cargar e iniciar el proceso.' : state.profile === 'superadmin' ? 'SuperAdmin tiene acceso total para cargar, procesar y administrar el período.' : 'Las fuentes se administran desde RRHH o Admin.';
  $$('[data-visible-for]').forEach((section) => {
    const profileVisible = section.dataset.visibleFor.split(',').includes(state.profile);
    const processVisible = !section.hasAttribute('data-supervisor-process') || state.processStarted || ['rrhh', 'admin', 'superadmin'].includes(state.profile);
    section.hidden = !profileVisible || !processVisible;
  });
  $$('[data-nav]').forEach((link) => {
    const target = document.querySelector(link.getAttribute('href'));
    const profileVisible = !target?.dataset.visibleFor || target.dataset.visibleFor.split(',').includes(state.profile);
    const processVisible = !target?.hasAttribute('data-supervisor-process') || state.processStarted || ['rrhh', 'admin', 'superadmin'].includes(state.profile);
    link.hidden = !profileVisible || !processVisible;
  });
  $('#supervisor-process-waiting').hidden = state.profile !== 'supervisor' || state.processStarted;
  const canUpload = ['rrhh', 'admin', 'superadmin'].includes(state.profile);
  $$('[data-file-input]').forEach((input) => { input.disabled = !canUpload; });
  $$('[data-browse]').forEach((button) => { button.disabled = !canUpload; });
  const masterApiButton = $('#load-master-api');
  if (masterApiButton) masterApiButton.disabled = !canUpload;
  $$('label.mini-upload').forEach((label) => label.classList.toggle('read-only', !canUpload));
  $('#download-template').disabled = !canUpload;
  const supervisorFile = $('#file-supervisor-variables');
  if (supervisorFile) supervisorFile.disabled = !['supervisor', 'superadmin'].includes(state.profile);
  const supervisorTemplate = $('#download-supervisor-template');
  if (supervisorTemplate) supervisorTemplate.disabled = !['supervisor', 'gerente', 'superadmin'].includes(state.profile);
  $('#calculate-button').hidden = ['supervisor', 'gerente'].includes(state.profile);
  $('#calculate-button-label').textContent = state.profile === 'admin' ? 'Iniciar y calcular' : 'Calcular conceptos';
  if (state.rows.length) {
    populateSupervisorFilter();
    renderResults();
    renderSupervisorWorksheet();
  } else {
    renderSupervisorSelectors();
    renderManualSummary();
    renderSupervisorWorksheet();
  }
}

function downloadTemplate() {
  loadXlsx().then(() => {
    const headers = [['RUT', 'NOMBRE TAC', 'SUPERVISOR', 'ZONA', 'Suma de TAG', 'Suma de Multa', 'Suma de Combustible - Consumo', 'Suma de Asignación Combustible', 'Suma de TOTAL CONSUMO']];
    const workbook = XLSX.utils.book_new();
    ['TECNICOS', 'ADMINISTRATIVOS', 'OTROS'].forEach((name) => XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(headers), name));
    XLSX.writeFile(workbook, 'Template_fuente_centralizada_NOW.xlsx');
  });
}

function downloadSupervisorTemplate() {
  loadXlsx().then(() => {
    const rows = visibleRows();
    const headers = ['RUT', 'NOMBRE', 'CARGO', 'CENTRO DE COSTO', 'FECHA DE INICIO', 'DÍAS TRABAJADOS', 'SUELDO BASE', 'HHEE CANTIDAD', 'HHEE MONTO', 'COMBUSTIBLE', 'TAG', 'EFICIENCIA', 'DESGASTE CONTRACTUAL', 'CONCURSO', 'DESGASTE ADICIONAL', 'MAESTRO GUIA', 'COMPENSACION', 'OBSERVACIONES'];
    const data = [headers, ...rows.map((row) => [row.rut, row.nombre, row.cargo, row.centroCosto, formatStartDate(row.fechaInicio), row.diasTrabajados, row.sueldoBase, row.hheeHoras, row.hheeMonto, row.fuel, row.tag, row.eficiencia, row.desgasteContractual, '', '', '', '', ''])];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(data), 'VARIABLES');
    XLSX.writeFile(workbook, 'Template_variables_supervisor_NOW.xlsx');
  });
}

function wireNavigation() {
  $$('[data-nav]').forEach((link) => link.addEventListener('click', () => {
    $$('[data-nav]').forEach((item) => item.classList.toggle('active', item === link));
  }));
}

function wireEvents() {
  wireUploadButtons();
  $('#calculate-button').addEventListener('click', calculate);
  $('#approve-visible').addEventListener('click', approveVisible);
  $('#reject-visible').addEventListener('click', rejectVisible);
  $('#download-template').addEventListener('click', downloadTemplate);
  $('#download-supervisor-template').addEventListener('click', downloadSupervisorTemplate);
  $('#load-master-api').addEventListener('click', loadMasterFromApi);
  $('#period').addEventListener('change', updatePeriod);
  $('#refresh-button').addEventListener('click', () => { renderSourceSummary(); renderResults(); });
  $('#search-results').addEventListener('input', renderResults);
  $('#supervisor-filter').addEventListener('change', renderResults);
  $('#result-status-filter').addEventListener('change', renderResults);
  $('#supervisor-entry-filter').addEventListener('change', renderSupervisorSelectors);
  $('#supervisor-concept-select').addEventListener('change', updateSupervisorConceptHelp);
  $('#save-supervisor-entry').addEventListener('click', saveSupervisorEntry);
  $('#supervisor-worksheet-body').addEventListener('change', handleSupervisorWorksheetChange);
  $('#supervisor-send-button').addEventListener('click', sendSupervisorVariables);
  $('#profile-select').addEventListener('change', (event) => {
    state.profile = event.target.value;
    applyProfileView();
  });
  wireNavigation();
}

updatePeriod();
wireEvents();
renderSourceSummary();
updateApprovalSummary();
applyProfileView();
loadAugustPilotSources();
