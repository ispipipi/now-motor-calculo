const sourceDefinitions = {
  rex: {
    label: 'Libro de remuneraciones REX+',
    description: 'Fuente principal de haberes, descuentos, días y líquido informado por REX+.',
    short: 'REX+',
    required: true,
    accept: '.xlsx,.xls,.csv',
    sheets: ['DETALLE'],
    columns: ['Nombre', 'Rut', 'Sueldo Base', 'Sueldo Líquido'],
  },
  pago: {
    label: 'Pago TAC',
    description: 'Datos operativos y estructura de la preliquidación.',
    short: 'TAC',
    required: true,
    accept: '.xlsx,.xls,.csv',
    sheets: ['LIBRO REM', 'VARIABLES', 'NOMINA RRHH', 'PRE_LIQ TAC'],
  },
  novedades: {
    label: 'Novedades RRHH (complemento)',
    description: 'Detalle complementario para HHEE, vacaciones, licencias y otras novedades.',
    short: 'RRHH',
    required: false,
    accept: '.xlsx,.xls,.csv',
    sheets: ['VIATICO', 'ANTICIPO', 'PRESTAMO CAJA', 'AHORRO', 'CARGA FAMILIAR', 'VACACIONES', 'RETENCION JUDICIAL', 'SEGURO', 'LICENCIAS', 'HHEE', 'TAG'],
  },
  bono: {
    label: 'Bono de producción',
    description: 'Fuente externa con RUT, nombre y monto.',
    short: '$',
    required: false,
    accept: '.xlsx,.xls,.csv,.json',
    sheets: [],
    columns: ['RUT', 'NOMBRE', 'MONTO'],
  },
};

const requiredSourceIds = Object.entries(sourceDefinitions)
  .filter(([, definition]) => definition.required)
  .map(([id]) => id);

const state = { sources: {} };
let xlsxLoadPromise;
const sourceGrid = document.querySelector('#source-grid');
const validationList = document.querySelector('#validation-list');
const validationEmpty = document.querySelector('#validation-empty');
const continueButton = document.querySelector('#continue-button');
const clearButton = document.querySelector('#clear-button');
const actionBar = document.querySelector('.action-bar');
const reviewStage = document.querySelector('#review-stage');
const reviewItems = document.querySelector('#review-items');
const backButton = document.querySelector('#back-button');
const reviewToCalculationButton = document.querySelector('#review-to-calculation-button');
const calculationStage = document.querySelector('#calculation-stage');
const calculationBackButton = document.querySelector('#calculation-back-button');
const calculateCombustibleButton = document.querySelector('#calculate-combustible-button');
const combustibleResultsStage = document.querySelector('#combustible-results-stage');
const combustibleBackButton = document.querySelector('#combustible-back-button');
const continueConcursoButton = document.querySelector('#continue-concurso-button');
const concursoHheeStage = document.querySelector('#concurso-hhee-stage');
const concursoHheeBackButton = document.querySelector('#concurso-hhee-back-button');
const continuePreliquidationButton = document.querySelector('#continue-preliquidation-button');
const preliquidationStage = document.querySelector('#preliquidation-stage');
const preliquidationBackButton = document.querySelector('#preliquidation-back-button');
const introSection = document.querySelector('.intro-section');
let preliquidationData = { rows: [], period: '' };

function normalize(value) {
  return String(value ?? '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');
}

function formatNumber(value) {
  return new Intl.NumberFormat('es-CL').format(value || 0);
}

function formatFileSize(bytes) {
  if (!bytes) return '';
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function sourceCard(definition, id) {
  const typeClass = definition.required ? 'required' : 'optional';
  const requiredText = definition.required ? 'Fuente requerida' : 'Fuente opcional';
  return `<article class="source-card ${typeClass}" data-source-card="${id}">
    <div class="source-top">
      <div class="source-icon">${definition.short}</div>
      <div class="source-heading"><h3>${definition.label}</h3><p>${definition.description}</p></div>
    </div>
    <label class="source-label" for="file-${id}">${requiredText}<span class="source-format">${definition.accept.replaceAll('.', '').replaceAll(',', ' / ').toUpperCase()}</span></label>
    <div class="drop-zone" data-drop-zone="${id}">
      <input id="file-${id}" data-file-input="${id}" type="file" accept="${definition.accept}">
      <button class="drop-button" data-browse="${id}" type="button">Seleccionar</button>
      <div class="drop-copy" data-file-copy="${id}"><strong>Sin archivo seleccionado</strong>Arrastra o busca un archivo</div>
    </div>
    <div class="source-status"><span class="status-pill neutral" data-card-status="${id}">Pendiente</span><span class="source-detail" data-card-detail="${id}"></span></div>
  </article>`;
}

function renderCards() {
  sourceGrid.innerHTML = Object.entries(sourceDefinitions).map(([id, definition]) => sourceCard(definition, id)).join('');
  for (const id of Object.keys(sourceDefinitions)) wireSourceCard(id);
}

function wireSourceCard(id) {
  const input = document.querySelector(`[data-file-input="${id}"]`);
  const browse = document.querySelector(`[data-browse="${id}"]`);
  const dropZone = document.querySelector(`[data-drop-zone="${id}"]`);
  browse.addEventListener('click', () => input.click());
  input.addEventListener('change', () => { if (input.files[0]) handleFile(id, input.files[0]); });
  for (const eventName of ['dragenter', 'dragover']) {
    dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.add('dragging'); });
  }
  for (const eventName of ['dragleave', 'drop']) {
    dropZone.addEventListener(eventName, (event) => { event.preventDefault(); dropZone.classList.remove('dragging'); });
  }
  dropZone.addEventListener('drop', (event) => { const file = event.dataTransfer.files[0]; if (file) handleFile(id, file); });
}

async function handleFile(id, file) {
  const definition = sourceDefinitions[id];
  const cardStatus = document.querySelector(`[data-card-status="${id}"]`);
  const cardDetail = document.querySelector(`[data-card-detail="${id}"]`);
  const fileCopy = document.querySelector(`[data-file-copy="${id}"]`);
  const dropZone = document.querySelector(`[data-drop-zone="${id}"]`);
  cardStatus.className = 'status-pill neutral';
  cardStatus.textContent = 'Revisando';
  cardDetail.textContent = '';
  fileCopy.innerHTML = `<strong>${escapeHtml(file.name)}</strong>${formatFileSize(file.size)}`;
  try {
    const report = await inspectFile(file, definition);
    state.sources[id] = { file, report };
    dropZone.classList.add('has-file');
    cardStatus.className = `status-pill ${report.level}`;
    cardStatus.textContent = report.statusLabel;
    cardDetail.textContent = report.detail;
  } catch (error) {
    state.sources[id] = { file, report: { level: 'error', statusLabel: 'No leído', detail: error.message, rows: 0, sheets: [], warnings: [error.message] } };
    dropZone.classList.remove('has-file');
    cardStatus.className = 'status-pill error';
    cardStatus.textContent = 'No leído';
    cardDetail.textContent = error.message;
  }
  renderValidation();
}

async function inspectFile(file, definition) {
  const extension = file.name.split('.').pop().toLowerCase();
  let workbook;
  if (extension === 'json') {
    const data = JSON.parse(await file.text());
    const rows = Array.isArray(data) ? data : Array.isArray(data.rows) ? data.rows : [data];
    return validateReport([{ name: 'JSON', rows, records: rows }], definition, file.name);
  }
  await loadXlsx();
  const buffer = await file.arrayBuffer();
  workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellNF: false, cellText: true });
  const sheets = workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
    const headerIndex = findHeaderIndex(matrix);
    const header = headerIndex >= 0 ? matrix[headerIndex].map((cell) => String(cell).trim()).filter(Boolean) : [];
    const rows = headerIndex >= 0 ? matrix.slice(headerIndex + 1).filter((row) => row.some((cell) => String(cell).trim() !== '')) : [];
    const rawHeader = headerIndex >= 0 ? matrix[headerIndex].map((cell) => String(cell).trim()) : [];
    const records = rows.map((row) => {
      const seenHeaders = new Map();
      return rawHeader.reduce((record, key, index) => {
        if (!key) return record;
        const count = (seenHeaders.get(key) ?? 0) + 1;
        seenHeaders.set(key, count);
        record[count === 1 ? key : `${key}_${count}`] = row[index] ?? '';
        return record;
      }, {});
    });
    return { name, header, rows, records };
  });
  return validateReport(sheets, definition, file.name);
}

function findHeaderIndex(matrix) {
  const detailedHeaderIndex = matrix.findIndex((row) => {
    const headers = row.map(normalize);
    const hasName = headers.includes('NOMBRE TAC') || headers.includes('NOMBRE');
    return headers.includes('RUT') && hasName && headers.includes('SUELDO BASE');
  });
  return detailedHeaderIndex >= 0
    ? detailedHeaderIndex
    : matrix.findIndex((row) => row.some((cell) => String(cell).trim() !== ''));
}

function loadXlsx() {
  if (window.XLSX) return Promise.resolve();
  if (xlsxLoadPromise) return xlsxLoadPromise;
  xlsxLoadPromise = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'vendor/xlsx.full.min.js';
    script.onload = () => resolve();
    script.onerror = () => reject(new Error('No se pudo cargar el lector de planillas.'));
    document.head.appendChild(script);
  });
  return xlsxLoadPromise;
}

function validateReport(sheets, definition, filename) {
  const warnings = [];
  const normalizedNames = new Set(sheets.map((sheet) => normalize(sheet.name)));
  const missingSheets = (definition.sheets || []).filter((sheet) => !matchesExpectedSheet(sheet, normalizedNames));
  if (missingSheets.length) warnings.push(`Hojas no encontradas: ${missingSheets.join(', ')}`);
  if (definition.columns?.length) {
    const allColumns = new Set(sheets.flatMap((sheet) => sheet.header.map(normalize)));
    const missingColumns = definition.columns.filter((column) => !allColumns.has(normalize(column)));
    if (missingColumns.length) warnings.push(`Columnas requeridas no encontradas: ${missingColumns.join(', ')}`);
  }
  const rows = sheets.reduce((total, sheet) => total + sheet.rows.length, 0);
  const nonEmptySheets = sheets.filter((sheet) => sheet.rows.length > 0).length;
  const level = warnings.length ? 'warning' : 'valid';
  return {
    level,
    statusLabel: warnings.length ? 'Revisar alertas' : 'Estructura OK',
    detail: `${formatNumber(rows)} fila${rows === 1 ? '' : 's'} · ${nonEmptySheets} hoja${nonEmptySheets === 1 ? '' : 's'} con datos`,
    rows,
    sheets: sheets.map((sheet) => ({ name: sheet.name, rows: sheet.rows.length, header: sheet.header, records: sheet.records ?? [] })),
    warnings,
    filename,
  };
}

function matchesExpectedSheet(expectedName, normalizedNames) {
  const normalized = normalize(expectedName);
  const aliases = normalized === 'PRE LIQ TAC' || normalized === 'PRE LIQU TAC'
    ? ['PRE LIQ TAC', 'PRE LIQU TAC']
    : [normalized];
  return aliases.some((alias) => normalizedNames.has(alias));
}

function renderValidation() {
  const entries = Object.entries(state.sources);
  const loaded = entries.length;
  const rows = entries.reduce((total, [, item]) => total + (item.report.rows || 0), 0);
  const warnings = entries.reduce((total, [, item]) => total + (item.report.warnings?.length || 0), 0);
  document.querySelector('#loaded-count').textContent = formatNumber(loaded);
  document.querySelector('#row-count').textContent = formatNumber(rows);
  document.querySelector('#warning-count').textContent = formatNumber(warnings);
  validationEmpty.hidden = loaded > 0;
  validationList.innerHTML = entries.map(([id, item]) => validationRow(id, item)).join('');
  const requiredReady = requiredSourceIds.every((id) => state.sources[id]);
  continueButton.disabled = !requiredReady;
  document.querySelector('#action-message').textContent = requiredReady ? 'Las fuentes requeridas están cargadas.' : 'Carga el Libro REX+ y Pago TAC para continuar.';
  if (!reviewStage.hidden) renderReview();
}

function validationRow(id, item) {
  const report = item.report;
  const sheetNames = report.sheets.map((sheet) => sheet.name).join(', ') || 'Sin hojas';
  return `<div class="validation-row"><strong>${sourceDefinitions[id].label}</strong><span>${escapeHtml(item.file.name)}</span><span class="validation-sheets" title="${escapeHtml(sheetNames)}">${escapeHtml(sheetNames)}</span><span class="status-pill ${report.level}">${report.statusLabel}</span><small>${escapeHtml(report.detail)}</small></div>`;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' })[character]);
}

function setActiveStep(stepNumber) {
  for (const step of document.querySelectorAll('.step')) step.classList.toggle('active', step.dataset.step === String(stepNumber));
}

function renderReview() {
  const entries = Object.entries(state.sources);
  const warnings = entries.reduce((total, [, item]) => total + (item.report.warnings?.length || 0), 0);
  const rows = entries.reduce((total, [, item]) => total + (item.report.rows || 0), 0);
  const required = requiredSourceIds.filter((id) => state.sources[id]).length;
  document.querySelector('#review-required-count').textContent = `${required} / ${requiredSourceIds.length}`;
  document.querySelector('#review-row-count').textContent = formatNumber(rows);
  document.querySelector('#review-warning-count').textContent = formatNumber(warnings);
  document.querySelector('#review-stage-status').textContent = warnings ? 'Revisar alertas' : 'Sin alertas de estructura';
  reviewItems.innerHTML = entries.length ? entries.map(([id, item]) => reviewRow(id, item)).join('') : '<div class="validation-empty"><p>No hay fuentes disponibles.</p></div>';
}

function pickColumn(record, names) {
  const columns = new Map(Object.entries(record).map(([key, value]) => [normalize(key), value]));
  for (const name of names) {
    const value = columns.get(normalize(name));
    if (value !== undefined && value !== null && String(value).trim() !== '') return value;
  }
  return '';
}

function parseAmount(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const text = String(value ?? '').trim().replace(/[$\s%]/g, '');
  if (!text) return 0;
  const hasComma = text.includes(',');
  const hasDot = text.includes('.');
  const normalized = hasComma && hasDot
    ? (text.lastIndexOf(',') > text.lastIndexOf('.') ? text.replace(/\./g, '').replace(',', '.') : text.replace(/,/g, ''))
    : hasComma
      ? (text.split(',').slice(1).every((part) => part.length === 3) ? text.replace(/,/g, '') : text.replace(',', '.'))
      : (hasDot && text.split('.').slice(1).every((part) => part.length === 3) ? text.replace(/\./g, '') : text);
  const amount = Number(normalized);
  return Number.isFinite(amount) ? amount : 0;
}

function roundAmount(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function formatCurrency(value) {
  return new Intl.NumberFormat('es-CL', { style: 'currency', currency: 'CLP', maximumFractionDigits: 0 }).format(Math.round(value || 0));
}

function buildCombustibleResult() {
  const pago = state.sources.pago?.report;
  const variables = pago?.sheets.find((sheet) => normalize(sheet.name) === 'VARIABLES');
  if (!variables?.records?.length) return { rows: [], sourceName: variables?.name ?? 'Hoja Variables', warning: 'No se encontraron registros en la hoja Variables.' };

  const rows = variables.records.map((record) => {
    const rut = String(pickColumn(record, ['RUT TAC', 'RUT'])).trim();
    if (!rut) return null;
    const asignacion = parseAmount(pickColumn(record, ['Asignacion de combustible y TAG', 'Asignación de combustible y TAG']));
    const adicional = parseAmount(pickColumn(record, ['Adicional']));
    const total = roundAmount(asignacion + adicional);
    const consumo = parseAmount(pickColumn(record, ['Consumo']));
    const saldo = roundAmount(total - consumo);
    return {
      rut,
      nombre: String(pickColumn(record, ['Nombre TAC', 'Nombre', 'NOMBRE'])).trim(),
      asignacion,
      adicional,
      total,
      consumo,
      saldo,
      movilizacionEspecial: saldo > 0 ? saldo : 0,
      eficienciaConsumo: saldo < 0 ? saldo : 0,
    };
  }).filter(Boolean);

  return {
    rows,
    sourceName: variables.name,
    warning: rows.length ? '' : 'La hoja Variables no contiene técnicos identificables.',
  };
}

function renderCombustibleResults(result) {
  const totals = result.rows.reduce((summary, row) => ({
    asignacion: summary.asignacion + row.asignacion,
    adicional: summary.adicional + row.adicional,
    total: summary.total + row.total,
    consumo: summary.consumo + row.consumo,
    saldo: summary.saldo + row.saldo,
  }), { asignacion: 0, adicional: 0, total: 0, consumo: 0, saldo: 0 });
  document.querySelector('#combustible-technician-count').textContent = formatNumber(result.rows.length);
  document.querySelector('#combustible-total-amount').textContent = formatCurrency(totals.total);
  document.querySelector('#combustible-balance-amount').textContent = formatCurrency(totals.saldo);
  document.querySelector('#combustible-source-name').textContent = result.sourceName;
  document.querySelector('#combustible-result-status').textContent = result.warning || 'Variables procesadas';
  const table = document.querySelector('#combustible-results-table');
  table.innerHTML = result.rows.length
    ? `<table><thead><tr><th>RUT</th><th>Nombre</th><th>Asignación</th><th>Adicional</th><th>Total</th><th>Consumo</th><th>Saldo</th></tr></thead><tbody>${result.rows.map((row) => `<tr><td>${escapeHtml(row.rut)}</td><td>${escapeHtml(row.nombre)}</td><td>${formatCurrency(row.asignacion)}</td><td>${formatCurrency(row.adicional)}</td><td>${formatCurrency(row.total)}</td><td>${formatCurrency(row.consumo)}</td><td class="${row.saldo < 0 ? 'negative' : 'positive'}">${formatCurrency(row.saldo)}</td></tr>`).join('')}</tbody></table>`
    : `<div class="results-empty"><strong>No fue posible calcular Combustible/TAG.</strong><span>${escapeHtml(result.warning)}</span></div>`;
}

function buildConcursoHheeResult() {
  const pago = state.sources.pago?.report;
  const novedades = state.sources.novedades?.report;
  const variables = pago?.sheets.find((sheet) => normalize(sheet.name) === 'VARIABLES');
  const hhee = novedades?.sheets.find((sheet) => normalize(sheet.name) === 'HHEE');
  const hheeContext = buildHheeCalculationContext(variables);
  const concursoRows = (variables?.records ?? []).map((record) => {
    const rut = String(pickColumn(record, ['RUT TAC', 'RUT'])).trim();
    if (!rut) return null;
    const concurso = parseAmount(pickColumn(record, ['Concurso TOA', 'Concurso']));
    const totalConcurso = parseAmount(pickColumn(record, ['Total concurso']));
    if (concurso === 0 && totalConcurso === 0) return null;
    return { rut, nombre: String(pickColumn(record, ['Nombre TAC', 'Nombre', 'NOMBRE'])).trim(), concurso, totalConcurso };
  }).filter(Boolean);
  const hheeRows = (hhee?.records ?? []).map((record) => {
    const firstNameColumn = String(record.NOMBRE ?? '').trim();
    const secondNameColumn = String(record.NOMBRE_2 ?? '').trim();
    const rutColumn = String(record.RUT ?? '').trim();
    const firstLooksLikeRut = isRutLike(firstNameColumn);
    const rut = firstLooksLikeRut ? firstNameColumn : (isRutLike(rutColumn) ? rutColumn : String(pickColumn(record, ['RUT TAC'])).trim());
    if (!rut) return null;
    const horas = parseAmount(pickColumn(record, ['HORAS', 'HHEE']));
    const rawAmount = pickColumn(record, ['Monto HHEE', 'Total HHEE', 'Monto', 'Total Hora']);
    const hasAmount = String(rawAmount).trim() !== '';
    const hheeAmount = calculateHheeAmount(rut, horas, hheeContext, { amount: parseAmount(rawAmount), hasAmount });
    return {
      rut,
      nombre: firstLooksLikeRut ? rutColumn : (secondNameColumn || firstNameColumn),
      horas,
      tipo: String(pickColumn(record, ['TIPO'])).trim(),
      area: String(pickColumn(record, ['AREA', 'ÁREA'])).trim(),
      estado: String(pickColumn(record, ['ESTADO'])).trim(),
      valorHoraExtra: hheeAmount.valorHoraExtra,
      monto: hheeAmount.monto,
      montoFuente: hheeAmount.fuente,
      montoInformado: hasAmount ? parseAmount(rawAmount) : 0,
      tieneMontoInformado: hasAmount,
      horasFuente: 'Novedades · HHEE',
    };
  }).filter(Boolean);
  return {
    concursoRows,
    hheeRows,
    concursoSourceName: variables?.name ?? 'Variables',
    hheeSourceName: hhee?.name ?? 'Novedades · HHEE',
  };
}

function isRutLike(value) {
  return /^\d{7,8}-[\dK]$/i.test(String(value ?? '').replace(/[.\s]/g, ''));
}

function buildHheeCalculationContext(variables) {
  const novedades = state.sources.novedades?.report;
  const pago = state.sources.pago?.report;
  const rex = state.sources.rex?.report;
  const salaryByRut = new Map();
  const variableByRut = new Map();
  const employeeSheets = [
    rex?.sheets.find((sheet) => normalize(sheet.name) === 'DETALLE'),
    novedades?.sheets.find((sheet) => normalize(sheet.name) === 'FUNCIONARIO'),
    pago?.sheets.find((sheet) => normalize(sheet.name) === 'NOMINA RRHH'),
  ].filter(Boolean);

  for (const sheet of employeeSheets) {
    for (const record of sheet.records ?? []) {
      const rut = String(pickColumn(record, ['RUT TAC', 'RUT'])).trim();
      const salary = parseAmount(pickColumn(record, ['VALOR SUELDO BASE', 'Sueldo Base', 'SUELDO BASE']));
      if (rut && salary > 0 && !salaryByRut.has(normalizeRutKey(rut))) salaryByRut.set(normalizeRutKey(rut), salary);
    }
  }

  for (const record of variables?.records ?? []) {
    const section = findHheeVariableSection(record);
    const rut = section?.rut || String(pickColumn(record, ['RUT TAC', 'RUT'])).trim();
    if (!rut) continue;
    const rawHours = pickColumn(record, ['Q HHEE', 'HORAS EXTRAS']);
    const rawAmount = pickColumn(record, ['Total Hora', 'Monto HHEE', 'Total HHEE']);
    if (!String(rawHours).trim() && !String(rawAmount).trim()) continue;
    variableByRut.set(normalizeRutKey(rut), {
      rut,
      nombre: section?.nombre || String(pickColumn(record, ['Nombre TAC', 'Nombre', 'NOMBRE'])).trim(),
      hours: parseAmount(rawHours),
      amount: parseAmount(rawAmount),
      hasAmount: String(rawAmount).trim() !== '',
    });
  }

  const fallbackSalary = salaryByRut.values().next().value || 0;
  const fallbackRate = fallbackSalary > 0 ? fallbackSalary / 180 * 1.5 : 0;
  return { salaryByRut, variableByRut, fallbackRate };
}

function findHheeVariableSection(record) {
  const entries = Object.entries(record);
  const hheeIndex = entries.findIndex(([key]) => normalize(key) === 'Q HHEE');
  if (hheeIndex < 0) return null;
  const rutEntry = entries.slice(0, hheeIndex).reverse().find(([key, value]) => isRutColumn(key) && isRutLike(value));
  if (!rutEntry) return null;
  const rutIndex = entries.findIndex(([key, value]) => key === rutEntry[0] && value === rutEntry[1]);
  const nameEntry = entries.slice(rutIndex + 1, hheeIndex).reverse().find(([key, value]) => normalize(key).startsWith('NOMBRE') && String(value).trim());
  return { rut: String(rutEntry[1]).trim(), nombre: String(nameEntry?.[1] ?? '').trim() };
}

function isRutColumn(value) {
  return /^RUT(?: TAC)?(?: \d+)?$/.test(normalize(value));
}

function calculateHheeAmount(rut, hours, context, sourceOverride = {}) {
  const key = normalizeRutKey(rut);
  const variable = context.variableByRut.get(key);
  const salary = context.salaryByRut.get(key) || 0;
  const sourceAmount = sourceOverride.hasAmount ? sourceOverride.amount : variable?.amount;
  const sourceHours = sourceOverride.hasAmount ? hours : variable?.hours;
  const sourceRate = sourceAmount > 0 && sourceHours > 0 ? sourceAmount / sourceHours : 0;
  const formulaRate = salary > 0 ? salary / 180 * 1.5 : context.fallbackRate;
  const valorHoraExtra = roundAmount(sourceRate || formulaRate);
  const usesSourceAmount = Boolean((sourceOverride.hasAmount || variable?.hasAmount) && (!sourceHours || sourceHours === hours));
  const monto = usesSourceAmount ? sourceAmount : (valorHoraExtra > 0 ? hours * valorHoraExtra : 0);
  return {
    valorHoraExtra,
    monto: roundAmount(monto),
    fuente: usesSourceAmount
      ? (sourceOverride.hasAmount ? 'Monto informado en HHEE' : 'Total Hora informado en Variables')
      : (valorHoraExtra > 0 ? 'Fórmula Pre_Liqu TAC' : 'Monto no disponible'),
  };
}

function renderConcursoHheeResults(result) {
  const concursoTotal = result.concursoRows.reduce((total, row) => total + row.totalConcurso, 0);
  const hheeHours = result.hheeRows.reduce((total, row) => total + row.horas, 0);
  const hheeAmount = result.hheeRows.reduce((total, row) => total + row.monto, 0);
  document.querySelector('#concurso-source-name').textContent = result.concursoSourceName;
  document.querySelector('#concurso-technician-count').textContent = formatNumber(result.concursoRows.length);
  document.querySelector('#concurso-total-amount').textContent = formatCurrency(concursoTotal);
  document.querySelector('#hhee-source-name').textContent = result.hheeSourceName;
  document.querySelector('#hhee-record-count').textContent = formatNumber(result.hheeRows.length);
  document.querySelector('#hhee-total-hours').textContent = String(roundAmount(hheeHours));
  document.querySelector('#hhee-total-amount').textContent = formatCurrency(hheeAmount);
  document.querySelector('#concurso-results-table').innerHTML = result.concursoRows.length
    ? `<table><thead><tr><th>RUT</th><th>Nombre</th><th>Concurso TOA</th><th>Total concurso</th></tr></thead><tbody>${result.concursoRows.map((row) => `<tr><td>${escapeHtml(row.rut)}</td><td>${escapeHtml(row.nombre)}</td><td>${formatCurrency(row.concurso)}</td><td>${formatCurrency(row.totalConcurso)}</td></tr>`).join('')}</tbody></table>`
    : '<div class="concept-empty">No hay valores de Concurso informados.</div>';
  document.querySelector('#hhee-results-table').innerHTML = result.hheeRows.length
    ? `<table><thead><tr><th>RUT</th><th>Nombre</th><th>Horas</th><th>Valor hora</th><th>Monto HHEE</th><th>Tipo</th></tr></thead><tbody>${result.hheeRows.map((row) => `<tr><td>${escapeHtml(row.rut)}</td><td>${escapeHtml(row.nombre)}</td><td>${row.horas}</td><td>${preliqMoney(row.valorHoraExtra)}</td><td>${preliqMoney(row.monto)}</td><td>${escapeHtml(row.tipo)}</td></tr>`).join('')}</tbody></table>`
    : '<div class="concept-empty">No hay HHEE informadas por RRHH.</div>';
}

function normalizeRutKey(value) {
  return String(value ?? '').replace(/[.\s]/g, '').toUpperCase();
}

function buildPreliquidationResult() {
  const pago = state.sources.pago?.report;
  const rex = state.sources.rex?.report?.sheets.find((sheet) => normalize(sheet.name) === 'DETALLE');
  const variables = pago?.sheets.find((sheet) => normalize(sheet.name) === 'VARIABLES');
  const libroRem = pago?.sheets.find((sheet) => normalize(sheet.name) === 'LIBRO REM');
  const hheeContext = buildHheeCalculationContext(variables);
  const concursoHhee = buildConcursoHheeResult();
  const bonusRecords = (state.sources.bono?.report?.sheets ?? []).flatMap((sheet) => sheet.records ?? []);
  const consolidated = new Map();

  function ensureRow(rut, nombre = '') {
    const key = normalizeRutKey(rut);
    if (!key) return null;
    if (!consolidated.has(key)) consolidated.set(key, {
      key,
      rut: String(rut).trim(),
      nombre: String(nombre ?? '').trim(),
      remuneracionesDesdeRex: false,
      remuneracionesFuente: '',
      rexEmpresa: '',
      rexProceso: '',
      rexSede: '',
      rexFechaInicio: '',
      rexFechaTermino: '',
      rexTipoContrato: '',
      rexCentroCosto: '',
      rexAgrupacion: '',
      rexHaberesExentos: 0,
      rexSumaHaberes: 0,
      rexTotalRebajas: 0,
      rexAfectoAfp: 0,
      rexAfectoCesantia: 0,
      rexAfectoImpuesto: 0,
      rexAlcanceLiquido: 0,
      rexAporteCaf: 0,
      supervisor: '',
      area: '',
      region: '',
      sede: '',
      centroCosto: '',
      tipoContrato: '',
      estado: '',
      sueldoBase: 0,
      diasOperando: 0,
      estadoFinal: '',
      cargo: '',
      antiguedadMeses: 0,
      diasNoOperando: 0,
      diasTrabajadosMes: 0,
      cumpleMeta: 0,
      rangoProductividad: 0,
      rangoCalidad: 0,
      llaveProductividadCalidad: '',
      capacidadOciosa: 0,
      totalMovil: 0,
      compensacion: 0,
      totalMovilCompensacion: 0,
      ingresoFullOperatividad: 0,
      porcentajeRendimiento: 0,
      diasJustificar: '',
      bonoProduccion100: 0,
      capacidadOciosaPx0: 0,
      permisoSinGoce: 0,
      otrosDesgastes: 0,
      maestroGuiaBonosExtra: 0,
      compensacionEmpresa: 0,
      diferenciaContratoZn: 0,
      totalOtrosDesgastes: 0,
      observacionVariables: '',
      compensacionIngresoAct: 0,
      alcanceLiquidoNew: 0,
      bonosTotales: 0,
      fechaDesvinculacion: '',
      coordenada: '',
      baseToa: 0,
      diferencia800: 0,
      difVariables: 0,
      bonoReferencia: 0,
      semanaCorrida: 0,
      vacacionesMonto: 0,
      hheeMontoLibro: 0,
      otrosDiferenciaSueldo: 0,
      aguinaldo: 0,
      totalIngresos: 0,
      gratificacion: 0,
      totalHaberesImponibles: 0,
      colacionEspecial: 0,
      desgasteHerramientasLibro: 0,
      viaticosLibro: 0,
      asignacionFamiliar: 0,
      totalHaberesNoImponibles: 0,
      totalHaberes: 0,
      saludTipo: '',
      montoPactadoSalud: 0,
      totalPactadoSalud: 0,
      totalSalud: 0,
      afpInstitucion: '',
      afp: 0,
      afc: 0,
      ahorroVoluntario: 0,
      totalDescuentosPrevisionales: 0,
      alcanceLiquido: 0,
      saludInstitucion: '',
      descuentoSaludSiete: 0,
      adicionalIsapre: 0,
      afpTasa: 0,
      afcTasa: 0,
      anticipoRemuneraciones: 0,
      prestamos: 0,
      cajaCompensacion: 0,
      saldoEficiencia: 0,
      prestamoSolidario: 0,
      anticipoAguinaldo: 0,
      retenciones: 0,
      prestamoFonasa: 0,
      anticipoViatico: 0,
      colectaFuncionario: 0,
      totalDescuentos: 0,
      sueldoLiquido: 0,
      diferencia: 0,
      alcanceSinBonos: 0,
      ingresoEmpresaPx: 0,
      adicionalEmpresa: 0,
      totalEmpresa: 0,
      costoRrhhTac: 0,
      costoFijo: 0,
      costoTotal: 0,
      saldoEmpresa: 0,
      margen: 0,
      estadoResultado: '',
      meses: 0,
      totalRgu: 0,
      totalActividades: 0,
      totalActividadesProduccion: 0,
      productionByFamily: [],
      productionRguExtensor: 0,
      productionRguIptv: 0,
      productionTotalBase: 0,
      totalActividadesMesAnterior: 0,
      productividadNominal: 0,
      metaRgu: 0,
      diasCierreEfectivo: 0,
      diasOperandoVariables: 0,
      diasPlanHabil: 0,
      porcentajeAsistencia: 0,
      porcentajeCierreActividad: 0,
      vacaciones: 0,
      licencias: 0,
      faltas: 0,
      reiterados: 0,
      calidad: 0,
      factorCalidad: 0,
      valor100Movil: 0,
      valorPagoAsistencia: 0,
      bono: 0,
      asignacionCombustibleTag: 0,
      adicionalCombustible: 0,
      totalCombustible: 0,
      consumo: 0,
      saldo: 0,
      movilizacionEspecial: 0,
      eficienciaConsumo: 0,
      totalFestivosTrabajados: 0,
      totalBonoFestivo: 0,
      incentivoAdicionalEspecial: 0,
      totalFeriados: 0,
      pxCeroDias: 0,
      incidenciaOperacionalNow: 0,
      compensacionBajoIngresoAct: 0,
      celular: 0,
      herramientasMenores: 0,
      otrosDesgastesHerramientas: 0,
      totalDesgasteHerramientas: 0,
      concursoBase: 0,
      concurso: 0,
      hheeHoras: 0,
      hheeHorasFuente: '',
      hheeValorHora: 0,
      hheeMonto: 0,
      hheeMontoFuente: '',
    });
    const row = consolidated.get(key);
    if (!row.nombre && nombre) row.nombre = String(nombre).trim();
    return row;
  }

  const rexColumns = new Set((rex?.header ?? []).map(normalize));
  const hasRexColumn = (names) => names.some((name) => rexColumns.has(normalize(name)));
  const readRexAmount = (record, names, fallback = 0) => {
    const value = pickColumn(record, names);
    return String(value).trim() === '' ? fallback : parseAmount(value);
  };
  const readRexText = (record, names, fallback = '') => {
    const value = pickColumn(record, names);
    return String(value).trim() || fallback;
  };

  for (const record of rex?.records ?? []) {
    const target = ensureRow(pickColumn(record, ['Rut', 'RUT']), pickColumn(record, ['Nombre', 'NOMBRE']));
    if (!target) continue;
    target.remuneracionesDesdeRex = true;
    target.remuneracionesFuente = 'Libro de remuneraciones REX+';
    target.rexEmpresa = readRexText(record, ['Nombre empresa']);
    target.rexProceso = readRexText(record, ['Proceso']);
    target.rexSede = readRexText(record, ['Sede']);
    target.sede = target.rexSede;
    target.rexFechaInicio = readRexText(record, ['Fecha Inicio']);
    target.rexFechaTermino = readRexText(record, ['Fecha Término']);
    target.rexTipoContrato = readRexText(record, ['Tipo Contrato']);
    target.rexCentroCosto = readRexText(record, ['Centro Costo']);
    target.rexAgrupacion = readRexText(record, ['Agrupación']);
    target.cargo = readRexText(record, ['Cargo'], target.cargo);
    target.sueldoBase = readRexAmount(record, ['Sueldo Base'], target.sueldoBase);
    target.gratificacion = readRexAmount(record, ['Gratificación'], target.gratificacion);
    target.diasTrabajadosMes = readRexAmount(record, ['Días Trabajados'], target.diasTrabajadosMes);
    target.licencias = readRexAmount(record, ['Dias con Licencia Medica'], target.licencias);
    target.asignacionFamiliar = readRexAmount(record, ['Cargas Familiares Simples'], target.asignacionFamiliar);
    target.afp = readRexAmount(record, ['Cotizacion AFP'], target.afp);
    target.totalSalud = readRexAmount(record, ['Cotizacion SALUD'], target.totalSalud);
    target.afc = readRexAmount(record, ['Seguro de Cesantia'], target.afc);
    target.ahorroVoluntario = readRexAmount(record, ['APVI Ahorro voluntario mensual'], target.ahorroVoluntario);
    target.anticipoRemuneraciones = readRexAmount(record, ['Anticipo'], target.anticipoRemuneraciones);
    target.prestamos = readRexAmount(record, ['Creditos personales CCAF'], target.prestamos);
    target.cajaCompensacion = readRexAmount(record, ['Ahorro en CCAF'], target.cajaCompensacion);
    target.totalHaberesNoImponibles = readRexAmount(record, ['Haberes Exentos'], target.totalHaberesNoImponibles);
    target.totalHaberes = readRexAmount(record, ['Suma Haberes'], target.totalHaberes);
    target.totalDescuentos = readRexAmount(record, ['Total Rebajas'], target.totalDescuentos);
    target.totalHaberesImponibles = readRexAmount(record, ['Total imponible sin tope'], target.totalHaberesImponibles);
    target.alcanceLiquido = readRexAmount(record, ['Alcance Líquido'], target.alcanceLiquido);
    target.sueldoLiquido = readRexAmount(record, ['Sueldo Líquido'], target.sueldoLiquido);
    target.saludInstitucion = readRexText(record, ['Inst. Salud'], target.saludInstitucion);
    target.afpInstitucion = readRexText(record, ['AFP'], target.afpInstitucion);
    target.rexHaberesExentos = target.totalHaberesNoImponibles;
    target.rexSumaHaberes = target.totalHaberes;
    target.rexTotalRebajas = target.totalDescuentos;
    target.rexAfectoAfp = readRexAmount(record, ['Afecto AFP']);
    target.rexAfectoCesantia = readRexAmount(record, ['Afecto Cesantía']);
    target.rexAfectoImpuesto = readRexAmount(record, ['Afecto Impuesto']);
    target.rexAlcanceLiquido = target.alcanceLiquido;
    target.rexAporteCaf = readRexAmount(record, ['Aporte a CCAF']);
    target.totalIngresos = target.rexSumaHaberes;
    target.totalDescuentosPrevisionales = roundAmount(target.afp + target.totalSalud + target.afc + target.ahorroVoluntario);
    if (hasRexColumn(['Tipo Contrato'])) target.tipoContrato = target.rexTipoContrato;
    if (hasRexColumn(['Centro Costo'])) target.centroCosto = target.rexCentroCosto;
    if (hasRexColumn(['Bono de Producción', 'Bono producción'])) target.bonoReferencia = readRexAmount(record, ['Bono de Producción', 'Bono producción']);
    if (hasRexColumn(['Horas Extra', 'HHEE', 'Monto HHEE'])) target.hheeMontoLibro = readRexAmount(record, ['Horas Extra', 'Monto HHEE', 'HHEE']);
    if (hasRexColumn(['Semana Corrida'])) target.semanaCorrida = readRexAmount(record, ['Semana Corrida']);
    if (hasRexColumn(['Vacaciones'])) target.vacacionesMonto = readRexAmount(record, ['Vacaciones']);
    if (hasRexColumn(['Aguinaldo'])) target.aguinaldo = readRexAmount(record, ['Aguinaldo']);
    if (hasRexColumn(['Movilización especial'])) target.movilizacionEspecial = readRexAmount(record, ['Movilización especial']);
    if (hasRexColumn(['Desgaste de herramientas'])) target.desgasteHerramientasLibro = readRexAmount(record, ['Desgaste de herramientas']);
    if (hasRexColumn(['Viáticos'])) target.viaticosLibro = readRexAmount(record, ['Viáticos']);
  }

  const consolidado = pago?.sheets.find((sheet) => normalize(sheet.name) === 'CONSOLIDADO');
  const productionByRut = new Map();
  for (const record of consolidado?.records ?? []) {
    const key = normalizeRutKey(pickColumn(record, ['RUT']));
    const family = String(pickColumn(record, ['Familia VAL', 'FAMILIA', 'Familia'])).trim();
    if (!key || !family) continue;
    const current = productionByRut.get(key) || {
      families: new Map(),
      rguExtensor: 0,
      rguIptv: 0,
      totalBase: 0,
      supervisor: '',
    };
    const familyTotals = current.families.get(family) || { activities: 0, rgu: 0 };
    familyTotals.activities += parseAmount(pickColumn(record, ['Q Actividades']));
    familyTotals.rgu += parseAmount(pickColumn(record, ['Total RGU']));
    current.families.set(family, familyTotals);
    current.rguExtensor += parseAmount(pickColumn(record, ['RGU Extensor']));
    current.rguIptv += parseAmount(pickColumn(record, ['RGU IPTV']));
    current.totalBase += parseAmount(pickColumn(record, ['Total Base']));
    current.supervisor ||= String(pickColumn(record, ['NOMBRE CORTO SUPERVISOR', 'NOMBRE SUPERVISOR'])).trim();
    productionByRut.set(key, current);
  }

  for (const record of variables?.records ?? []) {
    const target = ensureRow(pickColumn(record, ['RUT TAC', 'RUT']), pickColumn(record, ['Nombre TAC', 'Nombre', 'NOMBRE']));
    if (!target) continue;
    target.sueldoBase ||= hheeContext.salaryByRut.get(target.key) || 0;
    target.supervisor ||= String(pickColumn(record, ['Supervisor', 'SUPERVISOR'])).trim();
    const variableArea = String(pickColumn(record, ['Area', 'Área'])).trim();
    const variableRegion = String(pickColumn(record, ['REGION', 'Región', 'Region'])).trim();
    if (variableArea) target.area = variableArea;
    if (variableRegion) target.region = variableRegion;
    target.estado ||= String(pickColumn(record, ['Estado', 'ESTADO'])).trim();
    target.antiguedadMeses = parseAmount(pickColumn(record, ['Aniguedad Meses', 'Antigüedad Meses']));
    target.totalRgu = parseAmount(pickColumn(record, ['Total RGU']));
    target.totalActividades = parseAmount(pickColumn(record, ['Total Actividades', 'Total Actividades (N-1)']));
    target.totalActividadesMesAnterior = parseAmount(pickColumn(record, ['Total Actividades (N-1)']));
    target.productividadNominal = parseAmount(pickColumn(record, ['Productividad Nominal']));
    target.metaRgu = parseAmount(pickColumn(record, ['Meta RGU']));
    target.cumpleMeta = parseAmount(pickColumn(record, ['Cumple Meta %']));
    target.diasCierreEfectivo = parseAmount(pickColumn(record, ['Total días cierre efectivo']));
    target.diasOperandoVariables = parseAmount(pickColumn(record, ['Total días operando (1)']));
    target.diasNoOperando = parseAmount(pickColumn(record, ['Total días no operando (2)']));
    target.diasTrabajadosMes = parseAmount(pickColumn(record, ['Días Trabajados (Mes 30 días)', 'Dias trabajados']));
    target.diasPlanHabil = parseAmount(pickColumn(record, ['Total dias (1+2)', 'Días - Capacidad Ociosa']));
    target.porcentajeAsistencia = parseAmount(pickColumn(record, ['% Asistencia']));
    target.vacaciones = parseAmount(pickColumn(record, ['Vacaciones']));
    target.licencias = parseAmount(pickColumn(record, ['Incidencia - Rnow', 'Licencia']));
    target.faltas = parseAmount(pickColumn(record, ['Faltas']));
    target.reiterados = parseAmount(pickColumn(record, ['Reiterados']));
    target.calidad = parseAmount(pickColumn(record, ['Calidad']));
    target.factorCalidad = parseAmount(pickColumn(record, ['Factor Calidad']));
    target.rangoProductividad = parseAmount(pickColumn(record, ['Rango Tabla Productividad']));
    target.rangoCalidad = parseAmount(pickColumn(record, ['Rango Tabla Calidad']));
    target.llaveProductividadCalidad = String(pickColumn(record, ['Llave'])).trim();
    target.valor100Movil = parseAmount(pickColumn(record, ['Valor 100% Movil']));
    target.valorPagoAsistencia = parseAmount(pickColumn(record, ['Valor pago * asistencia']));
    target.capacidadOciosa = parseAmount(pickColumn(record, ['Capacidad Ociosa']));
    target.totalMovil = parseAmount(pickColumn(record, ['Total movil']));
    target.compensacion = parseAmount(pickColumn(record, ['Compensación']));
    target.totalMovilCompensacion = parseAmount(pickColumn(record, ['Total movil + compe']));
    target.ingresoFullOperatividad = parseAmount(pickColumn(record, ['Ingreso Full Operatividad']));
    target.porcentajeRendimiento = parseAmount(pickColumn(record, ['% Rendimiento']));
    target.diasJustificar = String(pickColumn(record, ['Dias a justificar'])).trim();
    target.bonoProduccion100 = parseAmount(pickColumn(record, ['Bono producción (100%)']));
    target.capacidadOciosaPx0 = parseAmount(pickColumn(record, ['Capacidad Ociosa (Px=0)']));
    target.asignacionCombustibleTag = parseAmount(pickColumn(record, ['Asignacion de combustible y TAG', 'Asignación de combustible y TAG']));
    target.adicionalCombustible = parseAmount(pickColumn(record, ['Adicional', 'Adicional combustible']));
    target.consumo = parseAmount(pickColumn(record, ['Consumo']));
    target.totalCombustible = roundAmount(target.asignacionCombustibleTag + target.adicionalCombustible);
    target.saldo = roundAmount(target.totalCombustible - target.consumo);
    target.movilizacionEspecial = target.saldo > 0 ? target.saldo : 0;
    target.eficienciaConsumo = target.saldo < 0 ? target.saldo : 0;
    target.concursoBase = parseAmount(pickColumn(record, ['Concurso TOA', 'Concurso']));
    target.concurso = parseAmount(pickColumn(record, ['Total concurso'])) || target.concursoBase;
    target.permisoSinGoce = parseAmount(pickColumn(record, ['Permiso sin gose', 'Permiso sin goce']));
    target.otrosDesgastes = parseAmount(pickColumn(record, ['Otros desgastes']));
    target.maestroGuiaBonosExtra = parseAmount(pickColumn(record, ['Maestro Guia + Bonos extra']));
    target.compensacionEmpresa = parseAmount(pickColumn(record, ['Compensacion empresa']));
    target.diferenciaContratoZn = parseAmount(pickColumn(record, ['Diferencia Contrato ZN']));
    target.totalOtrosDesgastes = parseAmount(pickColumn(record, ['Total Otros Desgastes']));
    target.observacionVariables = String(pickColumn(record, ['Observación', 'Observacion'])).trim();
    target.compensacionIngresoAct = parseAmount(pickColumn(record, ['Compensación x ingreso act', 'Compensacion x ingreso act']));
    target.alcanceLiquidoNew = parseAmount(pickColumn(record, ['Alcance Liquido New']));
    target.bonosTotales = parseAmount(pickColumn(record, ['Bonos Totales']));
    target.fechaDesvinculacion = String(pickColumn(record, ['Fecha desvinculacion', 'Fecha desvinculado'])).trim();
    target.coordenada = String(pickColumn(record, ['Coordenada'])).trim();
    target.baseToa = parseAmount(pickColumn(record, ['BASE TOA']));
    target.diferencia800 = parseAmount(pickColumn(record, ['Difrencia 800', 'Diferencia 800']));
    target.difVariables = parseAmount(pickColumn(record, ['Dif']));
    target.totalFestivosTrabajados = parseAmount(pickColumn(record, ['Total Festivos trabajados']));
    target.totalBonoFestivo = parseAmount(pickColumn(record, ['Total bono festivo']));
    target.incentivoAdicionalEspecial = parseAmount(pickColumn(record, ['Incentivo adicional especial']));
    target.totalFeriados = parseAmount(pickColumn(record, ['Total']));
    target.pxCeroDias = parseAmount(pickColumn(record, ['Total días Px=0']));
    target.incidenciaOperacionalNow = parseAmount(pickColumn(record, ['Total días Insidencia NOW', 'Total días Incidencia NOW']));
    target.compensacionBajoIngresoAct = parseAmount(pickColumn(record, ['Compensación bajo ingreso Act']));
    target.celular = roundAmount(target.diasTrabajadosMes * (1000 / 3));
    target.herramientasMenores = roundAmount(target.diasTrabajadosMes * (1000 / 3));
    target.totalDesgasteHerramientas = roundAmount(target.celular + target.herramientasMenores + target.totalOtrosDesgastes);
  }

  const familyOrder = ['Instalación', 'Visita Técnica', 'Cambio de domicilio', 'Upgrade', 'Downgrade'];
  for (const [key, detail] of productionByRut) {
    const target = consolidated.get(key);
    if (!target) continue;
    target.productionByFamily = familyOrder.map((family) => ({ family, ...detail.families.get(family) })).map((item) => ({
      family: item.family,
      activities: item.activities || 0,
      rgu: item.rgu || 0,
    }));
    target.totalActividadesProduccion = roundAmount(target.productionByFamily.reduce((total, item) => total + item.activities, 0));
    target.productionRguExtensor = roundAmount(detail.rguExtensor);
    target.productionRguIptv = roundAmount(detail.rguIptv);
    target.productionTotalBase = roundAmount(detail.totalBase);
    target.supervisor ||= detail.supervisor;
    target.porcentajeCierreActividad = target.diasPlanHabil > 0
      ? target.diasCierreEfectivo / target.diasPlanHabil
      : target.porcentajeAsistencia;
  }

  for (const record of libroRem?.records ?? []) {
    const rut = String(pickColumn(record, ['RUT', 'RUT TAC'])).trim();
    const target = ensureRow(rut, pickColumn(record, ['Nombre TAC', 'Nombre']));
    if (!target) continue;
    target.diasOperando = parseAmount(pickColumn(record, ['Dias Operando']));
    target.estadoFinal = String(pickColumn(record, ['Estado Final de mes'])).trim();
    const libroCargo = String(pickColumn(record, ['Cargo'])).trim();
    if (!target.cargo || !target.remuneracionesDesdeRex) target.cargo = libroCargo;
    const hasRexPayroll = target.remuneracionesDesdeRex;
    if (!hasRexPayroll) {
      target.sueldoBase ||= parseAmount(pickColumn(record, ['Sueldo Base']));
      target.gratificacion = parseAmount(pickColumn(record, ['Gratificación']));
      target.totalIngresos = parseAmount(pickColumn(record, ['Total Ingresos']));
      target.totalHaberesImponibles = parseAmount(pickColumn(record, ['Total Haberes Imponibles']));
      target.colacionEspecial = parseAmount(pickColumn(record, ['Colación especial']));
      target.movilizacionEspecial = parseAmount(pickColumn(record, ['Movilización especial']));
      target.desgasteHerramientasLibro = parseAmount(pickColumn(record, ['Desgaste de herramientas']));
      target.viaticosLibro = parseAmount(pickColumn(record, ['Viáticos']));
      target.asignacionFamiliar = parseAmount(pickColumn(record, ['Asignación Familiar']));
      target.totalHaberesNoImponibles = parseAmount(pickColumn(record, ['Total Haberes no imponibles']));
      target.totalHaberes = parseAmount(pickColumn(record, ['Total Haberes']));
      target.saludTipo = String(pickColumn(record, ['Tipo descuento salud'])).trim();
      target.montoPactadoSalud = parseAmount(pickColumn(record, ['Monto pactado']));
      target.totalPactadoSalud = parseAmount(pickColumn(record, ['Total Pactado']));
      target.totalSalud = parseAmount(pickColumn(record, ['Total Salud']));
      target.afpInstitucion = String(pickColumn(record, ['AFP Ins'])).trim();
      target.afp = parseAmount(pickColumn(record, ['AFP']));
      target.afc = parseAmount(pickColumn(record, ['AFC']));
      target.ahorroVoluntario = parseAmount(pickColumn(record, ['Ahorro voluntario']));
      target.totalDescuentosPrevisionales = parseAmount(pickColumn(record, ['Total descuentos previsionales']));
      target.alcanceLiquido = parseAmount(pickColumn(record, ['Alcance Liquido']));
      target.anticipoRemuneraciones = parseAmount(pickColumn(record, ['Anticipo remuneraciones']));
      target.prestamos = parseAmount(pickColumn(record, ['Prestamos']));
      target.cajaCompensacion = parseAmount(pickColumn(record, ['Caja de compensación']));
      target.saldoEficiencia = parseAmount(pickColumn(record, ['Saldo eficiencia de consumo']));
      target.prestamoSolidario = parseAmount(pickColumn(record, ['Prestamo solidario (3%)']));
      target.anticipoAguinaldo = parseAmount(pickColumn(record, ['Anticipo Aguinaldo']));
      target.retenciones = parseAmount(pickColumn(record, ['Retenciones']));
      target.prestamoFonasa = parseAmount(pickColumn(record, ['Prestamo Fonza', 'Prestamo Fonasa']));
      target.anticipoViatico = parseAmount(pickColumn(record, ['Anticipo Viatico']));
      target.totalDescuentos = parseAmount(pickColumn(record, ['Total Descuentos']));
      target.sueldoLiquido = parseAmount(pickColumn(record, ['Sueldo liquido']));
    }
    if (!hasRexColumn(['Bono de Producción', 'Bono producción'])) target.bonoReferencia = parseAmount(pickColumn(record, ['Bono de Producción']));
    if (!hasRexColumn(['Semana Corrida'])) target.semanaCorrida = parseAmount(pickColumn(record, ['Semana Corrida']));
    if (!hasRexColumn(['Vacaciones'])) target.vacacionesMonto = parseAmount(pickColumn(record, ['Vacaciones']));
    if (!hasRexColumn(['Horas Extra', 'HHEE', 'Monto HHEE'])) target.hheeMontoLibro = parseAmount(pickColumn(record, ['Horas Extra']));
    if (!hasRexColumn(['Otros (Diferencia sueldo)'])) target.otrosDiferenciaSueldo = parseAmount(pickColumn(record, ['Otros (Diferencia sueldo)']));
    if (!hasRexColumn(['Aguinaldo'])) target.aguinaldo = parseAmount(pickColumn(record, ['Aguinaldo']));
    target.diferencia = parseAmount(pickColumn(record, ['Diferencia']));
    target.alcanceSinBonos = parseAmount(pickColumn(record, ['Alcance sin bonos']));
    target.ingresoEmpresaPx = parseAmount(pickColumn(record, ['Ingreso Empresa (Px)']));
    target.adicionalEmpresa = parseAmount(pickColumn(record, ['Adicional']));
    target.totalEmpresa = parseAmount(pickColumn(record, ['Total']));
    target.costoRrhhTac = parseAmount(pickColumn(record, ['Costo RRHH TAC']));
    target.costoFijo = parseAmount(pickColumn(record, ['Costo Fijo']));
    target.costoTotal = parseAmount(pickColumn(record, ['Costo Total']));
    target.saldoEmpresa = parseAmount(pickColumn(record, ['Saldo']));
    target.margen = parseAmount(pickColumn(record, ['Margen']));
    target.estadoResultado = String(pickColumn(record, ['Estado'])).trim();
    target.meses = parseAmount(pickColumn(record, ['Meses']));
    target.descuentoSaludSiete = target.totalSalud;
    target.afpTasa = target.totalHaberesImponibles ? target.afp / target.totalHaberesImponibles : 0;
    target.afcTasa = target.totalHaberesImponibles ? target.afc / target.totalHaberesImponibles : 0;
    target.totalDesgasteHerramientas ||= target.desgasteHerramientasLibro;
  }

  const funcionario = state.sources.novedades?.report?.sheets.find((sheet) => normalize(sheet.name) === 'FUNCIONARIO');
  for (const record of funcionario?.records ?? []) {
    const rut = pickColumn(record, ['RUT']);
    const target = consolidated.get(normalizeRutKey(rut));
    if (!target) continue;
    target.nombre ||= String(pickColumn(record, ['NOMBRE'])).trim();
    target.sueldoBase ||= parseAmount(pickColumn(record, ['VALOR SUELDO BASE']));
    target.cargo ||= String(pickColumn(record, ['CARGO'])).trim();
    target.saludInstitucion ||= String(pickColumn(record, ['ISAPRE'])).trim();
    target.saludTipo ||= String(pickColumn(record, ['TIPO PACTO ISAPRE'])).trim();
    target.montoPactadoSalud ||= parseAmount(pickColumn(record, ['MONTO PACTADO']));
    target.afpInstitucion ||= String(pickColumn(record, ['PREVISION'])).trim();
  }

  const hheeByRut = new Map();
  for (const row of concursoHhee.hheeRows) {
    const key = normalizeRutKey(row.rut);
    if (!key) continue;
    const current = hheeByRut.get(key) || { ...row, horas: 0, montoInformado: 0, tieneMontoInformado: true };
    current.horas += row.horas;
    current.montoInformado += row.montoInformado;
    current.tieneMontoInformado = current.tieneMontoInformado && row.tieneMontoInformado;
    hheeByRut.set(key, current);
  }
  for (const variable of hheeContext.variableByRut.values()) {
    const key = normalizeRutKey(variable.rut);
    if (!key || variable.hours <= 0 || hheeByRut.has(key)) continue;
    hheeByRut.set(key, {
      rut: variable.rut,
      nombre: variable.nombre,
      horas: variable.hours,
      montoInformado: 0,
      tieneMontoInformado: false,
      horasFuente: 'Pago TAC · Variables (respaldo)',
    });
  }
  for (const row of hheeByRut.values()) {
    const target = ensureRow(row.rut, row.nombre);
    if (target) {
      const hheeAmount = calculateHheeAmount(row.rut, row.horas, hheeContext, { amount: row.montoInformado, hasAmount: row.tieneMontoInformado });
      target.sueldoBase ||= hheeContext.salaryByRut.get(target.key) || 0;
      target.hheeHoras += row.horas;
      target.hheeValorHora = hheeAmount.valorHoraExtra;
      target.hheeMonto += hheeAmount.monto;
      target.hheeMontoFuente = hheeAmount.fuente;
      target.hheeHorasFuente = row.horasFuente || target.hheeHorasFuente;
    }
  }
  for (const record of bonusRecords) {
    const rut = pickColumn(record, ['RUT', 'RUT TAC']);
    const target = ensureRow(rut, pickColumn(record, ['NOMBRE', 'Nombre', 'Nombre TAC']));
    if (target) target.bono += parseAmount(pickColumn(record, ['MONTO', 'Monto']));
  }

  for (const row of consolidated.values()) {
    if (!row.bono && row.bonoReferencia) row.bono = row.bonoReferencia;
    if (!row.hheeMonto && row.hheeMontoLibro) {
      row.hheeMonto = row.hheeMontoLibro;
      row.hheeMontoFuente = 'Libro Rem';
    }
  }

  const rows = [...consolidated.values()].map((row) => {
    if (!row.area) row.area = row.sede || row.centroCosto;
    const totalMonetario = row.bono + row.movilizacionEspecial + row.eficienciaConsumo + row.concurso + row.hheeMonto;
    return {
      ...row,
      bono: roundAmount(row.bono),
      asignacionCombustibleTag: roundAmount(row.asignacionCombustibleTag),
      adicionalCombustible: roundAmount(row.adicionalCombustible),
      totalCombustible: roundAmount(row.totalCombustible || row.asignacionCombustibleTag + row.adicionalCombustible),
      consumo: roundAmount(row.consumo),
      saldo: roundAmount(row.saldo || row.movilizacionEspecial + row.eficienciaConsumo),
      movilizacionEspecial: roundAmount(row.movilizacionEspecial),
      eficienciaConsumo: roundAmount(row.eficienciaConsumo),
      concursoBase: roundAmount(row.concursoBase),
      concurso: roundAmount(row.concurso),
      hheeHoras: roundAmount(row.hheeHoras),
      hheeValorHora: roundAmount(row.hheeValorHora),
      hheeMonto: roundAmount(row.hheeMonto),
      totalMonetario: roundAmount(totalMonetario),
    };
  }).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
  return { rows, period: document.querySelector('#period').value };
}

function preliqMoney(value) {
  return value ? formatCurrency(value) : '-';
}

function preliqNumber(value) {
  return value ? formatNumber(value) : '-';
}

function preliqPercent(value) {
  if (!value) return '-';
  const percent = value <= 2 ? value * 100 : value;
  return `${percent.toFixed(2).replace('.', ',')}%`;
}

function renderPreliqBlock(title, rows) {
  return `<section class="preliq-block"><h3>${title}</h3><div class="preliq-lines">${rows.map(([label, value]) => `<div><span>${label}</span><strong>${escapeHtml(value)}</strong></div>`).join('')}</div></section>`;
}

function renderPreliqTable(headers, rows, className = '') {
  return `<table class="preliq-sheet-table ${className}"><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead><tbody>${rows.map((row) => `<tr>${row.map((value) => `<td>${escapeHtml(value)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
}

function renderPreliqSheetPanel(title, rows, className = '') {
  return `<section class="preliq-sheet-panel ${className}"><h3>${escapeHtml(title)}</h3>${renderPreliqTable(['Descripción', 'Valor'], rows, 'preliq-key-value')}</section>`;
}

function renderPreliquidationSheet(row) {
  const activityRows = row.productionByFamily?.length
    ? row.productionByFamily.map((item) => [item.family, preliqNumber(item.activities), preliqNumber(item.rgu)])
    : [['Total', preliqNumber(row.totalActividades), preliqNumber(row.totalRgu)]];
  activityRows.push(['Total', preliqNumber(row.totalActividadesProduccion || row.totalActividades), preliqNumber(row.totalRgu)]);
  const activitySummaryRows = [
    ['Suma de RGU Extensor', preliqNumber(row.productionRguExtensor)],
    ['Suma de RGU IPTV', preliqNumber(row.productionRguIptv)],
    ['Suma de Total Base', preliqNumber(row.productionTotalBase)],
    ['Suma de Total RGU', preliqNumber(row.totalRgu)],
  ];
  const planningRows = [
    ['Total días hábiles', preliqNumber(row.diasPlanHabil), preliqNumber(row.diasCierreEfectivo)],
    ['Productividad', preliqNumber(row.metaRgu), preliqNumber(row.productividadNominal)],
    ['Calidad', preliqNumber(row.factorCalidad), preliqPercent(row.factorCalidad)],
  ];
  const cierrePercent = row.porcentajeCierreActividad || row.porcentajeAsistencia;
  const period = document.querySelector('#period').value || 'sin definir';
  return `<article class="preliq-sheet">
    <header class="preliq-sheet-header">
      <div class="preliq-sheet-brand"><img src="assets/now-logo.png" alt="NOW"><div><strong>NOW</strong><span>Remuneraciones</span></div></div>
      <div class="preliq-sheet-heading"><span>Pago TAC</span><strong>PRELIQUIDACIÓN</strong></div>
      <div class="preliq-sheet-period"><span>Período</span><strong>${escapeHtml(period)}</strong></div>
    </header>
    <div class="preliq-sheet-identity">
      <div><span>RUT</span><strong>${escapeHtml(row.rut || '-')}</strong></div>
      <div><span>Nombre TAC</span><strong>${escapeHtml(row.nombre || '-')}</strong></div>
      <div><span>Supervisor</span><strong>${escapeHtml(row.supervisor || '-')}</strong></div>
      <div><span>Cargo</span><strong>${escapeHtml(row.cargo || '-')}</strong></div>
      <div><span>Área / región</span><strong>${escapeHtml([row.area || row.sede, row.region || row.centroCosto].filter(Boolean).join(' / ') || '-')}</strong></div>
      <div><span>Estado</span><strong>${escapeHtml(row.estadoFinal || row.estado || '-')}</strong></div>
    </div>
    ${row.remuneracionesDesdeRex ? renderPreliqSheetPanel('REMUNERACIONES REX+', [
      ['Fuente', row.remuneracionesFuente],
      ['Empresa', row.rexEmpresa || '-'],
      ['Proceso', row.rexProceso || period],
      ['Sede', row.rexSede || row.sede || '-'],
      ['Tipo contrato', row.rexTipoContrato || row.tipoContrato || '-'],
      ['Centro costo', row.rexCentroCosto || row.centroCosto || '-'],
      ['Días trabajados', preliqNumber(row.diasTrabajadosMes)],
      ['Días con licencia médica', preliqNumber(row.licencias)],
      ['Suma haberes informada', preliqMoney(row.rexSumaHaberes)],
      ['Aporte CCAF empleador', preliqMoney(row.rexAporteCaf)],
      ['Total rebajas informado', preliqMoney(row.rexTotalRebajas)],
      ['Sueldo líquido informado', preliqMoney(row.sueldoLiquido)],
    ], 'preliq-panel-wide preliq-panel-source') : ''}
    <div class="preliq-sheet-grid">
      <section class="preliq-sheet-panel preliq-panel-wide"><h3>PRODUCCIÓN</h3><div class="preliq-production-tables">${renderPreliqTable(['Descripción', 'Q actividades', 'Total RGU'], activityRows)}${renderPreliqTable(['Descripción', 'RGU por actividad'], activitySummaryRows)}</div>${renderPreliqTable(['Descripción', 'Planificación', 'Resultado'], planningRows)}${renderPreliqTable(['Descripción', 'Valor'], [
        ['Total RGU', preliqNumber(row.totalRgu)],
        ['Productividad', preliqNumber(row.productividadNominal)],
      ], 'preliq-key-value')}</section>
      ${renderPreliqSheetPanel('PRODUCTIVIDAD', [
        ['Rango tabla de cálculo', preliqNumber(row.rangoProductividad)],
        ['Cumplimiento de meta', preliqPercent(row.cumpleMeta)],
        ['Días cierre efectivo', preliqNumber(row.diasCierreEfectivo)],
        ['Días hábiles', preliqNumber(row.diasPlanHabil)],
        ['% días con cierre de actividad', preliqPercent(cierrePercent)],
        ['Valor 100%', preliqMoney(row.valor100Movil)],
        ['Valor pago por asistencia', preliqMoney(row.valorPagoAsistencia)],
        ['Capacidad ociosa', preliqMoney(row.capacidadOciosa)],
        ['Compensación', preliqMoney(row.compensacion)],
        ['Total móvil + compensación', preliqMoney(row.totalMovilCompensacion)],
        ['Ingreso full operatividad', preliqMoney(row.ingresoFullOperatividad)],
        ['% rendimiento', preliqPercent(row.porcentajeRendimiento)],
      ])}
      ${renderPreliqSheetPanel('CALIDAD', [
        ['Total actividades mes anterior', preliqNumber(row.totalActividadesMesAnterior || row.totalActividades)],
        ['Total reiterados', preliqNumber(row.reiterados)],
        ['Ajuste de calidad', preliqNumber(row.calidad)],
        ['Factor calidad', preliqPercent(row.factorCalidad)],
        ['Rango tabla de cálculo', preliqNumber(row.rangoCalidad)],
        ['Llave', row.llaveProductividadCalidad || '-'],
      ])}
      ${renderPreliqSheetPanel('OPERATIVIDAD MES', [
        ['Total días con cierre de actividad', preliqNumber(row.diasCierreEfectivo)],
        ['% días con cierre efectivo', preliqPercent(cierrePercent)],
        ['Días no operando', preliqNumber(row.diasNoOperando)],
        ['Licencia', preliqNumber(row.licencias)],
        ['Vacaciones', preliqNumber(row.vacaciones)],
        ['Permiso sin goce', preliqNumber(row.permisoSinGoce)],
        ['Faltas', preliqNumber(row.faltas)],
        ['Total días operativos', preliqNumber(row.diasOperandoVariables || row.diasCierreEfectivo || row.diasOperando)],
        ['Horas extra', preliqNumber(row.hheeHoras)],
      ])}
      ${renderPreliqSheetPanel('BONO VARIABLE (Px)', [
        ['Coordenada', row.coordenada || '-'],
        ['Valor bono al 100%', preliqMoney(row.bonoProduccion100 || row.bonoReferencia)],
        ['% días con cierre de actividad', preliqPercent(cierrePercent)],
        ['Capacidad ociosa Px=0', preliqMoney(row.capacidadOciosaPx0)],
      ])}
      ${renderPreliqSheetPanel('BONO VARIABLE (OTROS)', [
        ['Total días Px=0', preliqNumber(row.pxCeroDias)],
        ['Total días incidencia operacional NOW', preliqNumber(row.incidenciaOperacionalNow)],
        ['Compensación bajo ingreso actividad', preliqMoney(row.compensacionBajoIngresoAct)],
        ['Bonos totales', preliqMoney(row.bonosTotales)],
      ])}
      ${renderPreliqSheetPanel('EFICIENCIA DE CONSUMO', [
        ['Asignación combustible y TAG', preliqMoney(row.asignacionCombustibleTag)],
        ['Asignación adicional', preliqMoney(row.adicionalCombustible)],
        ['Total asignación', preliqMoney(row.totalCombustible)],
        ['Consumo TAG / combustible', preliqMoney(row.consumo)],
        ['Saldo', preliqMoney(row.saldo)],
        ['Movilización especial', preliqMoney(row.movilizacionEspecial)],
        ['Eficiencia de consumo', preliqMoney(row.eficienciaConsumo)],
      ])}
      ${renderPreliqSheetPanel('DESGASTE DE HERRAMIENTAS', [
        ['Celular', preliqMoney(row.celular)],
        ['Herramientas menores', preliqMoney(row.herramientasMenores)],
        ['Otros', preliqMoney(row.otrosDesgastesHerramientas || row.totalOtrosDesgastes)],
        ['Total', preliqMoney(row.totalDesgasteHerramientas || row.desgasteHerramientasLibro)],
      ])}
      ${renderPreliqSheetPanel('OTROS', [
        ['Concurso TOA', preliqMoney(row.concursoBase)],
        ['Total concurso', preliqMoney(row.concurso)],
        ['BASE TOA', preliqMoney(row.baseToa)],
        ['Diferencia 800', preliqMoney(row.diferencia800)],
        ['Maestro guía + bonos extra', preliqMoney(row.maestroGuiaBonosExtra)],
        ['Compensación empresa', preliqMoney(row.compensacionEmpresa)],
        ['Viáticos', preliqMoney(row.viaticosLibro)],
        ['Horas extra informadas', preliqNumber(row.hheeHoras)],
        ['Valor hora extra', preliqMoney(row.hheeValorHora)],
        ['Monto HHEE', preliqMoney(row.hheeMonto)],
        ['Origen monto HHEE', row.hheeMontoFuente || '-'],
      ])}
      ${renderPreliqSheetPanel('FESTIVOS Y FERIADOS', [
        ['Total festivos trabajados', preliqNumber(row.totalFestivosTrabajados)],
        ['Total bono festivo', preliqMoney(row.totalBonoFestivo)],
        ['Incentivo adicional especial', preliqMoney(row.incentivoAdicionalEspecial)],
        ['Total', preliqMoney(row.totalFeriados)],
      ])}
      ${renderPreliqSheetPanel('RESUMEN', [
        ['Sueldo base', preliqMoney(row.sueldoBase)],
        ['Bono de producción', preliqMoney(row.bono)],
        ['Semana corrida', preliqMoney(row.semanaCorrida)],
        ['Vacaciones', preliqMoney(row.vacacionesMonto)],
        ['Aguinaldo', preliqMoney(row.aguinaldo)],
        ['Horas extra', preliqMoney(row.hheeMonto)],
        ['Otros (diferencia sueldo)', preliqMoney(row.otrosDiferenciaSueldo)],
        ['Gratificación', preliqMoney(row.gratificacion)],
        ['Total haberes imponibles', preliqMoney(row.totalHaberesImponibles)],
        ['Total haberes no imponibles', preliqMoney(row.totalHaberesNoImponibles)],
        ['Total haberes', preliqMoney(row.totalHaberes)],
      ], 'preliq-panel-emphasis')}
      ${renderPreliqSheetPanel('SALUD Y PREVISIÓN', [
        ['Institución de salud', row.saludInstitucion || '-'],
        ['Tipo descuento salud', row.saludTipo || '-'],
        ['Monto pactado', preliqMoney(row.montoPactadoSalud)],
        ['Total pactado', preliqMoney(row.totalPactadoSalud)],
        ['Descuento salud', preliqMoney(row.descuentoSaludSiete || row.totalSalud)],
        ['Adicional Isapre', preliqMoney(row.adicionalIsapre)],
        ['AFP institución', row.afpInstitucion || '-'],
        ['Tasa AFP', preliqPercent(row.afpTasa)],
        ['Descuento AFP', preliqMoney(row.afp)],
        ['Tasa AFC', preliqPercent(row.afcTasa)],
        ['Descuento AFC', preliqMoney(row.afc)],
        ['Ahorro voluntario', preliqMoney(row.ahorroVoluntario)],
        ['Total descuentos previsionales', preliqMoney(row.totalDescuentosPrevisionales)],
      ])}
      ${renderPreliqSheetPanel('DESCUENTOS', [
        ['Alcance líquido', preliqMoney(row.alcanceLiquido)],
        ['Anticipo remuneraciones', preliqMoney(row.anticipoRemuneraciones)],
        ['Préstamos', preliqMoney(row.prestamos)],
        ['Otros (Caja compensación)', preliqMoney(row.cajaCompensacion)],
        ['Saldo eficiencia de consumo', preliqMoney(row.saldoEficiencia)],
        ['Préstamo solidario (3%)', preliqMoney(row.prestamoSolidario)],
        ['Anticipo viático', preliqMoney(row.anticipoViatico)],
        ['Retenciones', preliqMoney(row.retenciones)],
        ['Colecta funcionario', preliqMoney(row.colectaFuncionario)],
        ['Préstamo Fonasa', preliqMoney(row.prestamoFonasa)],
        ['Total descuentos', preliqMoney(row.totalDescuentos)],
        ['Líquido a pago', preliqMoney(row.sueldoLiquido)],
      ], 'preliq-panel-emphasis')}
      ${renderPreliqSheetPanel('COSTO Y CONTROL', [
        ['Diferencia', preliqMoney(row.diferencia)],
        ['Alcance sin bonos', preliqMoney(row.alcanceSinBonos)],
        ['Ingreso empresa (Px)', preliqMoney(row.ingresoEmpresaPx)],
        ['Adicional', preliqMoney(row.adicionalEmpresa)],
        ['Total ingreso empresa', preliqMoney(row.totalEmpresa)],
        ['Costo RRHH TAC', preliqMoney(row.costoRrhhTac)],
        ['Costo fijo', preliqMoney(row.costoFijo)],
        ['Costo total', preliqMoney(row.costoTotal)],
        ['Saldo', preliqMoney(row.saldoEmpresa)],
        ['Margen', preliqPercent(row.margen)],
        ['Estado', row.estadoResultado || row.estadoFinal || '-'],
        ['Meses', preliqNumber(row.meses)],
      ], 'preliq-panel-wide preliq-panel-control')}
    </div>
  </article>`;
}

function renderPreliquidationWorker(row) {
  if (!row) {
    document.querySelector('#preliquidation-rut').textContent = '-';
    document.querySelector('#preliquidation-name').textContent = '-';
    document.querySelector('#preliquidation-supervisor').textContent = '-';
    document.querySelector('#preliquidation-area').textContent = '-';
    document.querySelector('#print-preliquidation-button').disabled = true;
    document.querySelector('#preliquidation-detail').innerHTML = '<div class="results-empty"><strong>No hay un trabajador seleccionado.</strong></div>';
    return;
  }
  document.querySelector('#print-preliquidation-button').disabled = false;
  document.querySelector('#preliquidation-rut').textContent = row.rut || '-';
  document.querySelector('#preliquidation-name').textContent = row.nombre || '-';
  document.querySelector('#preliquidation-supervisor').textContent = row.supervisor || '-';
  document.querySelector('#preliquidation-area').textContent = [row.area, row.region].filter(Boolean).join(' / ') || '-';
  document.querySelector('#preliquidation-detail').innerHTML = renderPreliquidationSheet(row);
  return;
  document.querySelector('#preliquidation-detail').innerHTML = `<div class="preliq-columns">${renderPreliqBlock('PRODUCCIÓN', [
    ['Total RGU', preliqNumber(row.totalRgu)],
    ['Actividades disponibles', preliqNumber(row.totalActividades)],
    ['Productividad', preliqNumber(row.productividadNominal)],
    ['Factor calidad final', preliqPercent(row.factorCalidad)],
    ['Bono de producción', preliqMoney(row.bono)],
  ])}${renderPreliqBlock('PRODUCTIVIDAD', [
    ['Meta RGU', preliqNumber(row.metaRgu)],
    ['Cumplimiento de meta', preliqPercent(row.cumpleMeta)],
    ['Rango tabla productividad', preliqNumber(row.rangoProductividad)],
    ['Días cierre efectivo', preliqNumber(row.diasCierreEfectivo)],
    ['Días hábiles', preliqNumber(row.diasPlanHabil)],
    ['% asistencia', preliqPercent(row.porcentajeAsistencia)],
    ['Valor 100%', preliqMoney(row.valor100Movil)],
    ['Valor pago por asistencia', preliqMoney(row.valorPagoAsistencia)],
    ['Capacidad ociosa', preliqMoney(row.capacidadOciosa)],
    ['Total móvil', preliqMoney(row.totalMovil)],
    ['Compensación', preliqMoney(row.compensacion)],
    ['Total móvil + compensación', preliqMoney(row.totalMovilCompensacion)],
    ['Ingreso full operatividad', preliqMoney(row.ingresoFullOperatividad)],
    ['% rendimiento', preliqPercent(row.porcentajeRendimiento)],
    ['Días a justificar', row.diasJustificar || '-'],
  ])}${renderPreliqBlock('CALIDAD', [
    ['Actividades mes anterior', preliqNumber(row.totalActividadesMesAnterior || row.totalActividades)],
    ['Reiterados', preliqNumber(row.reiterados)],
    ['Ajuste de calidad', preliqNumber(row.calidad)],
    ['Factor calidad', preliqPercent(row.factorCalidad)],
    ['Rango tabla calidad', preliqNumber(row.rangoCalidad)],
    ['Llave productividad / calidad', row.llaveProductividadCalidad || '-'],
  ])}${renderPreliqBlock('OPERATIVIDAD MES', [
    ['Días cierre efectivo', preliqNumber(row.diasCierreEfectivo)],
    ['Días no operando', preliqNumber(row.diasNoOperando)],
    ['% días con cierre', preliqPercent(row.porcentajeAsistencia)],
    ['Días trabajados mes', preliqNumber(row.diasTrabajadosMes)],
    ['Vacaciones', preliqNumber(row.vacaciones)],
    ['Licencias', preliqNumber(row.licencias)],
    ['Permiso sin goce', preliqNumber(row.permisoSinGoce)],
    ['Faltas', preliqNumber(row.faltas)],
    ['HHEE informadas', preliqNumber(row.hheeHoras)],
  ])}${renderPreliqBlock('EFICIENCIA DE CONSUMO', [
    ['Asignación combustible y TAG', preliqMoney(row.asignacionCombustibleTag)],
    ['Asignación adicional', preliqMoney(row.adicionalCombustible)],
    ['Total asignación', preliqMoney(row.totalCombustible)],
    ['Consumo', preliqMoney(row.consumo)],
    ['Saldo', preliqMoney(row.saldo)],
    ['Movilización especial', preliqMoney(row.movilizacionEspecial)],
    ['Eficiencia de consumo', preliqMoney(row.eficienciaConsumo)],
  ])}${renderPreliqBlock('BONO VARIABLE', [
    ['Valor bono al 100%', preliqMoney(row.bonoProduccion100 || row.bonoReferencia)],
    ['Coordenada', row.coordenada || '-'],
    ['Capacidad ociosa Px=0', preliqMoney(row.capacidadOciosaPx0)],
    ['Compensación bajo ingreso act.', preliqMoney(row.compensacionBajoIngresoAct)],
    ['Bonos totales', preliqMoney(row.bonosTotales)],
  ])}${renderPreliqBlock('OTROS', [
    ['Concurso TOA', preliqMoney(row.concursoBase)],
    ['Total concurso', preliqMoney(row.concurso)],
    ['BASE TOA', preliqMoney(row.baseToa)],
    ['Diferencia 800', preliqMoney(row.diferencia800)],
    ['Otros desgastes', preliqMoney(row.otrosDesgastes)],
    ['Maestro guía + bonos extra', preliqMoney(row.maestroGuiaBonosExtra)],
    ['Compensación empresa', preliqMoney(row.compensacionEmpresa)],
    ['Diferencia contrato ZN', preliqMoney(row.diferenciaContratoZn)],
    ['Total otros desgastes', preliqMoney(row.totalOtrosDesgastes)],
    ['Alcance líquido nuevo', preliqMoney(row.alcanceLiquidoNew)],
    ['Observación', row.observacionVariables || '-'],
    ['Sueldo base usado', preliqMoney(row.sueldoBase)],
    ['HHEE informadas', preliqNumber(row.hheeHoras)],
    ['Fuente horas HHEE', row.hheeHorasFuente || '-'],
    ['Valor hora extra', preliqMoney(row.hheeValorHora)],
    ['Monto HHEE', preliqMoney(row.hheeMonto)],
    ['Origen monto HHEE', row.hheeMontoFuente || '-'],
  ])}${renderPreliqBlock('DESGASTE DE HERRAMIENTAS', [
    ['Celular', preliqMoney(row.celular)],
    ['Herramientas menores', preliqMoney(row.herramientasMenores)],
    ['Otros', preliqMoney(row.otrosDesgastesHerramientas)],
    ['Total', preliqMoney(row.totalDesgasteHerramientas || row.desgasteHerramientasLibro)],
  ])}${renderPreliqBlock('FESTIVOS Y FERIADOS', [
    ['Festivos trabajados', preliqNumber(row.totalFestivosTrabajados)],
    ['Bono festivo', preliqMoney(row.totalBonoFestivo)],
    ['Incentivo adicional especial', preliqMoney(row.incentivoAdicionalEspecial)],
    ['Total', preliqMoney(row.totalFeriados)],
  ])}${renderPreliqBlock('HABERES IMPONIBLES', [
    ['Sueldo base', preliqMoney(row.sueldoBase)],
    ['Bono de producción', preliqMoney(row.bono || row.bonoReferencia)],
    ['Semana corrida', preliqMoney(row.semanaCorrida)],
    ['Vacaciones', preliqMoney(row.vacacionesMonto)],
    ['Horas extra', preliqMoney(row.hheeMonto)],
    ['Otros (diferencia sueldo)', preliqMoney(row.otrosDiferenciaSueldo)],
    ['Aguinaldo', preliqMoney(row.aguinaldo)],
    ['Total ingresos', preliqMoney(row.totalIngresos)],
    ['Gratificación', preliqMoney(row.gratificacion)],
    ['Total haberes imponibles', preliqMoney(row.totalHaberesImponibles)],
  ])}${renderPreliqBlock('HABERES NO IMPONIBLES', [
    ['Colación especial', preliqMoney(row.colacionEspecial)],
    ['Movilización especial', preliqMoney(row.movilizacionEspecial)],
    ['Desgaste de herramientas', preliqMoney(row.desgasteHerramientasLibro)],
    ['Viáticos', preliqMoney(row.viaticosLibro)],
    ['Asignación familiar', preliqMoney(row.asignacionFamiliar)],
    ['Total haberes no imponibles', preliqMoney(row.totalHaberesNoImponibles)],
    ['Total haberes', preliqMoney(row.totalHaberes)],
  ])}${renderPreliqBlock('DESCUENTOS PREVISIONALES', [
    ['Institución de salud', row.saludInstitucion || '-'],
    ['Tipo descuento salud', row.saludTipo || '-'],
    ['Monto pactado', preliqMoney(row.montoPactadoSalud)],
    ['Total pactado', preliqMoney(row.totalPactadoSalud)],
    ['Total salud', preliqMoney(row.totalSalud)],
    ['Descuento salud 7%', preliqMoney(row.descuentoSaludSiete)],
    ['Adicional Isapre', preliqMoney(row.adicionalIsapre)],
    ['AFP institución', row.afpInstitucion || '-'],
    ['Tasa AFP', preliqPercent(row.afpTasa)],
    ['AFP', preliqMoney(row.afp)],
    ['Tasa AFC', preliqPercent(row.afcTasa)],
    ['AFC', preliqMoney(row.afc)],
    ['Ahorro voluntario', preliqMoney(row.ahorroVoluntario)],
    ['Total descuentos previsionales', preliqMoney(row.totalDescuentosPrevisionales)],
  ])}${renderPreliqBlock('DESCUENTOS Y LÍQUIDO', [
    ['Alcance líquido', preliqMoney(row.alcanceLiquido)],
    ['Anticipo remuneraciones', preliqMoney(row.anticipoRemuneraciones)],
    ['Préstamos', preliqMoney(row.prestamos)],
    ['Caja de compensación', preliqMoney(row.cajaCompensacion)],
    ['Saldo eficiencia de consumo', preliqMoney(row.saldoEficiencia)],
    ['Préstamo solidario (3%)', preliqMoney(row.prestamoSolidario)],
    ['Anticipo aguinaldo', preliqMoney(row.anticipoAguinaldo)],
    ['Retenciones', preliqMoney(row.retenciones)],
    ['Colecta funcionario', preliqMoney(row.colectaFuncionario)],
    ['Préstamo Fonasa', preliqMoney(row.prestamoFonasa)],
    ['Anticipo viático', preliqMoney(row.anticipoViatico)],
    ['Total descuentos', preliqMoney(row.totalDescuentos)],
    ['Sueldo líquido', preliqMoney(row.sueldoLiquido)],
  ])}${renderPreliqBlock('COSTO Y CONTROL', [
    ['Diferencia', preliqMoney(row.diferencia)],
    ['Alcance sin bonos', preliqMoney(row.alcanceSinBonos)],
    ['Ingreso empresa (Px)', preliqMoney(row.ingresoEmpresaPx)],
    ['Adicional', preliqMoney(row.adicionalEmpresa)],
    ['Total ingreso empresa', preliqMoney(row.totalEmpresa)],
    ['Costo RRHH TAC', preliqMoney(row.costoRrhhTac)],
    ['Costo fijo', preliqMoney(row.costoFijo)],
    ['Costo total', preliqMoney(row.costoTotal)],
    ['Saldo empresa', preliqMoney(row.saldoEmpresa)],
    ['Margen', preliqPercent(row.margen)],
    ['Estado', row.estadoResultado || row.estadoFinal || '-'],
    ['Meses', preliqNumber(row.meses)],
  ])}${renderPreliqBlock('RESUMEN', [
    ['Bono de producción', preliqMoney(row.bono)],
    ['Movilización especial', preliqMoney(row.movilizacionEspecial)],
    ['Eficiencia de consumo', preliqMoney(row.eficienciaConsumo)],
    ['Concurso', preliqMoney(row.concurso)],
    ['HHEE', preliqMoney(row.hheeMonto)],
    ['Total haberes', preliqMoney(row.totalHaberes)],
    ['Total descuentos previsionales', preliqMoney(row.totalDescuentosPrevisionales)],
    ['Total descuentos', preliqMoney(row.totalDescuentos)],
    ['Sueldo líquido', preliqMoney(row.sueldoLiquido)],
    ['Total monetario', preliqMoney(row.totalMonetario)],
  ])}</div>`;
}

function renderPreliquidationResults(result) {
  preliquidationData = result;
  document.querySelector('#preliquidation-period').textContent = `Período ${result.period || 'sin definir'}`;
  const workerSelect = document.querySelector('#preliquidation-worker');
  document.querySelector('#preliquidation-counter').textContent = `${formatNumber(result.rows.length)} trabajador${result.rows.length === 1 ? '' : 'es'}`;
  workerSelect.innerHTML = result.rows.map((row) => `<option value="${escapeHtml(row.key)}">${escapeHtml(row.nombre || row.rut)} · ${escapeHtml(row.rut)}</option>`).join('');
  workerSelect.disabled = !result.rows.length;
  renderPreliquidationWorker(result.rows[0]);
}

function reviewRow(id, item) {
  const report = item.report;
  const sheetCount = report.sheets.filter((sheet) => sheet.rows > 0).length;
  const warning = report.warnings?.length ? `<span class="review-source-warning">${report.warnings.map(escapeHtml).join('<br>')}</span>` : '';
  return `<article class="review-source-row ${sourceDefinitions[id].required ? '' : 'optional'}"><div class="review-source-main"><span class="review-source-icon">${sourceDefinitions[id].short}</span><div class="review-source-name"><strong>${sourceDefinitions[id].label}</strong><span>${escapeHtml(item.file.name)}</span></div><span class="status-pill ${report.level}">${report.statusLabel}</span></div><div class="review-source-facts"><span>${formatNumber(report.rows)} filas detectadas</span><span>${formatNumber(sheetCount)} hojas con datos</span></div>${warning}</article>`;
}

clearButton.addEventListener('click', () => {
  state.sources = {};
  reviewStage.hidden = true;
  calculationStage.hidden = true;
  combustibleResultsStage.hidden = true;
  concursoHheeStage.hidden = true;
  preliquidationStage.hidden = true;
  actionBar.hidden = false;
  setActiveStep(1);
  renderCards();
  renderValidation();
});

continueButton.addEventListener('click', () => {
  if (continueButton.disabled) return;
  renderReview();
  actionBar.hidden = true;
  reviewStage.hidden = false;
  calculationStage.hidden = true;
  combustibleResultsStage.hidden = true;
  concursoHheeStage.hidden = true;
  preliquidationStage.hidden = true;
  setActiveStep(2);
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

backButton.addEventListener('click', () => {
  reviewStage.hidden = true;
  calculationStage.hidden = true;
  combustibleResultsStage.hidden = true;
  concursoHheeStage.hidden = true;
  preliquidationStage.hidden = true;
  actionBar.hidden = false;
  setActiveStep(1);
  introSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
});

reviewToCalculationButton.addEventListener('click', () => {
  if (continueButton.disabled) return;
  reviewStage.hidden = true;
  calculationStage.hidden = false;
  setActiveStep(3);
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

calculationBackButton.addEventListener('click', () => {
  calculationStage.hidden = true;
  reviewStage.hidden = false;
  setActiveStep(2);
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

calculateCombustibleButton.addEventListener('click', () => {
  const result = buildCombustibleResult();
  renderCombustibleResults(result);
  calculationStage.hidden = true;
  combustibleResultsStage.hidden = false;
  preliquidationStage.hidden = true;
  setActiveStep(3);
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

combustibleBackButton.addEventListener('click', () => {
  combustibleResultsStage.hidden = true;
  calculationStage.hidden = false;
  concursoHheeStage.hidden = true;
  preliquidationStage.hidden = true;
  setActiveStep(3);
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

continueConcursoButton.addEventListener('click', () => {
  renderConcursoHheeResults(buildConcursoHheeResult());
  combustibleResultsStage.hidden = true;
  concursoHheeStage.hidden = false;
  preliquidationStage.hidden = true;
  setActiveStep(3);
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

concursoHheeBackButton.addEventListener('click', () => {
  concursoHheeStage.hidden = true;
  combustibleResultsStage.hidden = false;
  preliquidationStage.hidden = true;
  setActiveStep(3);
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

continuePreliquidationButton.addEventListener('click', () => {
  renderPreliquidationResults(buildPreliquidationResult());
  concursoHheeStage.hidden = true;
  preliquidationStage.hidden = false;
  setActiveStep(3);
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

preliquidationBackButton.addEventListener('click', () => {
  preliquidationStage.hidden = true;
  concursoHheeStage.hidden = false;
  setActiveStep(3);
  window.scrollTo({ top: 0, behavior: 'smooth' });
});

document.querySelector('#preliquidation-worker').addEventListener('change', (event) => {
  const selected = preliquidationData.rows.find((row) => row.key === event.target.value);
  renderPreliquidationWorker(selected);
});

document.querySelector('#print-preliquidation-button').addEventListener('click', () => {
  if (preliquidationData.rows.length) window.print();
});

renderCards();
renderValidation();
