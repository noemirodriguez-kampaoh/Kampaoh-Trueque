// ============================================================
// KAMPAOH TRUEQUE — Apps Script v1.0
// Fase 0 (Estructura) + Fase 1 (PL-01) + Fase 2 (BDD + PL-02..06)
// Principio: NUNCA enviar emails automáticamente.
//             Solo crear borradores revisables en Gmail.
// ============================================================

// ── CONSTANTES ───────────────────────────────────────────────

const CFG = {
  SHEETS: {
    FORMULARIO:  'Formulario',
    BDD:         'BDD_Trueque',
    CAMPINGS:    'Campings',
    PLANTILLAS:  'Plantillas_Email',
    CONFIG:      'Config',
    DASHBOARD:   'Dashboard',
    LOG:         'Log_Acciones'
  },
  MENU: '⛺ Trueque',
  TRQ_PREFIX: 'TRQ-',
  FIRMA_DEFAULT: 'Equipo Kampaoh 💛'
};

// Headers esperados (para crear hojas nuevas)
const H = {
  FORMULARIO: ['Fecha','Talento','Descripción','Destino','Llegada','Salida',
               'Días actuación','Nombre','Email','Teléfono','Instagram/Web',
               'Notas extra','Vídeo','Estado','PL1'],
  FORM_TECH:  ['ID_Formulario','Importado_BDD','Fecha_PL1','Draft_PL1_ID'],
  BDD:        ['ID','Solicitud','Nombre','Email','Teléfono','DNI',
               'Nº Personas','Actividad','Propuesta','Camping','Llegada',
               'Salida','Localizador','Google Calendar','Estado','Estado Mails',
               'PL2','PL3','PL4','PL5','PL6'],
  CAMPINGS:   ['Camping','Email','Teléfono','Disponibilidad','Horario','Zona','Notas'],
  PLANTILLAS: ['ID','Uso','Mail','Nombre plantilla','Fase','Destinatario',
               'Tipo de actividad','Camping asociado','Asunto'],
  CONFIG:     ['Clave','Valor','Descripción'],
  LOG:        ['Fecha hora','Usuario','Acción','Hoja','Fila','ID_Formulario',
               'ID_Trueque','Plantilla','Destinatario','Resultado','Detalle',
               'Draft ID','Variables sin sustituir','Error'],
  DASHBOARD:  ['Dashboard Kampaoh Trueque']
};

// Dropdowns
const DD = {
  FORM_ESTADO:    ['Pendiente','Revisar','Sí','No'],
  FORM_PL1:       ['Pendiente','Borrador creado','Enviado','Falta email','Error plantilla','Error'],
  BDD_ESTADO:     ['Nuevo','Pendiente revisión','Aprobado','En proceso','Rechazado',
                   'Futuro','Reserva solicitada','Reserva confirmada','Calendar creado','Completado'],
  BDD_ESTADO_MAILS: ['Pendiente','Contactado','Cancelado','Contactado STOP','Completado'],
  BDD_CALENDAR:   ['Pendiente','Evento preparado','Evento creado',
                   'Error: falta localizador','Error: camping sin email','Revisar'],
  BDD_PL:         ['','Sí','Error']
};

// ── MENÚ ─────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(CFG.MENU)
    .addItem('1. Configurar estructura',               'configurarEstructura')
    .addSeparator()
    .addItem('2. Probar plantilla PL-01 (sin borrador)','probarPlantillaPL01')
    .addItem('3. Procesar fila seleccionada PL-01',    'procesarFilaSeleccionadaPL01')
    .addItem('4. Procesar pendientes PL-01',           'procesarPendientesPL01')
    .addSeparator()
    .addItem('5. Pasar fila seleccionada a BDD',       'pasarFilaSeleccionadaABDD')
    .addItem('6. Procesar aprobados a BDD',            'procesarAprobadosABDD')
    .addSeparator()
    .addItem('7. Procesar PL-02 Contactado',           'procesarPL02')
    .addItem('8. Procesar PL-03 Cancelado',            'procesarPL03')
    .addItem('9. Procesar PL-04 Contactado STOP',      'procesarPL04')
    .addItem('10. Procesar PL-05 Completado',          'procesarPL05')
    .addItem('11. Procesar PL-06 Aviso camping',       'procesarPL06')
    .addSeparator()
    .addItem('12. Procesar fila seleccionada según Estado Mails', 'procesarFilaSegunEstado')
    .addSeparator()
    .addItem('13. Abrir Dashboard',                    'abrirDashboard')
    .addItem('14. Ver Log_Acciones',                   'verLog')
    .addToUi();
}

// ── FASE 0: CONFIGURAR ESTRUCTURA ────────────────────────────

function configurarEstructura() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  ui.alert('⛺ Configurando estructura...', 'Haz clic en OK para continuar.', ui.ButtonSet.OK);

  try {
    archivarHojasAntiguas_(ss);
    asegurarHoja_(ss, CFG.SHEETS.FORMULARIO, H.FORMULARIO, true);
    asegurarHoja_(ss, CFG.SHEETS.BDD,        H.BDD);
    asegurarHoja_(ss, CFG.SHEETS.CAMPINGS,   H.CAMPINGS);
    asegurarHoja_(ss, CFG.SHEETS.PLANTILLAS, H.PLANTILLAS);
    asegurarHoja_(ss, CFG.SHEETS.CONFIG,     H.CONFIG);
    asegurarHoja_(ss, CFG.SHEETS.LOG,        H.LOG);
    asegurarHoja_(ss, CFG.SHEETS.DASHBOARD,  H.DASHBOARD);

    añadirColumnasTecnicasFormulario_(ss);
    ocultarColumnasTecnicasBDD_(ss);
    crearDesplegables_(ss);
    limpiarPlantillas_(ss);
    inicializarConfig_(ss);
    construirDashboard_(ss);

    registrarLog_({ accion: 'Configurar estructura', hoja: 'Sistema', resultado: 'OK',
                    detalle: 'Fase 0 completada correctamente' });

    ui.alert('✅ Estructura lista',
      'Hojas, columnas y desplegables configurados.\n\n' +
      'Siguiente paso: Prueba plantilla PL-01 desde el menú ⛺ Trueque.',
      ui.ButtonSet.OK);

  } catch (e) {
    registrarLog_({ accion: 'Configurar estructura', hoja: 'Sistema', resultado: 'Error', error: e.toString() });
    ui.alert('❌ Error', e.toString(), ui.ButtonSet.OK);
  }
}

// ---------- helpers de estructura ----------

function archivarHojasAntiguas_(ss) {
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name.toUpperCase().includes('ELIMINAR') && !name.startsWith('_ARCHIVO_')) {
      let base = '_ARCHIVO_' + name.replace(/ELIMINAR/gi, '').trim();
      let final = base, i = 1;
      while (ss.getSheetByName(final)) final = base + '_' + i++;
      sheet.setName(final);
      sheet.hideSheet();
    }
  });
}

function asegurarHoja_(ss, nombre, headers, preservar) {
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      darFormatoHeaders_(sheet, headers.length);
    }
  } else if (!preservar && headers && headers.length) {
    // Si la primera fila está vacía, poner headers
    const firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
    if (firstRow.every(c => c === '' || c === null)) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      darFormatoHeaders_(sheet, headers.length);
    }
  }
  return sheet;
}

function darFormatoHeaders_(sheet, n) {
  const r = sheet.getRange(1, 1, 1, n);
  r.setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
}

function añadirColumnasTecnicasFormulario_(ss) {
  const sheet = ss.getSheetByName(CFG.SHEETS.FORMULARIO);
  if (!sheet || sheet.getLastRow() === 0) return;

  const headers = getHeaders_(sheet);
  H.FORM_TECH.forEach(h => {
    if (!headers.includes(h)) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(h)
        .setBackground('#f0f0f0').setFontColor('#888').setFontStyle('italic');
      sheet.hideColumns(col);
    }
  });

  generarIDsFormulario_(ss);
}

function generarIDsFormulario_(ss) {
  const sheet = ss.getSheetByName(CFG.SHEETS.FORMULARIO);
  if (!sheet || sheet.getLastRow() <= 1) return;
  const headers = getHeaders_(sheet);
  const idCol = headers.indexOf('ID_Formulario');
  if (idCol === -1) return;
  const tz = Session.getScriptTimeZone();
  for (let r = 2; r <= sheet.getLastRow(); r++) {
    const cell = sheet.getRange(r, idCol + 1);
    if (!cell.getValue()) {
      cell.setValue('FORM-' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd') + '-R' + r);
    }
  }
}

function ocultarColumnasTecnicasBDD_(ss) {
  const sheet = ss.getSheetByName(CFG.SHEETS.BDD);
  if (!sheet || sheet.getLastRow() === 0) return;
  const headers = getHeaders_(sheet);
  ['PL2','PL3','PL4','PL5','PL6'].forEach(h => {
    const i = headers.indexOf(h);
    if (i >= 0) sheet.hideColumns(i + 1);
  });
}

function crearDesplegables_(ss) {
  crearDD_(ss, CFG.SHEETS.FORMULARIO, 'Estado',       DD.FORM_ESTADO);
  crearDD_(ss, CFG.SHEETS.FORMULARIO, 'PL1',          DD.FORM_PL1);
  crearDD_(ss, CFG.SHEETS.BDD,        'Estado',       DD.BDD_ESTADO);
  crearDD_(ss, CFG.SHEETS.BDD,        'Estado Mails', DD.BDD_ESTADO_MAILS);
  crearDD_(ss, CFG.SHEETS.BDD,        'Google Calendar', DD.BDD_CALENDAR);
  ['PL2','PL3','PL4','PL5','PL6'].forEach(c => crearDD_(ss, CFG.SHEETS.BDD, c, DD.BDD_PL));
  crearDDDesdeLista_(ss, CFG.SHEETS.BDD, 'Camping', CFG.SHEETS.CAMPINGS, 'Camping');
}

function crearDD_(ss, sheetName, colHeader, values) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() === 0) return;
  const i = getHeaders_(sheet).indexOf(colHeader);
  if (i === -1) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true).setAllowInvalid(false).build();
  sheet.getRange(2, i + 1, 1000, 1).setDataValidation(rule);
}

function crearDDDesdeLista_(ss, target, targetCol, src, srcCol) {
  const tSheet = ss.getSheetByName(target);
  const sSheet = ss.getSheetByName(src);
  if (!tSheet || !sSheet) return;
  const tHeaders = getHeaders_(tSheet);
  const col = tHeaders.indexOf(targetCol);
  if (col === -1) return;
  const sHeaders = getHeaders_(sSheet);
  const sCol = sHeaders.indexOf(srcCol);
  if (sCol === -1) return;
  if (sSheet.getLastRow() <= 1) return;
  const values = sSheet.getRange(2, sCol + 1, sSheet.getLastRow() - 1, 1)
    .getValues().flat().filter(v => v !== '');
  if (!values.length) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true).setAllowInvalid(true).build();
  tSheet.getRange(2, col + 1, 1000, 1).setDataValidation(rule);
}

function limpiarPlantillas_(ss) {
  const sheet = ss.getSheetByName(CFG.SHEETS.PLANTILLAS);
  if (!sheet || sheet.getLastRow() <= 1) return;
  const headers = getHeaders_(sheet);
  const mailIdx = headers.indexOf('Mail');
  const asuntoIdx = headers.indexOf('Asunto');
  for (let r = 2; r <= sheet.getLastRow(); r++) {
    if (mailIdx >= 0) {
      const c = sheet.getRange(r, mailIdx + 1);
      const clean = limpiarTexto_(c.getValue());
      if (clean !== c.getValue()) c.setValue(clean);
    }
    if (asuntoIdx >= 0) {
      const c = sheet.getRange(r, asuntoIdx + 1);
      const clean = limpiarTexto_(c.getValue());
      if (clean !== c.getValue()) c.setValue(clean);
    }
  }
}

function inicializarConfig_(ss) {
  const sheet = ss.getSheetByName(CFG.SHEETS.CONFIG);
  if (!sheet) return;
  if (sheet.getLastRow() > 1) return;   // ya tiene datos
  const rows = [
    H.CONFIG,
    ['firma',                     'Equipo Kampaoh 💛',               'Firma en emails'],
    ['nombre_firmante',           'Noemi Rodríguez',                 'Nombre del firmante'],
    ['cargo_firmante',            'Kampaoh · Departamento de Marketing', 'Cargo'],
    ['confirmacion_masiva_umbral','3',                               'A partir de cuántos borradores pedir confirmación'],
  ];
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  darFormatoHeaders_(sheet, 3);
}

function construirDashboard_(ss) {
  const sheet = ss.getSheetByName(CFG.SHEETS.DASHBOARD);
  if (!sheet) return;
  sheet.clearContents();

  const rows = [
    ['⛺ KAMPAOH TRUEQUE — Dashboard', '', ''],
    ['', '', ''],
    ['KPI', 'Valor', 'Descripción'],
    ['Total solicitudes recibidas',  "=COUNTA(Formulario!A2:A)",                                                          'Filas con fecha en Formulario'],
    ['Solicitudes pendientes revisar',"=COUNTIF(Formulario!N2:N,\"Pendiente\")",                                          'Estado = Pendiente'],
    ['Solicitudes aprobadas',        "=COUNTIF(Formulario!N2:N,\"Sí\")",                                                  'Estado = Sí'],
    ['Registros en BDD',             "=COUNTA(BDD_Trueque!A2:A)",                                                         'Filas con ID en BDD'],
    ['PL-01 creados',                "=COUNTIF(Formulario!O2:O,\"Borrador creado\")",                                     'PL1 = Borrador creado'],
    ['PL-02 pendientes',             "=SUMPRODUCT((BDD_Trueque!P2:P1000=\"Contactado\")*(BDD_Trueque!Q2:Q1000<>\"Sí\"))", 'Contactado sin PL2'],
    ['PL-03 pendientes',             "=SUMPRODUCT((BDD_Trueque!P2:P1000=\"Cancelado\")*(BDD_Trueque!R2:R1000<>\"Sí\"))",  'Cancelado sin PL3'],
    ['PL-04 pendientes',             "=SUMPRODUCT((BDD_Trueque!P2:P1000=\"Contactado STOP\")*(BDD_Trueque!S2:S1000<>\"Sí\"))", 'Contactado STOP sin PL4'],
    ['PL-05 pendientes',             "=SUMPRODUCT((BDD_Trueque!P2:P1000=\"Completado\")*(BDD_Trueque!T2:T1000<>\"Sí\"))", 'Completado sin PL5'],
    ['PL-06 pendientes',             "=SUMPRODUCT((BDD_Trueque!P2:P1000=\"Completado\")*(BDD_Trueque!U2:U1000<>\"Sí\"))", 'Completado sin PL6'],
    ['Calendar pendiente',           "=COUNTIF(BDD_Trueque!N2:N,\"Pendiente\")",                                          'Google Calendar = Pendiente'],
    ['Calendar creado',              "=COUNTIF(BDD_Trueque!N2:N,\"Evento creado\")",                                      'Google Calendar = Evento creado'],
    ['Errores en Log',               "=COUNTIF(Log_Acciones!J2:J,\"Error\")",                                             'Resultado = Error'],
  ];

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold').setFontColor('#1a1a2e');
  sheet.getRange(3, 1, 1, 3).setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
  for (let i = 4; i <= rows.length; i++) {
    sheet.getRange(i, 1, 1, 3).setBackground(i % 2 === 0 ? '#f3f3f3' : '#ffffff');
  }
  sheet.setColumnWidth(1, 260).setColumnWidth(2, 70).setColumnWidth(3, 320);
  sheet.setFrozenRows(3);
}

// ── UTILIDADES GENERALES ──────────────────────────────────────

function getHeaders_(sheet) {
  if (!sheet || sheet.getLastRow() === 0) return [];
  return sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
}

function getConfigVal_(key) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.CONFIG);
  if (!sheet || sheet.getLastRow() <= 1) return '';
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (const row of data) { if (row[0] === key) return row[1]; }
  return '';
}

function fmtFecha_(val) {
  if (!val) return '';
  try {
    const d = val instanceof Date ? val : new Date(val);
    if (isNaN(d)) return val.toString();
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } catch (_) { return val.toString(); }
}

function registrarLog_(p) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.LOG);
    if (!sheet) return;
    sheet.appendRow([
      new Date(),
      Session.getActiveUser().getEmail() || 'sistema',
      p.accion || '', p.hoja || '', p.fila || '',
      p.idFormulario || '', p.idTrueque || '',
      p.plantilla || '', p.destinatario || '',
      p.resultado || '', p.detalle || '',
      p.draftId || '', p.varsSin || '', p.error || ''
    ]);
  } catch (_) {}
}

// ── NORMALIZACIÓN DE CAMPINGS ─────────────────────────────────

const CAMPING_CLAVES = [
  { clave: 'las arenas',    nombre: 'Kampaoh Las Arenas' },
  { clave: 'ria de vigo',   nombre: 'Kampaoh Ría de Vigo' },
  { clave: 'isla cristina', nombre: 'Kampaoh Isla Cristina' },
  { clave: 'el palmar',     nombre: 'Kampaoh El Palmar' },
  { clave: 'palmar',        nombre: 'Kampaoh El Palmar' },
  { clave: 'trafalgar',     nombre: 'Kampaoh Trafalgar' },
  { clave: 'los canos',     nombre: 'Kampaoh Los Caños' },
  { clave: 'canos',         nombre: 'Kampaoh Los Caños' },
  { clave: 'cordoba',       nombre: 'Kampaoh Córdoba' },
];

function normText_(t) {
  if (!t) return '';
  return t.toString().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/^kampaoh\s*/i, '')
    .replace(/\s+/g, ' ').trim();
}

function normalizarCamping_(raw) {
  if (!raw) return { nombre: '', confianza: 'sin_dato' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.CAMPINGS);
  let catalogo = [];
  if (sheet && sheet.getLastRow() > 1) {
    catalogo = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().filter(v => v);
  }

  // 1. Exacto
  for (const c of catalogo) {
    if (c.toString().toLowerCase() === raw.toString().toLowerCase())
      return { nombre: c, confianza: 'exacto' };
  }
  // 2. Normalizado exacto
  const normRaw = normText_(raw);
  for (const c of catalogo) {
    if (normText_(c) === normRaw) return { nombre: c, confianza: 'normalizado' };
  }
  // 3. Clave parcial
  for (const { clave, nombre } of CAMPING_CLAVES) {
    if (normRaw.includes(clave)) return { nombre: nombre, confianza: 'parcial' };
  }
  // Sin match
  return { nombre: 'Revisar camping', confianza: 'sin_match' };
}

function getCampingData_(nombre) {
  if (!nombre || nombre === 'Revisar camping') return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.CAMPINGS);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const headers = getHeaders_(sheet);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  for (const row of rows) {
    const n = row[headers.indexOf('Camping')] || '';
    if (n === nombre || normText_(n) === normText_(nombre)) {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    }
  }
  return null;
}

// ── SISTEMA DE PLANTILLAS ─────────────────────────────────────

function getPlantilla_(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.SHEETS.PLANTILLAS);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const headers = getHeaders_(sheet);
  const idCol     = headers.indexOf('ID');
  const mailCol   = headers.indexOf('Mail');
  const asuntoCol = headers.indexOf('Asunto');
  if (idCol === -1) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  for (const row of rows) {
    if (row[idCol] && row[idCol].toString().trim() === id) {
      return {
        id,
        mail:   mailCol   >= 0 ? limpiarTexto_(row[mailCol])   : '',
        asunto: asuntoCol >= 0 ? limpiarTexto_(row[asuntoCol]) : ''
      };
    }
  }
  return null;
}

function limpiarTexto_(t) {
  if (!t || typeof t !== 'string') return t || '';
  let s = t.trim();
  while (s.startsWith('"""') && s.endsWith('"""')) s = s.slice(3, -3).trim();
  while (s.startsWith('""')  && s.endsWith('""'))  s = s.slice(2, -2).trim();
  if   (s.startsWith('"')   && s.endsWith('"') && s.length > 1) s = s.slice(1, -1).trim();
  s = s.replace(/^Asunto:\s*/i, '').replace(/^Cuerpo:\s*/i, '').replace(/ /g, ' ').trim();
  return s;
}

// Variables alias: canonical → lista de sinónimos
const VAR_ALIAS = {
  actividad:       ['talento'],
  camping:         ['destino'],
  artista:         ['nombre'],
  franja_horaria:  ['franja horaria', 'horario'],
  espacio:         ['espacio_hojacampings','zona'],
  notas_camping:   ['notas_hojacampings','material_camping'],
  email_camping:   [],
  telefono_camping:[],
};

function renderizar_(template, vars) {
  if (!template) return { texto: '', varsSin: [] };

  // Expandir alias bidireccional
  const v = Object.assign({}, vars);
  Object.entries(VAR_ALIAS).forEach(([canon, aliases]) => {
    aliases.forEach(alias => {
      if (v[canon] !== undefined && v[alias] === undefined) v[alias] = v[canon];
      if (v[alias] !== undefined && v[canon] === undefined) v[canon] = v[alias];
    });
  });

  const varsSin = [];
  let texto = template;

  // {{var}} primero, luego {var}
  texto = texto.replace(/\{\{([^}]+)\}\}/g, (m, k) => {
    const key = k.trim();
    if (v[key] !== undefined && v[key] !== '') return v[key];
    varsSin.push(`{{${key}}}`);
    return m;
  });
  texto = texto.replace(/\{([^}]+)\}/g, (m, k) => {
    const key = k.trim();
    if (v[key] !== undefined && v[key] !== '') return v[key];
    varsSin.push(`{${key}}`);
    return m;
  });

  return { texto, varsSin };
}

function varsFormulario_(rowData, headers) {
  const g = n => { const i = headers.indexOf(n); return i >= 0 ? (rowData[i] || '') : ''; };
  const llegada = g('Llegada'), salida = g('Salida');
  let fechaTexto = '';
  if (llegada && salida) fechaTexto = `del ${fmtFecha_(llegada)} al ${fmtFecha_(salida)}`;
  else if (llegada) fechaTexto = fmtFecha_(llegada);

  return {
    fecha: fmtFecha_(g('Fecha')), talento: g('Talento'), actividad: g('Talento'),
    descripcion: g('Descripción'), destino: g('Destino'), camping: g('Destino'),
    llegada: fmtFecha_(llegada), salida: fmtFecha_(salida), fecha_legible: fechaTexto,
    dias_actuacion: g('Días actuación'), nombre: g('Nombre'), artista: g('Nombre'),
    email: g('Email'), telefono: g('Teléfono'),
    instagram_web: g('Instagram/Web'), instagram: g('Instagram/Web'), web: g('Instagram/Web'),
    notas: g('Notas extra'), video: g('Vídeo'),
    firma: getConfigVal_('firma') || CFG.FIRMA_DEFAULT,
  };
}

function varsBDD_(rowData, headers, campingData) {
  const g = n => { const i = headers.indexOf(n); return i >= 0 ? (rowData[i] || '') : ''; };
  const llegada = g('Llegada'), salida = g('Salida');
  let fechaTexto = '';
  if (llegada && salida) fechaTexto = `del ${fmtFecha_(llegada)} al ${fmtFecha_(salida)}`;
  else if (llegada) fechaTexto = fmtFecha_(llegada);

  const vars = {
    id: g('ID'), fecha_solicitud: fmtFecha_(g('Solicitud')),
    nombre: g('Nombre'), artista: g('Nombre'), email: g('Email'),
    telefono: g('Teléfono'), dni: g('DNI'), personas: g('Nº Personas'), num_personas: g('Nº Personas'),
    actividad: g('Actividad'), talento: g('Actividad'), propuesta: g('Propuesta'),
    camping: g('Camping'), destino: g('Camping'),
    llegada: fmtFecha_(llegada), salida: fmtFecha_(salida), fecha: fechaTexto,
    localizador: g('Localizador'), estado: g('Estado'), estado_mails: g('Estado Mails'),
    firma: getConfigVal_('firma') || CFG.FIRMA_DEFAULT,
  };

  if (campingData) {
    vars.email_camping        = campingData['Email']         || '';
    vars.telefono_camping     = campingData['Teléfono']      || '';
    vars.disponibilidad_camping = campingData['Disponibilidad'] || '';
    vars.franja_horaria       = campingData['Horario']       || '';
    vars.horario              = campingData['Horario']       || '';
    vars.espacio              = campingData['Zona']          || '';
    vars.zona                 = campingData['Zona']          || '';
    vars.notas_camping        = campingData['Notas']         || '';
    vars.material_camping     = campingData['Notas']         || '';
    // Aliases de compatibilidad con espacios
    vars['franja horaria']    = campingData['Horario']       || '';
    vars.espacio_hojacampings = campingData['Zona']          || '';
    vars.notas_hojacampings   = campingData['Notas']         || '';
  }
  return vars;
}

function crearBorrador_(destinatario, asunto, cuerpoTexto) {
  const emails = destinatario.toString()
    .split(/[\n,;]+/).map(e => e.trim()).filter(e => e.includes('@'));
  if (!emails.length) throw new Error('Email no válido: ' + destinatario);

  const htmlBody = '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6;color:#333;">' +
    cuerpoTexto.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>') + '</div>';

  const opts = { htmlBody, from: 'marketing@kampaoh.com', name: 'Kampaoh Trueque' };
  if (emails.length > 1) opts.cc = emails.slice(1).join(',');

  const draft = GmailApp.createDraft(emails[0], asunto, cuerpoTexto, opts);
  return draft.getId();
}

// ── FASE 1: PL-01 ────────────────────────────────────────────

function probarPlantillaPL01() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const plt = getPlantilla_('PL-01');

  if (!plt) {
    ui.alert('❌ Plantilla no encontrada',
      'No existe PL-01 en la hoja Plantillas_Email.\nAsegúrate de que la columna ID tiene el valor "PL-01".',
      ui.ButtonSet.OK);
    return;
  }

  const sheet = ss.getSheetByName(CFG.SHEETS.FORMULARIO);
  if (!sheet || sheet.getLastRow() <= 1) {
    ui.alert('ℹ️ Sin datos', 'No hay datos en Formulario.', ui.ButtonSet.OK);
    return;
  }

  const headers = getHeaders_(sheet);
  const row = sheet.getRange(2, 1, 1, headers.length).getValues()[0];
  const vars = varsFormulario_(row, headers);

  const rAsunto = renderizar_(plt.asunto, vars);
  const rCuerpo = renderizar_(plt.mail,   vars);
  const varsSin = [...rAsunto.varsSin, ...rCuerpo.varsSin];

  let msg = '=== PREVIEW PL-01 (fila 2) ===\n\n';
  msg += 'ASUNTO:\n' + rAsunto.texto + '\n\n';
  msg += 'BODY (primeros 600 chars):\n' + rCuerpo.texto.substring(0, 600) + (rCuerpo.texto.length > 600 ? '…' : '') + '\n\n';
  if (varsSin.length) msg += '⚠️ Variables sin sustituir: ' + varsSin.join(', ') + '\n\n';
  msg += 'No se creó ningún borrador. Solo preview.';

  ui.alert('🔍 Preview PL-01', msg, ui.ButtonSet.OK);
}

function procesarFilaSeleccionadaPL01() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ui   = SpreadsheetApp.getUi();
  const hoja = ss.getActiveSheet();
  if (hoja.getName() !== CFG.SHEETS.FORMULARIO) {
    ui.alert('⚠️', 'Debes estar en la hoja "Formulario".', ui.ButtonSet.OK); return;
  }
  const row = ss.getActiveRange().getRow();
  if (row <= 1) { ui.alert('⚠️', 'Selecciona una fila de datos.', ui.ButtonSet.OK); return; }
  procesarFilaPL01_(hoja, row, true);
}

function procesarPendientesPL01() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CFG.SHEETS.FORMULARIO);
  if (!sheet || sheet.getLastRow() <= 1) {
    ui.alert('ℹ️', 'No hay datos en Formulario.', ui.ButtonSet.OK); return;
  }

  const headers   = getHeaders_(sheet);
  const pl1Col    = headers.indexOf('PL1');
  const emailCol  = headers.indexOf('Email');
  const draftCol  = headers.indexOf('Draft_PL1_ID');

  const pendientes = [];
  for (let r = 2; r <= sheet.getLastRow(); r++) {
    const pl1  = pl1Col  >= 0 ? sheet.getRange(r, pl1Col  + 1).getValue() : '';
    const draft = draftCol >= 0 ? sheet.getRange(r, draftCol + 1).getValue() : '';
    const email = emailCol >= 0 ? sheet.getRange(r, emailCol + 1).getValue() : '';
    if (!pl1 || pl1 === 'Pendiente') {
      if (draft) continue; // ya tiene draft
      if (email) pendientes.push(r);
    }
  }

  if (!pendientes.length) {
    ui.alert('✅', 'No hay pendientes de PL-01.', ui.ButtonSet.OK); return;
  }

  const umbral = parseInt(getConfigVal_('confirmacion_masiva_umbral')) || 3;
  if (pendientes.length > umbral) {
    const resp = ui.alert('⚠️ Confirmar proceso masivo',
      `Se van a crear ${pendientes.length} borradores de PL-01.\n¿Continuar?`,
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) { ui.alert('ℹ️', 'Cancelado.', ui.ButtonSet.OK); return; }
  }

  let ok = 0, err = 0;
  pendientes.forEach(r => { const res = procesarFilaPL01_(sheet, r, false); res.ok ? ok++ : err++; });
  ui.alert('✅ Listo', `Borradores creados: ${ok}\nErrores: ${err}\nRevisa Log_Acciones.`, ui.ButtonSet.OK);
}

function procesarFilaPL01_(sheet, rowNum, alerta) {
  const ui      = SpreadsheetApp.getUi();
  const headers = getHeaders_(sheet);
  const row     = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  const g       = n => { const i = headers.indexOf(n); return i >= 0 ? (row[i] || '') : ''; };

  const email     = g('Email');
  const nombre    = g('Nombre');
  const pl1Idx    = headers.indexOf('PL1');
  const draftIdx  = headers.indexOf('Draft_PL1_ID');
  const fechaIdx  = headers.indexOf('Fecha_PL1');
  const idFormVal = g('ID_Formulario') || 'Fila-' + rowNum;

  const pl1Val   = g('PL1');
  const draftVal = draftIdx >= 0 ? row[draftIdx] : '';

  // Anti-duplicado
  if (pl1Val === 'Borrador creado' || pl1Val === 'Enviado' || draftVal) {
    const msg = `Fila ${rowNum} (${nombre}): ya tiene PL-01 (${pl1Val}). Se omite.`;
    registrarLog_({ accion: 'Procesar PL-01', hoja: CFG.SHEETS.FORMULARIO, fila: rowNum,
                    idFormulario: idFormVal, plantilla: 'PL-01', destinatario: email,
                    resultado: 'Saltado', detalle: msg });
    if (alerta) ui.alert('ℹ️ Ya procesado', msg, ui.ButtonSet.OK);
    return { ok: false };
  }

  // Email
  if (!email || !email.toString().includes('@')) {
    if (pl1Idx >= 0) sheet.getRange(rowNum, pl1Idx + 1).setValue('Falta email');
    registrarLog_({ accion: 'Procesar PL-01', hoja: CFG.SHEETS.FORMULARIO, fila: rowNum,
                    idFormulario: idFormVal, plantilla: 'PL-01', resultado: 'Error', error: 'Falta email' });
    if (alerta) ui.alert('⚠️ Falta email', `Fila ${rowNum}: sin email válido.`, ui.ButtonSet.OK);
    return { ok: false };
  }

  // Plantilla
  const plt = getPlantilla_('PL-01');
  if (!plt || !plt.mail) {
    if (pl1Idx >= 0) sheet.getRange(rowNum, pl1Idx + 1).setValue('Error plantilla');
    registrarLog_({ accion: 'Procesar PL-01', hoja: CFG.SHEETS.FORMULARIO, fila: rowNum,
                    idFormulario: idFormVal, plantilla: 'PL-01', resultado: 'Error', error: 'Plantilla PL-01 no encontrada' });
    if (alerta) ui.alert('❌', 'No se encontró la plantilla PL-01.', ui.ButtonSet.OK);
    return { ok: false };
  }

  // Renderizar
  const vars    = varsFormulario_(row, headers);
  const rAsunto = renderizar_(plt.asunto, vars);
  const rCuerpo = renderizar_(plt.mail,   vars);
  const varsSin = [...rAsunto.varsSin, ...rCuerpo.varsSin];

  if (!rAsunto.texto || !rCuerpo.texto) {
    if (pl1Idx >= 0) sheet.getRange(rowNum, pl1Idx + 1).setValue('Error');
    registrarLog_({ accion: 'Procesar PL-01', hoja: CFG.SHEETS.FORMULARIO, fila: rowNum,
                    idFormulario: idFormVal, plantilla: 'PL-01', resultado: 'Error',
                    error: 'Asunto o cuerpo vacío tras renderizado' });
    if (alerta) ui.alert('❌', 'Asunto o cuerpo quedaron vacíos.', ui.ButtonSet.OK);
    return { ok: false };
  }

  // Crear borrador
  try {
    const draftId = crearBorrador_(email, rAsunto.texto, rCuerpo.texto);
    if (pl1Idx  >= 0) sheet.getRange(rowNum, pl1Idx  + 1).setValue('Borrador creado');
    if (fechaIdx >= 0) sheet.getRange(rowNum, fechaIdx + 1).setValue(new Date());
    if (draftIdx >= 0) sheet.getRange(rowNum, draftIdx + 1).setValue(draftId);

    registrarLog_({ accion: 'Procesar PL-01', hoja: CFG.SHEETS.FORMULARIO, fila: rowNum,
                    idFormulario: idFormVal, plantilla: 'PL-01', destinatario: email,
                    resultado: 'OK', detalle: `Borrador creado para ${nombre}`,
                    draftId, varsSin: varsSin.join(', ') });

    if (alerta) {
      let msg = `✅ Borrador PL-01 creado para:\n${nombre} (${email})\n\nRevisa tus borradores en Gmail.`;
      if (varsSin.length) msg += `\n\n⚠️ Variables sin sustituir: ${varsSin.join(', ')}`;
      ui.alert('✅ Borrador creado', msg, ui.ButtonSet.OK);
    }
    return { ok: true, draftId };

  } catch (e) {
    if (pl1Idx >= 0) sheet.getRange(rowNum, pl1Idx + 1).setValue('Error');
    registrarLog_({ accion: 'Procesar PL-01', hoja: CFG.SHEETS.FORMULARIO, fila: rowNum,
                    idFormulario: idFormVal, plantilla: 'PL-01', destinatario: email,
                    resultado: 'Error', error: e.toString() });
    if (alerta) ui.alert('❌ Error Gmail', e.toString(), ui.ButtonSet.OK);
    return { ok: false };
  }
}

// ── FASE 2: PASO A BDD ────────────────────────────────────────

function pasarFilaSeleccionadaABDD() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ui   = SpreadsheetApp.getUi();
  const hoja = ss.getActiveSheet();
  if (hoja.getName() !== CFG.SHEETS.FORMULARIO) {
    ui.alert('⚠️', 'Debes estar en la hoja "Formulario".', ui.ButtonSet.OK); return;
  }
  const row = ss.getActiveRange().getRow();
  if (row <= 1) { ui.alert('⚠️', 'Selecciona una fila de datos.', ui.ButtonSet.OK); return; }
  pasarFilaABDD_(hoja, row, true);
}

function procesarAprobadosABDD() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CFG.SHEETS.FORMULARIO);
  if (!sheet || sheet.getLastRow() <= 1) { ui.alert('ℹ️', 'No hay datos.', ui.ButtonSet.OK); return; }

  const headers     = getHeaders_(sheet);
  const estadoCol   = headers.indexOf('Estado');
  const importadoCol = headers.indexOf('Importado_BDD');

  const pendientes = [];
  for (let r = 2; r <= sheet.getLastRow(); r++) {
    const estado    = estadoCol    >= 0 ? sheet.getRange(r, estadoCol    + 1).getValue() : '';
    const importado = importadoCol >= 0 ? sheet.getRange(r, importadoCol + 1).getValue() : '';
    if (estado === 'Sí' && importado !== 'Sí') pendientes.push(r);
  }

  if (!pendientes.length) { ui.alert('✅', 'No hay aprobados sin importar.', ui.ButtonSet.OK); return; }

  const umbral = parseInt(getConfigVal_('confirmacion_masiva_umbral')) || 3;
  if (pendientes.length > umbral) {
    const resp = ui.alert('⚠️ Confirmar',
      `Se van a pasar ${pendientes.length} registros a BDD_Trueque.\n¿Continuar?`,
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  let ok = 0, err = 0;
  pendientes.forEach(r => { const res = pasarFilaABDD_(sheet, r, false); res.ok ? ok++ : err++; });
  ui.alert('✅ Listo', `Importados: ${ok}\nErrores: ${err}`, ui.ButtonSet.OK);
}

function pasarFilaABDD_(sheetForm, rowNum, alerta) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const ui      = SpreadsheetApp.getUi();
  const headers = getHeaders_(sheetForm);
  const row     = sheetForm.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  const g       = n => { const i = headers.indexOf(n); return i >= 0 ? (row[i] || '') : ''; };

  const estadoVal    = g('Estado');
  const importadoVal = g('Importado_BDD');
  const email        = g('Email');
  const idForm       = g('ID_Formulario') || 'Fila-' + rowNum;

  if (estadoVal !== 'Sí') {
    if (alerta) ui.alert('⚠️', `Fila ${rowNum}: Estado no es "Sí" (es "${estadoVal}").`, ui.ButtonSet.OK);
    return { ok: false };
  }
  if (importadoVal === 'Sí') {
    if (alerta) ui.alert('ℹ️', `Fila ${rowNum}: ya importada a BDD.`, ui.ButtonSet.OK);
    return { ok: false };
  }

  const bddSheet = ss.getSheetByName(CFG.SHEETS.BDD);
  if (!bddSheet) { if (alerta) ui.alert('❌', 'No existe BDD_Trueque.', ui.ButtonSet.OK); return { ok: false }; }

  // Advertir duplicado de email (sin bloquear)
  if (email && bddSheet.getLastRow() > 1) {
    const bddHeaders = getHeaders_(bddSheet);
    const emailBDDCol = bddHeaders.indexOf('Email');
    if (emailBDDCol >= 0) {
      const emails = bddSheet.getRange(2, emailBDDCol + 1, bddSheet.getLastRow() - 1, 1).getValues().flat();
      if (emails.includes(email) && alerta) {
        const resp = ui.alert('⚠️ Posible duplicado',
          `El email "${email}" ya existe en BDD_Trueque.\n¿Continuar igualmente?`,
          ui.ButtonSet.YES_NO);
        if (resp !== ui.Button.YES) return { ok: false };
      }
    }
  }

  // Generar ID y datos
  const nuevaId    = generarIdTRQ_(bddSheet);
  const campingNorm = normalizarCamping_(g('Destino'));
  const propuesta  = generarPropuesta_(row, headers, campingNorm.nombre);

  const bddRow = new Array(H.BDD.length).fill('');
  bddRow[0]  = nuevaId;                   // ID
  bddRow[1]  = fmtFecha_(g('Fecha'));     // Solicitud
  bddRow[2]  = g('Nombre');              // Nombre
  bddRow[3]  = email;                    // Email
  bddRow[4]  = g('Teléfono');            // Teléfono
  bddRow[5]  = '';                       // DNI
  bddRow[6]  = '';                       // Nº Personas
  bddRow[7]  = g('Talento');             // Actividad
  bddRow[8]  = propuesta;               // Propuesta
  bddRow[9]  = campingNorm.nombre;      // Camping
  bddRow[10] = g('Llegada');             // Llegada
  bddRow[11] = g('Salida');              // Salida
  bddRow[12] = '';                       // Localizador
  bddRow[13] = 'Pendiente';             // Google Calendar
  bddRow[14] = 'Pendiente revisión';    // Estado
  bddRow[15] = 'Pendiente';             // Estado Mails
  // PL2-PL6 vacíos (índices 16-20)

  try {
    bddSheet.appendRow(bddRow);

    const impIdx = headers.indexOf('Importado_BDD');
    if (impIdx >= 0) sheetForm.getRange(rowNum, impIdx + 1).setValue('Sí');

    registrarLog_({ accion: 'Pasar a BDD', hoja: CFG.SHEETS.FORMULARIO, fila: rowNum,
                    idFormulario: idForm, idTrueque: nuevaId, resultado: 'OK',
                    detalle: `${g('Nombre')} importado. Camping: ${campingNorm.nombre} (${campingNorm.confianza})` });

    if (alerta) {
      let msg = `✅ Importado:\n• ID: ${nuevaId}\n• Nombre: ${g('Nombre')}\n• Camping: ${campingNorm.nombre}`;
      if (campingNorm.confianza === 'sin_match') msg += '\n\n⚠️ Camping no reconocido. Revísalo en BDD.';
      if (campingNorm.confianza === 'parcial')   msg += '\n\n⚠️ Camping mapeado por match parcial. Verifica.';
      ui.alert('✅ Importado', msg, ui.ButtonSet.OK);
    }
    return { ok: true, id: nuevaId };

  } catch (e) {
    registrarLog_({ accion: 'Pasar a BDD', hoja: CFG.SHEETS.FORMULARIO, fila: rowNum,
                    idFormulario: idForm, resultado: 'Error', error: e.toString() });
    if (alerta) ui.alert('❌', e.toString(), ui.ButtonSet.OK);
    return { ok: false };
  }
}

function generarIdTRQ_(bddSheet) {
  if (!bddSheet || bddSheet.getLastRow() <= 1) return 'TRQ-001';
  const headers = getHeaders_(bddSheet);
  const col = headers.indexOf('ID');
  if (col === -1) return 'TRQ-001';
  const ids = bddSheet.getRange(2, col + 1, bddSheet.getLastRow() - 1, 1).getValues().flat()
    .filter(v => v && v.toString().startsWith('TRQ-'))
    .map(v => parseInt(v.toString().replace('TRQ-', '')) || 0);
  return 'TRQ-' + String(ids.length ? Math.max(...ids) + 1 : 1).padStart(3, '0');
}

function generarPropuesta_(row, headers, campingNorm) {
  const g = n => { const i = headers.indexOf(n); return i >= 0 ? (row[i] || '') : ''; };
  const talento = g('Talento'), llegada = g('Llegada'), salida = g('Salida');

  let p = talento ? `Propone realizar una actividad de ${talento.toLowerCase()}` : 'Propone actividad';
  if (campingNorm && campingNorm !== 'Revisar camping') p += ` en ${campingNorm}`;
  if (llegada && salida) p += ` del ${fmtFecha_(llegada)} al ${fmtFecha_(salida)}`;
  else if (llegada) p += ` el ${fmtFecha_(llegada)}`;
  p += '.';
  const dias = g('Días actuación');
  const desc = g('Descripción'), ig = g('Instagram/Web'), vid = g('Vídeo'), notas = g('Notas extra');
  if (dias)  p += ` Indica ${dias} día(s) de actuación.`;
  if (desc)  p += ` ${desc}.`;
  if (ig)    p += ` Perfil: ${ig}.`;
  if (vid)   p += ` Vídeo: ${vid}.`;
  if (notas) p += ` Notas: ${notas}.`;
  return p.trim();
}

// ── FASE 2: PL-02 a PL-06 ────────────────────────────────────

function procesarPL02() { procesarPLBDD_('PL-02','Contactado','PL2'); }
function procesarPL03() { procesarPLBDD_('PL-03','Cancelado','PL3'); }
function procesarPL04() { procesarPLBDD_('PL-04','Contactado STOP','PL4'); }
function procesarPL05() { procesarPLBDD_('PL-05','Completado','PL5'); }

function procesarPL06() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CFG.SHEETS.BDD);
  if (!sheet || sheet.getLastRow() <= 1) { ui.alert('ℹ️','Sin datos en BDD.',ui.ButtonSet.OK); return; }

  const headers = getHeaders_(sheet);
  const emCol   = headers.indexOf('Estado Mails');
  const pl6Col  = headers.indexOf('PL6');
  const campCol = headers.indexOf('Camping');

  const pendientes = [];
  for (let r = 2; r <= sheet.getLastRow(); r++) {
    const em   = emCol   >= 0 ? sheet.getRange(r, emCol   + 1).getValue() : '';
    const pl6  = pl6Col  >= 0 ? sheet.getRange(r, pl6Col  + 1).getValue() : '';
    const camp = campCol >= 0 ? sheet.getRange(r, campCol + 1).getValue() : '';
    if (em === 'Completado' && !pl6 && camp) pendientes.push(r);
  }

  if (!pendientes.length) { ui.alert('✅','No hay pendientes de PL-06.',ui.ButtonSet.OK); return; }

  const umbral = parseInt(getConfigVal_('confirmacion_masiva_umbral')) || 3;
  if (pendientes.length > umbral) {
    const resp = ui.alert('⚠️ Confirmar',
      `Se van a crear ${pendientes.length} borradores de PL-06 (aviso camping).\n¿Continuar?`,
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  let ok = 0, err = 0;
  pendientes.forEach(r => { const res = procesarFilaPL06_(sheet, r, false); res.ok ? ok++ : err++; });
  ui.alert('✅ Listo', `PL-06 creados: ${ok}\nErrores: ${err}`, ui.ButtonSet.OK);
}

function procesarPLBDD_(pltId, estadoReq, colPL) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const ui    = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(CFG.SHEETS.BDD);
  if (!sheet || sheet.getLastRow() <= 1) { ui.alert('ℹ️','Sin datos en BDD.',ui.ButtonSet.OK); return; }

  const headers = getHeaders_(sheet);
  const emCol   = headers.indexOf('Estado Mails');
  const plCol   = headers.indexOf(colPL);

  const pendientes = [];
  for (let r = 2; r <= sheet.getLastRow(); r++) {
    const em = emCol >= 0 ? sheet.getRange(r, emCol + 1).getValue() : '';
    const pl = plCol >= 0 ? sheet.getRange(r, plCol + 1).getValue() : '';
    if (em === estadoReq && !pl) pendientes.push(r);
  }

  if (!pendientes.length) {
    ui.alert('✅', `No hay pendientes de ${pltId} (Estado Mails = "${estadoReq}").`, ui.ButtonSet.OK); return;
  }

  const umbral = parseInt(getConfigVal_('confirmacion_masiva_umbral')) || 3;
  if (pendientes.length > umbral) {
    const resp = ui.alert('⚠️ Confirmar',
      `Se van a crear ${pendientes.length} borradores de ${pltId}.\n¿Continuar?`,
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  let ok = 0, err = 0;
  pendientes.forEach(r => {
    const res = procesarFilaPL_BDD_(sheet, r, pltId, colPL, false);
    res.ok ? ok++ : err++;
  });
  ui.alert('✅ Listo', `${pltId} creados: ${ok}\nErrores: ${err}`, ui.ButtonSet.OK);
}

function procesarFilaPL_BDD_(sheet, rowNum, pltId, colPL, alerta) {
  const ui      = SpreadsheetApp.getUi();
  const headers = getHeaders_(sheet);
  const row     = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  const g       = n => { const i = headers.indexOf(n); return i >= 0 ? (row[i] || '') : ''; };

  const email     = g('Email');
  const nombre    = g('Nombre');
  const idTrueque = g('ID');
  const plColIdx  = headers.indexOf(colPL);
  const plVal     = plColIdx >= 0 ? row[plColIdx] : '';

  if (plVal === 'Sí') {
    if (alerta) ui.alert('ℹ️', `${pltId} ya procesado para esta fila.`, ui.ButtonSet.OK);
    return { ok: false };
  }
  if (!email || !email.toString().includes('@')) {
    if (plColIdx >= 0) sheet.getRange(rowNum, plColIdx + 1).setValue('Error');
    registrarLog_({ accion: `Procesar ${pltId}`, hoja: CFG.SHEETS.BDD, fila: rowNum,
                    idTrueque, plantilla: pltId, resultado: 'Error', error: 'Falta email' });
    if (alerta) ui.alert('⚠️', 'Falta email.', ui.ButtonSet.OK);
    return { ok: false };
  }

  const plt = getPlantilla_(pltId);
  if (!plt || !plt.mail) {
    if (plColIdx >= 0) sheet.getRange(rowNum, plColIdx + 1).setValue('Error');
    registrarLog_({ accion: `Procesar ${pltId}`, hoja: CFG.SHEETS.BDD, fila: rowNum,
                    idTrueque, plantilla: pltId, resultado: 'Error', error: 'Plantilla no encontrada' });
    if (alerta) ui.alert('❌', `Plantilla ${pltId} no encontrada.`, ui.ButtonSet.OK);
    return { ok: false };
  }

  const campingData = getCampingData_(g('Camping'));
  const vars        = varsBDD_(row, headers, campingData);
  const rAsunto     = renderizar_(plt.asunto, vars);
  const rCuerpo     = renderizar_(plt.mail,   vars);
  const varsSin     = [...rAsunto.varsSin, ...rCuerpo.varsSin];

  try {
    const draftId = crearBorrador_(email, rAsunto.texto, rCuerpo.texto);
    if (plColIdx >= 0) sheet.getRange(rowNum, plColIdx + 1).setValue('Sí');

    registrarLog_({ accion: `Procesar ${pltId}`, hoja: CFG.SHEETS.BDD, fila: rowNum,
                    idTrueque, plantilla: pltId, destinatario: email, resultado: 'OK',
                    draftId, varsSin: varsSin.join(', ') });

    if (alerta) {
      let msg = `✅ Borrador ${pltId} creado para:\n${nombre} (${email})`;
      if (varsSin.length) msg += `\n\n⚠️ Variables sin sustituir: ${varsSin.join(', ')}`;
      ui.alert(`✅ ${pltId}`, msg, ui.ButtonSet.OK);
    }
    return { ok: true };

  } catch (e) {
    if (plColIdx >= 0) sheet.getRange(rowNum, plColIdx + 1).setValue('Error');
    registrarLog_({ accion: `Procesar ${pltId}`, hoja: CFG.SHEETS.BDD, fila: rowNum,
                    idTrueque, plantilla: pltId, destinatario: email,
                    resultado: 'Error', error: e.toString() });
    if (alerta) ui.alert('❌', e.toString(), ui.ButtonSet.OK);
    return { ok: false };
  }
}

function procesarFilaPL06_(sheet, rowNum, alerta) {
  const ui      = SpreadsheetApp.getUi();
  const headers = getHeaders_(sheet);
  const row     = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  const g       = n => { const i = headers.indexOf(n); return i >= 0 ? (row[i] || '') : ''; };

  const camping   = g('Camping');
  const idTrueque = g('ID');
  const pl6Idx    = headers.indexOf('PL6');
  const pl6Val    = pl6Idx >= 0 ? row[pl6Idx] : '';

  if (pl6Val === 'Sí') {
    if (alerta) ui.alert('ℹ️', 'PL-06 ya procesado.', ui.ButtonSet.OK); return { ok: false };
  }
  if (!camping) {
    if (pl6Idx >= 0) sheet.getRange(rowNum, pl6Idx + 1).setValue('Error');
    registrarLog_({ accion: 'Procesar PL-06', hoja: CFG.SHEETS.BDD, fila: rowNum, idTrueque,
                    plantilla: 'PL-06', resultado: 'Error', error: 'Camping vacío' });
    if (alerta) ui.alert('⚠️', 'Esta fila no tiene camping.', ui.ButtonSet.OK);
    return { ok: false };
  }

  const campingData = getCampingData_(camping);
  if (!campingData) {
    if (pl6Idx >= 0) sheet.getRange(rowNum, pl6Idx + 1).setValue('Error');
    registrarLog_({ accion: 'Procesar PL-06', hoja: CFG.SHEETS.BDD, fila: rowNum, idTrueque,
                    plantilla: 'PL-06', resultado: 'Error', error: `Camping "${camping}" no en catálogo` });
    if (alerta) ui.alert('❌', `Camping "${camping}" no encontrado en hoja Campings.`, ui.ButtonSet.OK);
    return { ok: false };
  }

  const emailCamping = campingData['Email'] || '';
  if (!emailCamping) {
    if (pl6Idx >= 0) sheet.getRange(rowNum, pl6Idx + 1).setValue('Camping sin email');
    registrarLog_({ accion: 'Procesar PL-06', hoja: CFG.SHEETS.BDD, fila: rowNum, idTrueque,
                    plantilla: 'PL-06', resultado: 'Error', error: `Camping "${camping}" sin email` });
    if (alerta) ui.alert('⚠️', `El camping "${camping}" no tiene email.`, ui.ButtonSet.OK);
    return { ok: false };
  }

  const plt = getPlantilla_('PL-06');
  if (!plt || !plt.mail) {
    if (pl6Idx >= 0) sheet.getRange(rowNum, pl6Idx + 1).setValue('Error');
    registrarLog_({ accion: 'Procesar PL-06', hoja: CFG.SHEETS.BDD, fila: rowNum, idTrueque,
                    plantilla: 'PL-06', resultado: 'Error', error: 'Plantilla PL-06 no encontrada' });
    if (alerta) ui.alert('❌', 'Plantilla PL-06 no encontrada.', ui.ButtonSet.OK);
    return { ok: false };
  }

  const vars    = varsBDD_(row, headers, campingData);
  const rAsunto = renderizar_(plt.asunto, vars);
  const rCuerpo = renderizar_(plt.mail,   vars);
  const varsSin = [...rAsunto.varsSin, ...rCuerpo.varsSin];

  try {
    const draftId = crearBorrador_(emailCamping, rAsunto.texto, rCuerpo.texto);
    if (pl6Idx >= 0) sheet.getRange(rowNum, pl6Idx + 1).setValue('Sí');

    registrarLog_({ accion: 'Procesar PL-06', hoja: CFG.SHEETS.BDD, fila: rowNum,
                    idTrueque, plantilla: 'PL-06', destinatario: emailCamping,
                    resultado: 'OK', draftId, varsSin: varsSin.join(', ') });

    if (alerta) {
      let msg = `✅ Borrador PL-06 para camping:\n${camping} (${emailCamping})`;
      if (varsSin.length) msg += `\n\n⚠️ Variables sin sustituir: ${varsSin.join(', ')}`;
      ui.alert('✅ PL-06', msg, ui.ButtonSet.OK);
    }
    return { ok: true };

  } catch (e) {
    if (pl6Idx >= 0) sheet.getRange(rowNum, pl6Idx + 1).setValue('Error');
    registrarLog_({ accion: 'Procesar PL-06', hoja: CFG.SHEETS.BDD, fila: rowNum,
                    idTrueque, plantilla: 'PL-06', destinatario: emailCamping,
                    resultado: 'Error', error: e.toString() });
    if (alerta) ui.alert('❌', e.toString(), ui.ButtonSet.OK);
    return { ok: false };
  }
}

// Procesar fila seleccionada en BDD según su Estado Mails
function procesarFilaSegunEstado() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ui   = SpreadsheetApp.getUi();
  const hoja = ss.getActiveSheet();
  if (hoja.getName() !== CFG.SHEETS.BDD) {
    ui.alert('⚠️', 'Debes estar en la hoja "BDD_Trueque".', ui.ButtonSet.OK); return;
  }
  const row = ss.getActiveRange().getRow();
  if (row <= 1) { ui.alert('⚠️', 'Selecciona una fila de datos.', ui.ButtonSet.OK); return; }

  const headers = getHeaders_(hoja);
  const emIdx   = headers.indexOf('Estado Mails');
  if (emIdx === -1) { ui.alert('❌', 'No existe columna "Estado Mails".', ui.ButtonSet.OK); return; }

  const estado = hoja.getRange(row, emIdx + 1).getValue();
  const mapa = {
    'Contactado':      { plt: 'PL-02', col: 'PL2' },
    'Cancelado':       { plt: 'PL-03', col: 'PL3' },
    'Contactado STOP': { plt: 'PL-04', col: 'PL4' },
    'Completado':      { plt: 'PL-05', col: 'PL5' },
  };

  const accion = mapa[estado];
  if (!accion) {
    ui.alert('ℹ️', `Estado Mails = "${estado}" no tiene plantilla automática.\nVálidos: Contactado, Cancelado, Contactado STOP, Completado.`, ui.ButtonSet.OK);
    return;
  }

  procesarFilaPL_BDD_(hoja, row, accion.plt, accion.col, true);

  if (estado === 'Completado') {
    const resp = ui.alert('¿También PL-06?', '¿Preparar también el aviso al camping (PL-06)?', ui.ButtonSet.YES_NO);
    if (resp === ui.Button.YES) procesarFilaPL06_(hoja, row, true);
  }
}

// ── NAVEGACIÓN ────────────────────────────────────────────────

function abrirDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.DASHBOARD);
  if (sheet) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  else SpreadsheetApp.getUi().alert('ℹ️', 'Dashboard no encontrado. Ejecuta "Configurar estructura".', SpreadsheetApp.getUi().ButtonSet.OK);
}

function verLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.LOG);
  if (sheet) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  else SpreadsheetApp.getUi().alert('ℹ️', 'Log no encontrado. Ejecuta "Configurar estructura".', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── CALENDAR (Fase 4 — stub preparado) ───────────────────────
// function crearEventoCalendar_(row, headers) {
//   TODO: implementar con CalendarApp en Fase 4
// }
