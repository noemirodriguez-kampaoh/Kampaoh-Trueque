/****************************************************************************
 * KAMPAOH · TRUEQUE — Apps Script (Fase 0 + Fase 1) — v2 reforzada
 * --------------------------------------------------------------------------
 * Este código vive DENTRO de tu propia Google Sheet (Extensiones > Apps Script).
 * Una vez pegado, opera sobre tus pestañas reales y tu Gmail real.
 *
 * FASE 0  -> Limpia y prepara la estructura (renombra hojas, archiva las
 *            "ELIMINAR" de forma segura, crea Config / Log_Acciones / Dashboard,
 *            añade columnas técnicas, oculta columnas y crea los desplegables).
 * FASE 1  -> Cuando entra una solicitud nueva en "Formulario":
 *            genera ID_Formulario, prepara el BORRADOR de PL-01 en Gmail,
 *            marca PL1 = "Borrador creado", evita duplicados y deja registro
 *            en "Log_Acciones".
 *
 * SEMÁNTICA DE ESTADOS (importante):
 *   - PL1 = "Borrador creado"  -> YA GESTIONADO (borrador en Gmail). NO implica
 *     que se haya enviado. El envío final lo haces tú a mano.
 *   - En BDD_Trueque, PL2..PL6 = "Sí" tendrá el MISMO significado: gestionado /
 *     borrador creado, para evitar duplicados aunque el envío sea manual.
 *
 * REGLA: NUNCA se envía un email automáticamente. Solo borradores revisables.
 *
 * Menú "⛺ Trueque" (aparece arriba al abrir la hoja):
 *   1) Configurar estructura (Fase 0)
 *   2) Instalar automatización del formulario (Fase 1)
 *   3) Procesar solicitudes pendientes (PL-01)        <- también para probar
 *   · Probar plantilla PL-01 (sin crear borrador)      <- prueba 100% segura
 ****************************************************************************/


/* ============================ CONFIGURACIÓN ============================ */
const NOMBRES = {
  FORMULARIO: { limpio: 'Formulario',        simple: 'Formulario' },
  BDD:        { limpio: 'BDD_Trueque',       simple: 'BDD_Trueque (SIMPLE)' },
  CAMPINGS:   { limpio: 'Campings',          simple: 'Campings (SIMPLE)' },
  PLANTILLAS: { limpio: 'Plantillas_Email',  simple: 'Plantillas_Email (SIMPLE)' },
  CONFIG:     { limpio: 'Config',            simple: 'Config' },
  LOG:        { limpio: 'Log_Acciones',      simple: 'Log_Acciones' },
  DASHBOARD:  { limpio: 'Dashboard',         simple: 'Dashboard' }
};

const COLS_TECNICAS_FORM = ['ID_Formulario', 'Importado_BDD', 'Fecha_PL1', 'Draft_PL1_ID'];
const COLS_PL_BDD = ['PL2', 'PL3', 'PL4', 'PL5', 'PL6'];

const ESTADOS = {
  FORM: ['Sí', 'No', 'Revisar', 'Pendiente'],
  BDD: ['Nuevo', 'Pendiente revisión', 'Aprobado', 'En proceso', 'Rechazado',
        'Futuro', 'Reserva solicitada', 'Reserva confirmada', 'Calendar creado', 'Completado'],
  MAILS: ['Sin emails', 'PL1 preparado', 'PL1 enviado', 'PL2 pendiente', 'PL2 preparado',
          'PL2 enviado', 'Rechazo enviado', 'Futuro enviado', 'Camping avisado', 'Completo'],
  CALENDAR: ['Pendiente', 'Evento preparado', 'Evento creado', 'Error', 'Revisar'],
  PL_SI: ['Sí']
};

const FIRMA = 'Noemi Rodríguez\nKampaoh · Departamento de Marketing';
const CALENDAR_ID = 'noemirodriguez@kampaoh.com';


/* ============================ MENÚ ============================ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⛺ Trueque')
    .addItem('1) Configurar estructura (Fase 0)', 'setupFase0')
    .addItem('2) Instalar automatización del formulario (Fase 1)', 'instalarTriggers_PL1')
    .addItem('3) Procesar solicitudes pendientes (PL-01)', 'procesarSolicitudesPendientesPL1')
    .addSeparator()
    .addItem('Probar plantilla PL-01 (sin crear borrador)', 'testLeerPlantillaPL01')
    .addItem('Ver registro de acciones', 'irALog_')
    .addToUi();
}

function irALog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = getHojaFlexible_(ss, 'LOG');
  if (log) ss.setActiveSheet(log);
  else SpreadsheetApp.getUi().alert('Aún no existe "Log_Acciones". Ejecuta primero la Fase 0.');
}


/* ============================================================ */
/* =====================  FASE 0: ESTRUCTURA  ================= */
/* ============================================================ */

// Ejecuta cada paso por separado: si uno falla, lo registra y sigue,
// y al final te muestra exactamente qué pasos fallaron.
function setupFase0() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pasos = [
    ['Renombrar hojas', renombrarHojas_],
    ['Archivar hojas ELIMINAR', archivarHojasEliminar_],
    ['Columnas técnicas Formulario', asegurarColumnasTecnicasFormulario_],
    ['Hoja Config', asegurarHojaConfig_],
    ['Hoja Log_Acciones', asegurarHojaLog_],
    ['Hoja Dashboard', asegurarHojaDashboard_],
    ['Desplegables / validaciones', configurarValidaciones_],
    ['Ocultar columnas técnicas', ocultarColumnasTecnicas_]
  ];

  const lineas = [];
  let huboError = false, huboAviso = false;

  pasos.forEach(function (p) {
    const nombrePaso = p[0], fn = p[1];
    try {
      const r = String(fn(ss));
      if (r.indexOf('⚠') !== -1) huboAviso = true;
      lineas.push('✓ ' + nombrePaso + ': ' + r);
    } catch (err) {
      huboError = true;
      const msg = String(err && err.message ? err.message : err);
      lineas.push('✗ ' + nombrePaso + ' FALLÓ: ' + msg);
      registrarLog_('Setup Fase 0', 'Config', '-', nombrePaso, 'Error', msg);
    }
  });

  registrarLog_('Setup Fase 0', 'Config', '-', '-',
    huboError ? 'Completado con errores' : (huboAviso ? 'Completado con avisos' : 'OK'),
    lineas.join(' | '));

  let cabecera = '✅ Fase 0 completada correctamente:\n\n';
  if (huboError) cabecera = '⚠ Fase 0 terminó CON ERRORES. Revisa los pasos con ✗ y la hoja "Log_Acciones":\n\n';
  else if (huboAviso) cabecera = '⚠ Fase 0 completada, pero con AVISOS que debes revisar:\n\n';

  SpreadsheetApp.getUi().alert(cabecera + lineas.join('\n'));
}

// (Ajuste 7) Seguridad: si coexisten el nombre limpio y el "(SIMPLE)", NO se
// renombra ni se fusiona nada; se avisa para que lo resuelvas a mano.
function renombrarHojas_(ss) {
  const hechos = [];
  const avisos = [];
  ['BDD', 'CAMPINGS', 'PLANTILLAS'].forEach(function (k) {
    const n = NOMBRES[k];
    const limpio = buscarHoja_(ss, n.limpio);
    const simple = (n.simple !== n.limpio) ? buscarHoja_(ss, n.simple) : null;
    if (limpio && simple) {
      avisos.push('⚠ Coexisten "' + n.limpio + '" y "' + n.simple + '": NO se renombra para no fusionar/confundir. Revísalo a mano.');
      return;
    }
    if (limpio) return;
    if (simple) { simple.setName(n.limpio); hechos.push(n.simple + ' → ' + n.limpio); }
  });
  let out = hechos.length ? ('Renombradas: ' + hechos.join(', ')) : 'Nombres ya correctos';
  if (avisos.length) out += ' || ' + avisos.join(' ');
  return out;
}

// (Ajuste 1) Archivado seguro: si el nombre de archivo ya existe, añade
// sufijo único con fecha/hora y, si hiciera falta, un contador.
function archivarHojasEliminar_(ss) {
  let n = 0;
  ss.getSheets().forEach(function (h) {
    const nombre = h.getName();
    if (/ELIMINAR/i.test(nombre) && !/^_ARCHIVO_/.test(nombre)) {
      h.setName(nombreArchivoUnico_(ss, '_ARCHIVO_' + nombre));
      h.hideSheet();
      n++;
    }
  });
  return n ? (n + ' hoja(s) "ELIMINAR" archivadas y ocultas (no borradas)') : 'No hay hojas "ELIMINAR" pendientes';
}

function nombreArchivoUnico_(ss, base) {
  if (!buscarHoja_(ss, base)) return base;
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'Europe/Madrid', 'yyyyMMdd_HHmmss');
  let candidato = base + '_' + stamp;
  let i = 2;
  while (buscarHoja_(ss, candidato)) { candidato = base + '_' + stamp + '_' + i; i++; }
  return candidato;
}

function asegurarColumnasTecnicasFormulario_(ss) {
  const hoja = getHojaFlexible_(ss, 'FORMULARIO');
  if (!hoja) throw new Error('No se encontró la hoja "Formulario".');
  const ultCol = Math.max(hoja.getLastColumn(), 1);
  const headers = hoja.getRange(1, 1, 1, ultCol).getValues()[0];
  const presentes = new Set(headers.map(normalizeHeader_));
  let next = ultCol + 1;
  const añadidas = [];
  COLS_TECNICAS_FORM.forEach(function (h) {
    if (!presentes.has(normalizeHeader_(h))) {
      hoja.getRange(1, next).setValue(h);
      presentes.add(normalizeHeader_(h));
      añadidas.push(h);
      next++;
    }
  });
  return añadidas.length ? ('Añadidas a Formulario: ' + añadidas.join(', ')) : 'Ya existían';
}

function asegurarHojaConfig_(ss) {
  let h = getHojaFlexible_(ss, 'CONFIG');
  let creada = false;
  if (!h) { h = ss.insertSheet(NOMBRES.CONFIG.limpio); creada = true; }
  if (String(h.getRange('A1').getValue() || '').trim() === '') {
    h.getRange('A1').setValue('AJUSTES (no borres las claves de la columna A)').setFontWeight('bold');
    const ajustes = [
      ['Modo emails', 'Borrador'],
      ['Calendar ID', CALENDAR_ID],
      ['Firma (línea 1)', 'Noemi Rodríguez'],
      ['Firma (línea 2)', 'Kampaoh · Departamento de Marketing']
    ];
    h.getRange(2, 1, ajustes.length, 2).setValues(ajustes);

    h.getRange('A7').setValue('ALIAS CAMPINGS (lo que escribe la gente → camping oficial)').setFontWeight('bold');
    h.getRange('A8:B8').setValues([['Alias (minúsculas, sin tildes)', 'Camping oficial']]).setFontWeight('bold');
    const alias = [
      ['cordoba', 'Kampaoh Córdoba'],
      ['palmar', 'Kampaoh El Palmar'],
      ['las arenas', 'Kampaoh Las Arenas'],
      ['arenas', 'Kampaoh Las Arenas'],
      ['vigo', 'Kampaoh Ría de Vigo'],
      ['ria de vigo', 'Kampaoh Ría de Vigo'],
      ['isla cristina', 'Kampaoh Isla Cristina'],
      ['cristina', 'Kampaoh Isla Cristina'],
      ['trafalgar', 'Kampaoh Trafalgar'],
      ['los canos', 'Kampaoh Los Caños'],
      ['canos', 'Kampaoh Los Caños']
    ];
    h.getRange(9, 1, alias.length, 2).setValues(alias);
    h.setColumnWidth(1, 280);
    h.setColumnWidth(2, 240);
  }
  return creada ? 'Creada' : 'Ya existía';
}

function asegurarHojaLog_(ss) {
  let h = getHojaFlexible_(ss, 'LOG');
  let creada = false;
  if (!h) { h = ss.insertSheet(NOMBRES.LOG.limpio); creada = true; }
  if (String(h.getRange('A1').getValue() || '').trim() === '') {
    h.getRange(1, 1, 1, 7)
      .setValues([['Fecha y hora', 'Acción', 'Hoja', 'Referencia', 'Plantilla', 'Resultado', 'Detalle']])
      .setFontWeight('bold');
    h.setFrozenRows(1);
    h.setColumnWidth(1, 150);
    h.setColumnWidth(7, 320);
  }
  return creada ? 'Creada' : 'Ya existía';
}

function asegurarHojaDashboard_(ss) {
  let h = getHojaFlexible_(ss, 'DASHBOARD');
  let creada = false;
  if (!h) { h = ss.insertSheet(NOMBRES.DASHBOARD.limpio); creada = true; }
  if (String(h.getRange('A1').getValue() || '').trim() === '') {
    h.getRange('A1').setValue('Dashboard Trueque').setFontWeight('bold').setFontSize(16);
    h.getRange('A3').setValue('Los indicadores (KPIs) y la tabla filtrable se construyen en la Fase 5.');
  }
  return creada ? 'Creada (esqueleto)' : 'Ya existía';
}

function configurarValidaciones_(ss) {
  const hechos = [];
  const form = getHojaFlexible_(ss, 'FORMULARIO');
  if (form) {
    const cm = buildColMap_(form);
    if (aplicarLista_(form, col_(cm, 'estado'), ESTADOS.FORM)) hechos.push('Formulario.Estado');
  }
  const bdd = getHojaFlexible_(ss, 'BDD');
  if (bdd) {
    const cm = buildColMap_(bdd);
    if (aplicarLista_(bdd, col_(cm, 'estado'), ESTADOS.BDD)) hechos.push('BDD.Estado');
    if (aplicarLista_(bdd, col_(cm, 'estado mails'), ESTADOS.MAILS)) hechos.push('BDD.EstadoMails');
    if (aplicarLista_(bdd, col_(cm, 'google calendar'), ESTADOS.CALENDAR)) hechos.push('BDD.Calendar');
    COLS_PL_BDD.forEach(function (pl) { aplicarLista_(bdd, col_(cm, pl.toLowerCase()), ESTADOS.PL_SI); });
    hechos.push('BDD.PL2-PL6');

    const camp = getHojaFlexible_(ss, 'CAMPINGS');
    const cCamping = col_(cm, 'camping');
    if (camp && cCamping) {
      const filas = Math.max(camp.getMaxRows() - 1, 1);
      const rangoCampings = camp.getRange(2, 1, filas, 1);
      try {
        const regla = SpreadsheetApp.newDataValidation()
          .requireValueInRange(rangoCampings, true).setAllowInvalid(false).build();
        bdd.getRange(2, cCamping, Math.max(bdd.getMaxRows() - 1, 1), 1).setDataValidation(regla);
        hechos.push('BDD.Camping (desde Campings)');
      } catch (e) { /* ignora */ }
    }
  }
  return 'Configurados: ' + (hechos.join(', ') || 'ninguno');
}

function aplicarLista_(hoja, col, lista) {
  if (!col) return false;
  try {
    const regla = SpreadsheetApp.newDataValidation().requireValueInList(lista, true).setAllowInvalid(true).build();
    hoja.getRange(2, col, Math.max(hoja.getMaxRows() - 1, 1), 1).setDataValidation(regla);
    return true;
  } catch (e) { return false; }
}

function ocultarColumnasTecnicas_(ss) {
  const hechos = [];
  const form = getHojaFlexible_(ss, 'FORMULARIO');
  if (form) {
    const cm = buildColMap_(form);
    COLS_TECNICAS_FORM.forEach(function (h) {
      const c = col_(cm, h.toLowerCase());
      if (c) try { form.hideColumns(c); } catch (e) {}
    });
    hechos.push('Formulario (técnicas)');
  }
  const bdd = getHojaFlexible_(ss, 'BDD');
  if (bdd) {
    const cm = buildColMap_(bdd);
    COLS_PL_BDD.forEach(function (pl) {
      const c = col_(cm, pl.toLowerCase());
      if (c) try { bdd.hideColumns(c); } catch (e) {}
    });
    hechos.push('BDD (PL2-PL6)');
  }
  return 'Ocultas: ' + (hechos.join(', ') || 'ninguna');
}


/* ============================================================ */
/* ===============  FASE 1: FORMULARIO -> PL-01  ============== */
/* ============================================================ */

// (Ajuste 8) Instala el disparador y explica claramente sus límites.
function instalarTriggers_PL1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onFormSubmitTrueque') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmitTrueque').forSpreadsheet(ss).onFormSubmit().create();

  let formInfo = '';
  let sinForm = false;
  try {
    const f = getHojaFlexible_(ss, 'FORMULARIO');
    const url = f ? f.getFormUrl() : null;
    if (url) {
      formInfo = '\n\nFormulario vinculado detectado ✓';
    } else {
      sinForm = true;
      formInfo = '\n\n⚠ No detecto un Google Form vinculado a "Formulario". El disparador automático solo se activa con envíos reales del formulario; mientras tanto usa la opción 3 del menú.';
    }
  } catch (e) {}

  SpreadsheetApp.getUi().alert(
    'Automatización instalada.\n\n' +
    'IMPORTANTE: "onFormSubmitTrueque" SOLO se ejecuta con RESPUESTAS REALES del formulario vinculado a esta hoja.\n' +
    'Para datos que ya existían o filas pegadas a mano, usa:\n' +
    '⛺ Trueque > 3) Procesar solicitudes pendientes (PL-01).' +
    formInfo
  );
  registrarLog_('Instalar trigger', 'Formulario', '-', 'PL-01', 'OK',
    'Trigger onFormSubmit instalado' + (sinForm ? ' (sin form vinculado detectado)' : ''));
}

// Se dispara solo con cada nueva respuesta real del formulario.
function onFormSubmitTrueque(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = getHojaFlexible_(ss, 'FORMULARIO');
  if (!hoja) return;
  let row;
  if (e && e.range && e.range.getSheet().getName() === hoja.getName()) row = e.range.getRow();
  else row = hoja.getLastRow();
  if (row < 2) return;
  procesarFilaFormulario_(ss, hoja, row, null);
}

// Recorre el formulario y prepara PL-01 para lo que falte. Sirve para procesar
// datos previos y para PROBAR sin formulario.
function procesarSolicitudesPendientesPL1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const hoja = getHojaFlexible_(ss, 'FORMULARIO');
  if (!hoja) { ui.alert('No se encontró la hoja "Formulario".'); return; }

  const last = hoja.getLastRow();
  if (last < 2) { ui.alert('No hay solicitudes en el formulario.'); return; }

  const plantilla = getPlantilla_(ss, 'PL-01');
  if (!plantilla.encontrada) {
    ui.alert('No encuentro la plantilla PL-01.\nRevisa que en "Plantillas_Email" haya una fila con ID = PL-01, ' +
             'con el cuerpo en la columna "Mail" y el asunto en "Asunto".');
    return;
  }
  // (Ajuste 4) Cuerpo vacío -> no generamos nada.
  if (String(plantilla.cuerpo || '').trim() === '') {
    ui.alert('La plantilla PL-01 existe pero su cuerpo (columna "Mail") está VACÍO.\n' +
             'Rellénalo antes de generar borradores.');
    return;
  }

  const cont = { creado: 0, falta_email: 0, incompleta: 0, omitido: 0, plantilla_vacia: 0, error: 0, vacia: 0 };
  for (let row = 2; row <= last; row++) {
    const r = procesarFilaFormulario_(ss, hoja, row, plantilla);
    if (cont[r] !== undefined) cont[r]++;
  }
  ui.alert('Proceso PL-01 terminado:\n\n' +
    '• Borradores creados: ' + cont.creado + '\n' +
    '• Sin email (marcadas "Falta email"): ' + cont.falta_email + '\n' +
    '• Incompletas (sin propuesta, no procesadas): ' + cont.incompleta + '\n' +
    '• Omitidas (ya gestionadas): ' + cont.omitido + '\n' +
    '• Plantilla vacía: ' + cont.plantilla_vacia + '\n' +
    '• Errores: ' + cont.error + '\n\n' +
    'Los borradores quedan en Gmail para que los revises y los envíes tú.');
}

// Núcleo idempotente: prepara el borrador PL-01 para UNA fila.
function procesarFilaFormulario_(ss, hoja, row, plantillaOpt) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (err) { return 'error'; }
  try {
    const cm = buildColMap_(hoja);
    const cId      = col_(cm, 'id_formulario');
    const cFechaP1 = col_(cm, 'fecha_pl1');
    const cDraft   = col_(cm, 'draft_pl1_id');
    const cPl1     = col_(cm, 'pl1');
    const cEmail   = col_(cm, 'email');
    const cNombre  = col_(cm, 'nombre');
    const cTalento = col_(cm, 'talento');
    const cDestino = col_(cm, 'destino');
    const cDesc    = col_(cm, 'descripcion');

    const vals = hoja.getRange(row, 1, 1, hoja.getLastColumn()).getValues()[0];
    const get = function (c) { return c ? vals[c - 1] : ''; };

    const nombreV  = String(get(cNombre) || '').trim();
    const emailV   = String(get(cEmail) || '').trim();
    const talentoV = String(get(cTalento) || '').trim();
    const destinoV = String(get(cDestino) || '').trim();
    const descV    = String(get(cDesc) || '').trim();

    // (Ajuste 3) Fila totalmente vacía -> ignorar en silencio.
    if (!nombreV && !emailV) return 'vacia';
    // (Ajuste 3) Hay nombre/email pero ningún campo de propuesta -> probablemente
    // prueba incompleta. No creamos borrador.
    if (!talentoV && !destinoV && !descV) return 'incompleta';

    // Asegurar ID_Formulario.
    let idForm = String(get(cId) || '').trim();
    if (!idForm && cId) { idForm = nextIdFormulario_(hoja, cId); hoja.getRange(row, cId).setValue(idForm); }

    // ¿Ya gestionado? -> no duplicar. ("Borrador creado" = gestionado, no enviado.)
    const pl1 = String(get(cPl1) || '').trim().toLowerCase();
    if (pl1 === 'borrador creado' || pl1 === 'enviado' || String(get(cDraft) || '').trim()) return 'omitido';

    // Email válido.
    if (!esEmailValido_(emailV)) {
      if (cPl1) hoja.getRange(row, cPl1).setValue('Falta email');
      registrarLog_('Preparar PL-01', hoja.getName(), idForm || ('fila ' + row), 'PL-01', 'Falta email', 'Sin email válido');
      return 'falta_email';
    }

    // Plantilla (se lee en vivo desde la hoja).
    const plantilla = plantillaOpt || getPlantilla_(ss, 'PL-01');
    if (!plantilla.encontrada) {
      if (cPl1) hoja.getRange(row, cPl1).setValue('Error: plantilla PL-01');
      registrarLog_('Preparar PL-01', hoja.getName(), idForm, 'PL-01', 'Error', 'PL-01 no encontrada en Plantillas_Email');
      return 'error';
    }
    // (Ajuste 4) Cuerpo vacío -> no crear borrador vacío.
    if (String(plantilla.cuerpo || '').trim() === '') {
      if (cPl1) hoja.getRange(row, cPl1).setValue('Error: plantilla vacía');
      registrarLog_('Preparar PL-01', hoja.getName(), idForm, 'PL-01', 'Error', 'El cuerpo (columna "Mail") de PL-01 está vacío');
      return 'plantilla_vacia';
    }

    // Variables y render.
    const tz = ss.getSpreadsheetTimeZone();
    const vars = buildVarsFormulario_(vals, cm, tz, leerFirma_(ss));
    const asunto = renderPlantilla_(plantilla.asunto, vars) || 'Tu propuesta · Kampaoh Trueque';
    const cuerpo = renderPlantilla_(plantilla.cuerpo, vars);

    // Crear BORRADOR (nunca se envía).
    try {
      const draft = crearBorradorGmail_(emailV, asunto, cuerpo);
      if (cDraft)   hoja.getRange(row, cDraft).setValue(draft.getId());
      if (cFechaP1) hoja.getRange(row, cFechaP1).setValue(new Date());
      if (cPl1)     hoja.getRange(row, cPl1).setValue('Borrador creado');
      registrarLog_('Preparar PL-01', hoja.getName(), idForm, 'PL-01', 'Borrador creado', 'Destinatario: ' + emailV);
      return 'creado';
    } catch (err) {
      if (cPl1) hoja.getRange(row, cPl1).setValue('Error al crear borrador');
      registrarLog_('Preparar PL-01', hoja.getName(), idForm, 'PL-01', 'Error', String(err));
      return 'error';
    }
  } finally {
    lock.releaseLock();
  }
}

function nextIdFormulario_(hoja, cId) {
  const last = hoja.getLastRow();
  let max = 0;
  if (last >= 2) {
    const vals = hoja.getRange(2, cId, last - 1, 1).getValues();
    vals.forEach(function (r) {
      const m = String(r[0] || '').match(/FORM-(\d+)/i);
      if (m) { const n = parseInt(m[1], 10); if (n > max) max = n; }
    });
  }
  return 'FORM-' + String(max + 1).padStart(3, '0');
}


/* ============================================================ */
/* ==================  PRUEBA SEGURA (Ajuste 10)  ============ */
/* ============================================================ */

// Lee PL-01 y muestra asunto/cuerpo renderizado con datos FICTICIOS.
// NO crea ningún borrador ni toca ninguna fila.
function testLeerPlantillaPL01() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const plantilla = getPlantilla_(ss, 'PL-01');
  if (!plantilla.encontrada) {
    ui.alert('PL-01 no encontrada en "Plantillas_Email" (busca una fila con ID = PL-01).');
    return;
  }
  if (String(plantilla.cuerpo || '').trim() === '') {
    ui.alert('PL-01 existe pero su cuerpo (columna "Mail") está VACÍO.');
    return;
  }

  const datos = {
    nombre: 'Laura Pruebas', artista: 'Laura Pruebas',
    talento: 'Taller de cerámica', actividad: 'Taller de cerámica',
    descripcion: 'Taller creativo para familias',
    destino: 'palmar', camping: 'palmar',
    fecha: '22/06/2026', fecha_solicitud: '22/06/2026',
    llegada: '05/07/2026', salida: '07/07/2026',
    dias_actuacion: '2',
    email: 'laura.pruebas@example.com', telefono: '600000000',
    instagram_web: '@laura_ceramica', instagram: '@laura_ceramica', web: '@laura_ceramica',
    video: '', notas: '(datos de prueba)',
    firma: leerFirma_(ss)
  };

  const asunto = renderPlantilla_(plantilla.asunto, datos) || '(sin asunto)';
  const cuerpoRender = renderPlantilla_(plantilla.cuerpo, datos);
  const cuerpoTexto = /<[a-z][\s\S]*>/i.test(cuerpoRender) ? htmlAPlano_(cuerpoRender) : cuerpoRender;

  Logger.log('ASUNTO PL-01: ' + asunto);
  Logger.log('CUERPO PL-01:\n' + cuerpoTexto);

  const preview = cuerpoTexto.length > 1200
    ? cuerpoTexto.substring(0, 1200) + '\n\n[...recortado: cuerpo completo en Apps Script > Registros de ejecución]'
    : cuerpoTexto;

  ui.alert('PRUEBA PL-01 — no se ha creado ningún borrador\n\n' +
           'ASUNTO:\n' + asunto + '\n\n' +
           'CUERPO:\n' + preview);
}


/* ============================================================ */
/* =====================  HELPERS COMUNES  =================== */
/* ============================================================ */

function getHojaFlexible_(ss, key) {
  const n = NOMBRES[key];
  if (!n) return null;
  return buscarHoja_(ss, n.limpio) || buscarHoja_(ss, n.simple);
}

function buscarHoja_(ss, name) {
  const hs = ss.getSheets();
  for (let i = 0; i < hs.length; i++) {
    if (hs[i].getName().trim() === String(name).trim()) return hs[i];
  }
  return null;
}

function normalizeHeader_(s) {
  return String(s || '').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

function buildColMap_(hoja) {
  const map = new Map();
  const ult = Math.max(hoja.getLastColumn(), 1);
  const headers = hoja.getRange(1, 1, 1, ult).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    const h = normalizeHeader_(headers[i]);
    if (h && !map.has(h)) map.set(h, i + 1);
  }
  return map;
}

// Devuelve la columna por cabecera: 1) coincidencia exacta, 2) "contiene".
function col_(map, key) {
  const kn = normalizeHeader_(key);
  if (map.has(kn)) return map.get(kn);
  let res = 0;
  map.forEach(function (idx, h) { if (!res && h.indexOf(kn) !== -1) res = idx; });
  return res;
}

// Lee una plantilla por su ID (p.ej. "PL-01"). Cuerpo = "Mail"; Asunto = "Asunto".
function getPlantilla_(ss, id) {
  const hoja = getHojaFlexible_(ss, 'PLANTILLAS');
  if (!hoja) return { encontrada: false };
  const cm = buildColMap_(hoja);
  const cId     = col_(cm, 'id');
  const cCuerpo = col_(cm, 'mail');
  const cAsunto = col_(cm, 'asunto');
  const cDest   = col_(cm, 'destinatario');
  const last = hoja.getLastRow();
  if (last < 2 || !cId) return { encontrada: false };

  const vals = hoja.getRange(2, 1, last - 1, hoja.getLastColumn()).getValues();
  const norm = function (s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); };
  const objetivo = norm(id);
  for (let i = 0; i < vals.length; i++) {
    if (norm(vals[i][cId - 1]) === objetivo) {
      return {
        encontrada: true,
        asunto: String(cAsunto ? vals[i][cAsunto - 1] : ''),
        cuerpo: String(cCuerpo ? vals[i][cCuerpo - 1] : ''),
        destinatario: String(cDest ? vals[i][cDest - 1] : '')
      };
    }
  }
  return { encontrada: false };
}

// (Ajuste 6) Sustituye {variable} y {{variable}} (sin distinguir mayúsculas).
function renderPlantilla_(texto, vars) {
  let out = String(texto || '');
  Object.keys(vars).forEach(function (k) {
    const re = new RegExp('\\{\\{?\\s*' + k + '\\s*\\}?\\}', 'gi');
    out = out.replace(re, vars[k] == null ? '' : String(vars[k]));
  });
  return out;
}

// (Ajuste 6) Variables disponibles + alias frecuentes:
//   {artista}=={nombre}, {camping}=={destino}, {instagram} y {web} == {instagram_web}
function buildVarsFormulario_(vals, cm, tz, firma) {
  const g = function (key) { const c = col_(cm, key); return c ? vals[c - 1] : ''; };
  const nombre        = String(g('nombre') || '').trim();
  const talento       = String(g('talento') || '').trim();
  const destino       = String(g('destino') || '').trim();
  const instagram_web = String(g('instagram') || '').trim();
  return {
    nombre: nombre,
    artista: nombre,
    talento: talento,
    actividad: talento,
    descripcion: String(g('descripcion') || '').trim(),
    destino: destino,
    camping: destino,
    fecha: formatFecha_(g('fecha'), tz),
    fecha_solicitud: formatFecha_(g('fecha'), tz),
    llegada: formatFecha_(g('llegada'), tz),
    salida: formatFecha_(g('salida'), tz),
    dias_actuacion: String(g('dias actuacion') || '').trim(),
    email: String(g('email') || '').trim(),
    telefono: String(g('telefono') || '').trim(),
    instagram_web: instagram_web,
    instagram: instagram_web,
    web: instagram_web,
    video: String(g('video') || '').trim(),
    notas: String(g('notas') || '').trim(),
    firma: firma
  };
}

function formatFecha_(value, tz) {
  if (value instanceof Date) return Utilities.formatDate(value, tz || 'Europe/Madrid', 'dd/MM/yyyy');
  return String(value || '').trim();
}

function esEmailValido_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

// Crea un BORRADOR en Gmail (texto plano limpio + versión HTML). Nunca envía.
function crearBorradorGmail_(to, subject, body) {
  const esHtml = /<[a-z][\s\S]*>/i.test(body);
  const html = esHtml ? body : escaparYBr_(body);
  const plano = esHtml ? htmlAPlano_(body) : body;   // (Ajuste 5)
  return GmailApp.createDraft(to, subject, plano, { htmlBody: html });
}

function escaparYBr_(t) {
  return String(t || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
}

// (Ajuste 5) Conversión HTML -> texto plano legible (saltos de línea reales,
// viñetas para <li>, entidades decodificadas, sin líneas en blanco de más).
function htmlAPlano_(html) {
  let t = String(html || '');
  t = t.replace(/<\s*br\s*\/?>/gi, '\n');
  t = t.replace(/<\s*li[^>]*>/gi, '\n• ');
  t = t.replace(/<\/\s*(p|div|li|tr|h[1-6]|ul|ol)\s*>/gi, '\n');
  t = t.replace(/<[^>]+>/g, '');
  t = t.replace(/&nbsp;/gi, ' ')
       .replace(/&amp;/gi, '&')
       .replace(/&lt;/gi, '<')
       .replace(/&gt;/gi, '>')
       .replace(/&quot;/gi, '"')
       .replace(/&#39;/gi, "'");
  t = t.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n');
  return t.trim();
}

function leerFirma_(ss) {
  try {
    const c = getHojaFlexible_(ss, 'CONFIG');
    if (c) {
      const l1 = c.getRange('B4').getValue();
      const l2 = c.getRange('B5').getValue();
      if (l1 || l2) return [l1, l2].filter(Boolean).join('\n');
    }
  } catch (e) {}
  return FIRMA;
}

// Registra una acción en "Log_Acciones" (crea la hoja si no existe).
function registrarLog_(accion, hoja, ref, plantilla, resultado, detalle) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let log = getHojaFlexible_(ss, 'LOG');
    if (!log) {
      log = ss.insertSheet(NOMBRES.LOG.limpio);
      log.getRange(1, 1, 1, 7)
        .setValues([['Fecha y hora', 'Acción', 'Hoja', 'Referencia', 'Plantilla', 'Resultado', 'Detalle']])
        .setFontWeight('bold');
      log.setFrozenRows(1);
    }
    log.appendRow([new Date(), accion, hoja, ref, plantilla, resultado, detalle]);
  } catch (e) { /* el log nunca debe romper el flujo principal */ }
}


/* ============================================================ */
/* ====  ANTI-DUPLICADOS PL2-PL6 EN BDD (para Fase 2+)  ====== */
/* ====  "Sí" = gestionado / borrador creado (no enviado).     */
/* ============================================================ */

function plMarcadoBDD_(hojaBDD, row, headerPL) {
  const cm = buildColMap_(hojaBDD);
  const c = col_(cm, headerPL.toLowerCase());
  if (!c) return false;
  const v = String(hojaBDD.getRange(row, c).getValue() || '').trim().toLowerCase();
  return v === 'sí' || v === 'si';
}

function marcarPLBDD_(hojaBDD, row, headerPL) {
  const cm = buildColMap_(hojaBDD);
  const c = col_(cm, headerPL.toLowerCase());
  if (c) hojaBDD.getRange(row, c).setValue('Sí');
}
