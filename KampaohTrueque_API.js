// ============================================================
// KAMPAOH TRUEQUE — Apps Script API unificada
//
// UNA SOLA doGet que hace dos cosas:
//   → ?action=getData        → sirve JSON al dashboard web
//   → ?nombre=&email=&…      → recibe envíos del formulario público
//
// Cómo distinguirlos: si llega el parámetro "action=getData"
// es el dashboard. Cualquier otra llamada es un envío de formulario.
// ============================================================

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};

  const jsonOut = d => ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON);

  // ── RAMA A: Dashboard ─────────────────────────────────────
  if (params.action === 'getData')      { try { return jsonOut(buildDashboardData_()); }    catch(e){ return jsonOut({error:e.toString()}); } }
  if (params.action === 'getPreview')   { try { return jsonOut(buildEmail_(params.id, params.pl)); } catch(e){ return jsonOut({ok:false,error:e.toString()}); } }
  if (params.action === 'sendMail')     { try { return jsonOut(emailSend_(params.id, params.pl)); }  catch(e){ return jsonOut({ok:false,error:e.toString()}); } }
  if (params.action === 'updateEstado')     { try { return jsonOut(estadoUpdate_(params.id, params.estado)); }     catch(e){ return jsonOut({ok:false,error:e.toString()}); } }
  if (params.action === 'updateEstadoProc') { try { return jsonOut(estadoProcUpdate_(params.id, params.estado)); } catch(e){ return jsonOut({ok:false,error:e.toString()}); } }
  if (params.action === 'colorConflictos')  { try { return jsonOut(colorConflictos_()); } catch(e){ return jsonOut({ok:false,error:e.toString()}); } }
  if (params.action === 'updateCalGuest')   { try { return jsonOut(api_updateCalGuests_(params.id)); } catch(e){ return jsonOut({ok:false,error:e.toString()}); } }
  if (params.action === 'respuesta')       { try { return registrarRespuesta_(params.id, params.resp); } catch(e){ return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;color:red">Error: '+e.toString()+'</p>'); } }
  if (params.action === 'valoracion')     { try { return registrarValoracion_(params.id, params.stars); } catch(e){ return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;color:red">Error: '+e.toString()+'</p>'); } }
  if (params.action === 'debugConflictos') { try {
    var ss3 = SpreadsheetApp.getActiveSpreadsheet();
    var bdd3 = ss3.getSheetByName('BDD_Trueque');
    var hdrs3 = api_hdrs_(bdd3);
    var ci3 = function(n){ return hdrs3.indexOf(n); };
    var data3 = bdd3.getRange(2,1,bdd3.getLastRow()-1,hdrs3.length).getValues();
    var cancelados3 = ['Rechaza','Cancelado','Cancelado STOP'];
    var filas = data3.filter(function(r){ return r[ci3('ID')]; }).map(function(r){
      return {
        id: r[ci3('ID')].toString(),
        nombre: r[ci3('Nombre')] || '',
        camping: r[ci3('Camping')] || '',
        llegada_raw: r[ci3('Llegada')] ? r[ci3('Llegada')].toString() : '',
        llegada_fmt: fmtFechaAPI_(r[ci3('Llegada')]),
        salida_fmt:  fmtFechaAPI_(r[ci3('Salida')]),
        estado: r[ci3('Estado')] || '',
        cancelado: cancelados3.indexOf((r[ci3('Estado')]||'').toString()) >= 0
      };
    });
    var conflictos = [];
    filas.forEach(function(a, i){
      if (!a.camping || !a.llegada_fmt || a.cancelado) return;
      filas.forEach(function(b, j){
        if (j <= i) return;
        if (b.cancelado) return;
        if (a.camping === b.camping && a.llegada_fmt === b.llegada_fmt) {
          conflictos.push({ id_a: a.id, nombre_a: a.nombre, id_b: b.id, nombre_b: b.nombre,
            camping: a.camping, llegada: a.llegada_fmt });
        }
      });
    });
    return jsonOut({ filas: filas, conflictos: conflictos });
  } catch(e){ return jsonOut({error:e.toString()}); } }

  if (params.action === 'debugCamping') { try {
    var ss2 = SpreadsheetApp.getActiveSpreadsheet();
    var sh2 = ss2.getSheetByName('Campings');
    var h2  = sh2 ? api_hdrs_(sh2) : [];
    var r2  = (sh2 && sh2.getLastRow() > 1) ? sh2.getRange(2,1,sh2.getLastRow()-1,Math.max(h2.length,1)).getValues() : [];
    var cc  = h2.indexOf('Camping');
    var nombres = r2.map(function(r){ return r[cc]; });
    var bdd2  = ss2.getSheetByName('BDD_Trueque');
    var bh2   = bdd2 ? api_hdrs_(bdd2) : [];
    var brows = (bdd2 && bdd2.getLastRow() > 1) ? bdd2.getRange(2,1,bdd2.getLastRow()-1,Math.max(bh2.length,1)).getValues() : [];
    var campingsEnBDD = brows.map(function(r){ var i=bh2.indexOf('Camping'); return i>=0?r[i]:''; }).filter(function(v){ return v; });
    return jsonOut({ campingsSheet: nombres, campingsEnBDD: campingsEnBDD });
  } catch(e){ return jsonOut({error:e.toString()}); } }

  // ── RAMA B: Formulario público envía una solicitud ────────
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Formulario') || ss.getActiveSheet();

    // Crear cabeceras si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Fecha','Talento','Descripción','Destino','Llegada','Salida',
        'Días actuación','Nombre','Email','Teléfono','Instagram/Web',
        'Vídeo','Notas extra','Estado','DNI','Acompañante'
      ]);
    }

    sheet.appendRow([
      params.fecha        || new Date().toISOString(),
      params.talento      || '',
      params.descripcion  || '',
      params.destinos     || '',
      params.llegada      || '',
      params.salida       || '',
      params.dias         || '',
      params.nombre       || '',
      params.email        || '',
      params.telefono     || '',
      params.igweb        || '',
      params.video        || '',
      params.extra        || '',
      'Pendiente',
      params.dni          || '',
      params.acompanante  || ''
    ]);

    // Enviar email de confirmación al artista (solo si tiene email)
    if (params.email) {
      sendConfirmacionArtista_(params.nombre || '', params.email);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Email de confirmación al artista ─────────────────────────
// (tu función original, renombrada con _ para uso interno)

function sendConfirmacionArtista_(nombre, email) {
  const html = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:sans-serif;background:#f4f4f4;margin:0;padding:0;">
<div style="max-width:580px;margin:32px auto;background:#fff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.1);">
  <div style="background:#111;padding:36px 40px;text-align:center;">
    <div style="font-size:1.8rem;margin-bottom:12px;">💛</div>
    <h1 style="color:#F5C800;font-size:1.4rem;margin:0 0 6px;">Kampaoh Trueque</h1>
    <p style="color:rgba(255,255,255,.5);font-size:.85rem;margin:0;">Programa 2026</p>
  </div>
  <div style="padding:36px 40px;">
    <p style="color:#333;font-size:.97rem;line-height:1.7;">Hola${nombre ? ', ' + nombre : ''} 💛</p>
    <p style="color:#555;font-size:.93rem;line-height:1.75;">Muchas gracias por tu interés en participar en <strong>Kampaoh Trueque</strong>.</p>
    <p style="color:#555;font-size:.93rem;line-height:1.75;">Hemos recibido correctamente tu formulario y queremos agradecerte el tiempo dedicado a compartir tu propuesta con nosotros.</p>
    <p style="color:#555;font-size:.93rem;line-height:1.75;">Kampaoh Trueque nace con la idea de ofrecer a nuestros huéspedes experiencias diferentes y únicas durante su estancia, por lo que estudiaremos cada propuesta con detalle.</p>
    <p style="color:#555;font-size:.93rem;line-height:1.75;">Durante los próximos días revisaremos toda la información recibida y valoraremos tu perfil.</p>
    <p style="color:#555;font-size:.93rem;line-height:1.75;">En caso de seguir adelante con la propuesta, recibirás un nuevo email con más información sobre los próximos pasos a seguir.</p>
    <p style="color:#555;font-size:.93rem;line-height:1.75;">Gracias de nuevo por tu interés y por querer formar parte de esta iniciativa.</p>
    <p style="color:#333;font-size:.93rem;margin-top:24px;">Un saludo,<br><strong>Equipo Kampaoh</strong> 💛</p>
  </div>
  <div style="background:#f0f0f0;padding:16px 40px;text-align:center;font-size:.75rem;color:#aaa;">
    &copy; 2026 Kampaoh &middot; Programa Trueque &middot;
    <a href="https://es.kampaoh.com" style="color:#aaa;text-decoration:none;">es.kampaoh.com</a>
  </div>
</div>
</body>
</html>`;

  GmailApp.sendEmail(email, 'Hemos recibido tu solicitud — Kampaoh Trueque', '', {
    htmlBody: html,
    name:    'Kampaoh Trueque',
    from:    'hola@kampaoh.com',
    replyTo: 'hola@kampaoh.com'
  });
}

// ── Dashboard: construir payload JSON ────────────────────────

function buildDashboardData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetForm = ss.getSheetByName('Formulario');
  const sheetBDD  = ss.getSheetByName('BDD_Trueque');
  const sheetLog  = ss.getSheetByName('Log_Acciones');

  // KPIs desde Formulario
  let totalSolicitudes = 0, pendientes = 0, aprobadas = 0, pl01Creados = 0;

  if (sheetForm && sheetForm.getLastRow() > 1) {
    const fHeaders = sheetForm.getRange(1, 1, 1, sheetForm.getLastColumn()).getValues()[0];
    const formData = sheetForm.getRange(2, 1, sheetForm.getLastRow() - 1, sheetForm.getLastColumn()).getValues();
    const estadoIdx = fHeaders.indexOf('Estado');
    const pl1Idx    = fHeaders.indexOf('PL1');

    formData.forEach(row => {
      if (!row[0]) return;
      totalSolicitudes++;
      const estado = estadoIdx >= 0 ? row[estadoIdx] : '';
      const pl1    = pl1Idx    >= 0 ? row[pl1Idx]    : '';
      if (!estado || estado === 'Pendiente') pendientes++;
      if (estado === 'Sí') aprobadas++;
      if (pl1 === 'Borrador creado' || pl1 === 'Enviado') pl01Creados++;
    });
  }

  // KPIs + filas desde BDD_Trueque
  let registrosBDD = 0;
  let pl02p = 0, pl03p = 0, pl04p = 0, pl05p = 0, pl06p = 0;
  let calendarCreado = 0;
  const rows = [];

  if (sheetBDD && sheetBDD.getLastRow() > 1) {
    const bddHeaders = sheetBDD.getRange(1, 1, 1, sheetBDD.getLastColumn()).getValues()[0];
    const bddData    = sheetBDD.getRange(2, 1, sheetBDD.getLastRow() - 1, sheetBDD.getLastColumn()).getValues();
    // Búsqueda de columna insensible a mayúsculas
    const ci = function(name) {
      var lower = name.toLowerCase();
      for (var _i = 0; _i < bddHeaders.length; _i++) {
        if (bddHeaders[_i].toString().toLowerCase() === lower) return _i;
      }
      return -1;
    };
    // La columna puede llamarse "Estado email" o "Estado Mails" (nombre antiguo)
    var estadoEmailIdx = ci('Estado email') >= 0 ? ci('Estado email') : ci('Estado Mails');

    // Cargar PropertiesService una sola vez como capa de persistencia
    var sentProps = {};
    try { sentProps = PropertiesService.getScriptProperties().getProperties(); } catch(ep) {}

    bddData.forEach(function(row) {
      const id = row[ci('ID')];
      if (!id) return;
      registrosBDD++;
      var idStr = id.toString();

      const estadoEmail = (estadoEmailIdx >= 0 ? row[estadoEmailIdx] : '') || '';
      const estadoRaw   = (row[ci('Estado')] || '').toString().trim();
      const estadoProc  = estadoRaw === 'Cancelada' ? 'Cancelado' : estadoRaw === 'Cancelada STOP' ? 'Cancelado STOP' : estadoRaw;
      const formulario  = row[ci('Formulario')]       || '';
      const pl2 = row[ci('PL2')]             || '';
      const pl3 = row[ci('PL3')]             || '';
      const pl4 = row[ci('PL4')]             || '';
      const pl5 = row[ci('PL5')]             || '';
      const pl6 = row[ci('PL6')]             || '';
      // PL7/PL8/PL9: hoja primero, PropertiesService como fallback
      const pl7 = row[ci('PL7')] || sentProps['pl_' + idStr + '_PL7'] || '';
      const pl8 = row[ci('PL8')] || sentProps['pl_' + idStr + '_PL8'] || '';
      const pl9 = row[ci('PL9')] || sentProps['pl_' + idStr + '_PL9'] || '';
      const cal = row[ci('Google Calendar')] || '';

      if (estadoEmail === 'Pendiente'      && pl2 !== 'Sí') pl02p++;
      if (estadoProc  === 'Cancelado'      && pl3 !== 'Sí') pl03p++;
      if (estadoProc  === 'Cancelado STOP' && pl4 !== 'Sí') pl04p++;
      if (estadoProc  === 'En proceso'     && pl5 !== 'Sí') pl05p++;
      if (estadoProc  === 'En proceso'     && pl6 !== 'Sí') pl06p++;
      if (cal === 'Evento creado') calendarCreado++;

      rows.push({
        id:             id.toString(),
        fechaSolicitud: fmtFechaAPI_(row[ci('Solicitud')]),
        nombre:         row[ci('Nombre')]      || '',
        email:          row[ci('Email')]       || '',
        telefono:       row[ci('Teléfono')]    || '',
        dni:            row[ci('DNI')]         || '',
        personas:       row[ci('Nº Personas')] || '',
        actividad:      row[ci('Actividad')]   || '',
        propuesta:      row[ci('Propuesta')]   || '',
        camping:        row[ci('Camping')]     || '',
        llegada:        fmtFechaAPI_(row[ci('Llegada')]),
        salida:         fmtFechaAPI_(row[ci('Salida')]),
        localizador:    row[ci('Localizador')] || '',
        calendar:       cal,
        estado:         estadoProc,
        estadoEmail:    estadoEmail,
        formulario:     formulario,
        pl2, pl3, pl4, pl5, pl6, pl7, pl8, pl9,
        valoracion:     (row[ci('Valoracion')] || sentProps['val_' + idStr] || '').toString(),
        conflicto:      false
      });
    });

    // Conflicto: mismo camping + misma fecha de llegada (solo filas activas)
    var cancelados = ['Rechaza', 'No acepta', 'Cancelado', 'Cancelado STOP'];
    rows.forEach(function(rowObj, i) {
      if (!rowObj.camping || !rowObj.llegada) return;
      if (cancelados.indexOf(rowObj.estado) >= 0) return;
      for (var j = 0; j < rows.length; j++) {
        if (j === i) continue;
        if (cancelados.indexOf(rows[j].estado) >= 0) continue;
        if (rows[j].camping === rowObj.camping && rows[j].llegada === rowObj.llegada) {
          rowObj.conflicto = true;
          break;
        }
      }
    });
  }

  // Errores desde Log
  let errores = 0;
  if (sheetLog && sheetLog.getLastRow() > 1) {
    const logHeaders = sheetLog.getRange(1, 1, 1, sheetLog.getLastColumn()).getValues()[0];
    const resIdx = logHeaders.indexOf('Resultado');
    if (resIdx >= 0) {
      errores = sheetLog
        .getRange(2, resIdx + 1, sheetLog.getLastRow() - 1, 1)
        .getValues().flat()
        .filter(v => v === 'Error').length;
    }
  }

  return {
    kpis: {
      totalSolicitudes, pendientes, aprobadas,
      registrosBDD, pl01Creados,
      pl02Pendientes: pl02p, pl03Pendientes: pl03p,
      pl04Pendientes: pl04p, pl05Pendientes: pl05p,
      pl06Pendientes: pl06p, calendarCreado, errores
    },
    rows
  };
}

// ── Dashboard: Preview y Envío real de emails ────────────────
//  Estas funciones son completamente autónomas (no dependen del
//  script principal). Funcionan solas en cualquier proyecto GAS.

function buildEmail_(trqId, plId) {
  if (!trqId || !plId) throw new Error('Faltan parámetros id y pl');

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const bdd = ss.getSheetByName('BDD_Trueque');
  if (!bdd || bdd.getLastRow() <= 1) throw new Error('BDD vacía o no encontrada');

  const headers = api_hdrs_(bdd);
  const idCol   = headers.indexOf('ID');
  if (idCol < 0) throw new Error('Columna ID no encontrada en BDD_Trueque');

  const allRows = bdd.getRange(2, 1, bdd.getLastRow() - 1, headers.length).getValues();
  const rowData = allRows.find(function(r){ return r[idCol] && r[idCol].toString() === trqId; });
  if (!rowData) throw new Error('Registro ' + trqId + ' no encontrado');

  const plt = api_plt_(plId);
  if (!plt || !plt.mail) throw new Error('Plantilla ' + plId + ' no encontrada. Comprueba la hoja Plantillas_Email.');

  var get = function(n){ var i = headers.indexOf(n); return i >= 0 ? (rowData[i] || '').toString() : ''; };
  var campingNom  = get('Camping');
  var campingData = api_camping_(campingNom);
  var vars        = api_vars_(rowData, headers, campingData);

  var rAsunto = api_render_(plt.asunto, vars);
  var rCuerpo = api_render_(plt.mail,   vars);

  var to, toName;
  if (plId === 'PL-06' || plId === 'PL-08') {
    if (!campingData) throw new Error('Camping "' + campingNom + '" no encontrado en la hoja Campings');
    var rawEmail = (campingData['Email'] || '').toString();
    var emails   = rawEmail.split(/[\n,;]+/).map(function(s){ return s.trim(); }).filter(function(s){ return s.indexOf('@') >= 0; });
    if (!emails.length) throw new Error('El camping "' + campingNom + '" no tiene email configurado');
    to     = emails.join(',');
    toName = campingNom;
  } else {
    to     = get('Email');
    toName = get('Nombre');
    if (!to) throw new Error('El artista ' + trqId + ' no tiene email registrado en BDD');
  }

  var bodyText = rCuerpo.texto;
  var bodyHtml = '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.75;color:#333;">' +
    bodyText.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>') + '</div>';

  // PL-02: inyectar botones de respuesta
  if (plId === 'PL-02') {
    var apiUrl = ScriptApp.getService().getUrl();
    var urlSi  = apiUrl + '?action=respuesta&id=' + trqId + '&resp=si';
    var urlNo  = apiUrl + '?action=respuesta&id=' + trqId + '&resp=no';
    bodyHtml += '<div style="text-align:center;margin:32px 0 16px">' +
      '<a href="' + urlSi + '" style="display:inline-block;background:#059669;color:#fff;padding:13px 30px;border-radius:8px;text-decoration:none;font-weight:700;font-size:15px;margin:0 8px">✅ Sí, acepto el trueque</a>' +
      '<a href="' + urlNo + '" style="display:inline-block;background:#DC2626;color:#fff;padding:13px 30px;border-radius:8px;text-decoration:none;font-weight:700;font-size:15px;margin:0 8px">❌ No puedo esta vez</a>' +
      '</div>';
  }

  // PL-09: inyectar botones de valoración (1-5 estrellas)
  if (plId === 'PL-09') {
    var apiUrl09 = ScriptApp.getService().getUrl();
    var starColors = ['#EF4444','#F97316','#F59E0B','#84CC16','#059669'];
    var starLabels = ['★','★★','★★★','★★★★','★★★★★'];
    var starBtns = '';
    for (var si = 0; si < 5; si++) {
      var urlStar = apiUrl09 + '?action=valoracion&id=' + trqId + '&stars=' + (si + 1);
      starBtns += '<a href="' + urlStar + '" style="display:inline-block;background:' + starColors[si] + ';color:#fff;padding:12px 16px;border-radius:8px;text-decoration:none;font-weight:700;font-size:15px;margin:0 4px">' + starLabels[si] + '</a>';
    }
    bodyHtml += '<div style="text-align:center;margin:28px 0 16px;padding:24px;background:#FFFBEB;border-radius:12px;border:1px solid #FCD34D">' +
      '<p style="margin-bottom:20px;font-size:15px;font-weight:600;color:#333;">¿Cómo fue tu experiencia? Toca una estrella:</p>' +
      starBtns +
      '</div>';
  }

  var allVarsSin = rAsunto.varsSin.concat(rCuerpo.varsSin).filter(function(v,i,a){ return a.indexOf(v)===i; });

  return { ok: true, trqId: trqId, plId: plId, to: to, toName: toName,
           subject: rAsunto.texto, bodyText: bodyText, bodyHtml: bodyHtml, varsSin: allVarsSin,
           campingDebug: { nombre: campingNom, found: !!campingData } };
}

function emailSend_(trqId, plId) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var bdd = ss.getSheetByName('BDD_Trueque');
  if (!bdd) throw new Error('Hoja BDD_Trueque no encontrada');

  var headers = api_hdrs_(bdd);
  var idCol   = headers.indexOf('ID');
  if (idCol < 0) throw new Error('Columna ID no encontrada');

  var allRows = bdd.getRange(2, 1, Math.max(bdd.getLastRow() - 1, 1), headers.length).getValues();
  var rowIdx  = -1;
  for (var i = 0; i < allRows.length; i++) {
    if (allRows[i][idCol] && allRows[i][idCol].toString() === trqId) { rowIdx = i; break; }
  }
  if (rowIdx < 0) throw new Error('Registro ' + trqId + ' no encontrado');

  var plMap = { 'PL-02':'PL2', 'PL-03':'PL3', 'PL-04':'PL4', 'PL-05':'PL5', 'PL-06':'PL6', 'PL-07':'PL7', 'PL-08':'PL8', 'PL-09':'PL9' };
  var plColName = plMap[plId];

  // Búsqueda de columna case-insensitive
  var ciIdx = function(name) {
    var lo = name.toLowerCase();
    for (var k = 0; k < headers.length; k++) {
      if (headers[k].toString().toLowerCase() === lo) return k;
    }
    return -1;
  };

  if (plColName) {
    var plColIdx = ciIdx(plColName);
    var alreadySent = (plColIdx >= 0 && allRows[rowIdx][plColIdx] === 'Sí');
    if (!alreadySent) {
      try { alreadySent = PropertiesService.getScriptProperties().getProperty('pl_' + trqId + '_' + plColName) === 'Sí'; } catch(ep) {}
    }
    if (alreadySent) {
      return { ok: false, error: plId + ' ya fue enviado (anti-duplicado).', duplicate: true };
    }
  }

  var email = buildEmail_(trqId, plId);

  // Crear BORRADOR (no envío directo) — siempre desde marketing@kampaoh.com
  GmailApp.createDraft(email.to, email.subject, email.bodyText, {
    htmlBody: email.bodyHtml,
    name: 'Kampaoh Trueque',
    from: 'marketing@kampaoh.com'
  });

  var sheetRow = rowIdx + 2;
  if (plColName) {
    var plci = ciIdx(plColName);
    if (plci >= 0) bdd.getRange(sheetRow, plci + 1).setValue('Sí');
    // PropertiesService como capa de persistencia independiente de columnas en hoja
    try { PropertiesService.getScriptProperties().setProperty('pl_' + trqId + '_' + plColName, 'Sí'); } catch(ep) {}
  }
  if (plId === 'PL-02') {
    var emci = headers.indexOf('Estado email') >= 0 ? headers.indexOf('Estado email') : headers.indexOf('Estado Mails');
    if (emci >= 0) bdd.getRange(sheetRow, emci + 1).setValue('Contactado');
  }

  // PL-06: marcar Completado + crear evento en Google Calendar
  if (plId === 'PL-06') {
    var emci06 = headers.indexOf('Estado email') >= 0 ? headers.indexOf('Estado email') : headers.indexOf('Estado Mails');
    if (emci06 >= 0) bdd.getRange(sheetRow, emci06 + 1).setValue('Completado');
    try {
      var calOk = api_createCalEvent_(allRows[rowIdx], headers, email.to);
      if (calOk) {
        var gcCol = headers.indexOf('Google Calendar');
        if (gcCol >= 0) bdd.getRange(sheetRow, gcCol + 1).setValue('Evento creado');
      }
    } catch(calErr) {}
  }

  try {
    var logSheet = ss.getSheetByName('Log_Acciones');
    if (logSheet) logSheet.appendRow([new Date(), Session.getActiveUser().getEmail(),
      'Borrador ' + plId + ' (dashboard)', 'BDD_Trueque', '', '', trqId, plId, email.to, 'OK', email.subject, '', email.varsSin.join(', '), '']);
  } catch(e) {}

  return { ok: true, draft: true, to: email.to, subject: email.subject };
}

function estadoUpdate_(trqId, estado) {
  if (!trqId || !estado) throw new Error('Faltan parámetros id y estado');

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var bdd = ss.getSheetByName('BDD_Trueque');
  if (!bdd) throw new Error('Hoja BDD_Trueque no encontrada');

  var headers = api_hdrs_(bdd);
  var idCol   = headers.indexOf('ID');
  var emCol   = headers.indexOf('Estado email') >= 0 ? headers.indexOf('Estado email') : headers.indexOf('Estado Mails');
  if (idCol < 0 || emCol < 0) throw new Error('Columna ID o Estado email/Estado Mails no encontrada');

  var allRows = bdd.getRange(2, 1, Math.max(bdd.getLastRow() - 1, 1), headers.length).getValues();
  var rowIdx  = -1;
  for (var i = 0; i < allRows.length; i++) {
    if (allRows[i][idCol] && allRows[i][idCol].toString() === trqId) { rowIdx = i; break; }
  }
  if (rowIdx < 0) throw new Error('Registro ' + trqId + ' no encontrado');

  bdd.getRange(rowIdx + 2, emCol + 1).setValue(estado);

  try {
    var logSheet = ss.getSheetByName('Log_Acciones');
    if (logSheet) logSheet.appendRow([new Date(), Session.getActiveUser().getEmail(),
      'Actualizar Estado email (dashboard)', 'BDD_Trueque', '', '', trqId, '', '', 'OK', 'Estado email → ' + estado, '', '', '']);
  } catch(e) {}

  return { ok: true, estado: estado };
}

function estadoProcUpdate_(trqId, estado) {
  if (!trqId) throw new Error('Falta parámetro id');

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var bdd = ss.getSheetByName('BDD_Trueque');
  if (!bdd) throw new Error('Hoja BDD_Trueque no encontrada');

  var headers = api_hdrs_(bdd);
  var idCol   = headers.indexOf('ID');
  var estCol  = headers.indexOf('Estado');
  if (idCol < 0 || estCol < 0) throw new Error('Columna ID o Estado no encontrada');

  var allRows = bdd.getRange(2, 1, Math.max(bdd.getLastRow() - 1, 1), headers.length).getValues();
  var rowIdx  = -1;
  for (var i = 0; i < allRows.length; i++) {
    if (allRows[i][idCol] && allRows[i][idCol].toString() === trqId) { rowIdx = i; break; }
  }
  if (rowIdx < 0) throw new Error('Registro ' + trqId + ' no encontrado');

  bdd.getRange(rowIdx + 2, estCol + 1).setValue(estado);

  try {
    var logSheet = ss.getSheetByName('Log_Acciones');
    if (logSheet) logSheet.appendRow([new Date(), Session.getActiveUser().getEmail(),
      'Actualizar Estado (dashboard)', 'BDD_Trueque', '', '', trqId, '', '', 'OK', 'Estado → ' + estado, '', '', '']);
  } catch(e) {}

  return { ok: true, estado: estado };
}

// ── Helpers autónomos (no dependen del script principal) ─────

function api_hdrs_(sheet) {
  if (!sheet || sheet.getLastRow() === 0) return [];
  return sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
}

function api_plt_(id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Plantillas_Email');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var hdrs    = api_hdrs_(sheet);
  var idCol   = hdrs.indexOf('ID');
  var mailCol = hdrs.indexOf('Mail');
  var subjCol = hdrs.indexOf('Asunto');
  if (idCol < 0) return null;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, hdrs.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][idCol] && rows[i][idCol].toString().trim() === id) {
      return { id: id, mail: api_clean_(mailCol >= 0 ? rows[i][mailCol] : ''), asunto: api_clean_(subjCol >= 0 ? rows[i][subjCol] : '') };
    }
  }
  return null;
}

function api_clean_(t) {
  if (!t || typeof t !== 'string') return t ? t.toString() : '';
  var s = t.trim();
  while (s.indexOf('"""') === 0 && s.lastIndexOf('"""') === s.length - 3) s = s.slice(3, -3).trim();
  while (s.indexOf('""')  === 0 && s.lastIndexOf('""')  === s.length - 2) s = s.slice(2, -2).trim();
  if (s.charAt(0) === '"' && s.charAt(s.length-1) === '"' && s.length > 1) s = s.slice(1, -1).trim();
  return s.replace(/^Asunto:\s*/i,'').replace(/^Cuerpo:\s*/i,'').replace(/ /g,' ').trim();
}

function api_camping_(nombre) {
  if (!nombre) return null;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Campings');
  if (!sheet || sheet.getLastRow() <= 1) return null;
  var hdrs = api_hdrs_(sheet);
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, hdrs.length).getValues();
  var norm = function(s){ return s.toString().toLowerCase().replace(/[áàä]/g,'a').replace(/[éèë]/g,'e').replace(/[íìï]/g,'i').replace(/[óòö]/g,'o').replace(/[úùü]/g,'u').replace(/^kampaoh\s*/,'').trim(); };
  for (var i = 0; i < rows.length; i++) {
    var n = rows[i][hdrs.indexOf('Camping')] || '';
    if (n.toString().toLowerCase() === nombre.toString().toLowerCase() || norm(n) === norm(nombre)) {
      var obj = {};
      hdrs.forEach(function(h, j){ obj[h] = rows[i][j]; });
      return obj;
    }
  }
  return null;
}

function api_fmt_(val) {
  if (!val) return '';
  try {
    var d = val instanceof Date ? val : new Date(val);
    if (isNaN(d)) return val.toString();
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } catch(e) { return val ? val.toString() : ''; }
}

function api_cfgVal_(key) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet || sheet.getLastRow() <= 1) return '';
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) { if (data[i][0] === key) return data[i][1]; }
  } catch(e) {}
  return '';
}

function api_vars_(rowData, headers, campingData) {
  var g = function(n){ var i = headers.indexOf(n); return i >= 0 ? (rowData[i] || '').toString() : ''; };
  var llegada = g('Llegada'), salida = g('Salida');
  var fechaTxt = '';
  if (llegada && salida) fechaTxt = 'del ' + api_fmt_(llegada) + ' al ' + api_fmt_(salida);
  else if (llegada) fechaTxt = api_fmt_(llegada);
  var firma = api_cfgVal_('firma') || 'Equipo Kampaoh 💛';
  var v = {
    id: g('ID'), nombre: g('Nombre'), artista: g('Nombre'), email: g('Email'),
    telefono: g('Teléfono'), dni: g('DNI'), personas: g('Nº Personas'), num_personas: g('Nº Personas'),
    actividad: g('Actividad'), talento: g('Actividad'), propuesta: g('Propuesta'),
    camping: g('Camping'), destino: g('Camping'),
    llegada: api_fmt_(llegada), salida: api_fmt_(salida), fecha: fechaTxt,
    localizador: g('Localizador'), estado: g('Estado'), estado_email: g('Estado email'), formulario: g('Formulario'), firma: firma
  };
  if (campingData) {
    // Horario específico por actividad → fallback a horario general
    var actividad = g('Actividad');
    var horario = '';
    if (actividad) {
      var normAct = actividad.toString().toLowerCase()
        .replace(/[áàä]/g,'a').replace(/[éèë]/g,'e')
        .replace(/[íìï]/g,'i').replace(/[óòö]/g,'o').replace(/[úùü]/g,'u').trim();
      var colAct = '';
      if (normAct === 'yoga' || normAct === 'pilates' || normAct === 'meditacion') {
        colAct = 'Horario_Yoga_Pilates_Meditación';
      } else if (normAct === 'workshop') {
        colAct = 'Horario_ workshop';
      } else if (normAct === 'dj' || normAct === 'actuacion' || normAct === 'musica') {
        colAct = 'Horario_DJ_Actuacion_Musica';
      }
      if (colAct) horario = campingData[colAct] || '';
    }
    if (!horario) horario = campingData['Horario'] || '';

    v.email_camping          = campingData['Email']    || '';
    v.telefono_camping       = campingData['Teléfono'] || '';
    v.franja_horaria         = horario;
    v['franja horaria']      = horario;
    v.horario                = horario;
    v.hora                   = g('Horario') || horario;
    v.espacio                = campingData['Zona']     || '';
    v.zona                   = campingData['Zona']     || '';
    v.espacio_hojacampings   = campingData['Zona']     || '';
    v.notas_camping          = campingData['Notas']    || '';
    v.notas_hojacampings     = campingData['Notas']    || '';
    v.material_camping       = campingData['Notas']    || '';
  }
  return v;
}

function api_createCalEvent_(rowData, headers, campingEmail) {
  var g = function(n){ var i=headers.indexOf(n); return i>=0?(rowData[i]||''):''; };
  var nombre      = g('Nombre').toString();
  var actividad   = g('Actividad').toString();
  var camping     = g('Camping').toString();
  var llegada     = g('Llegada');
  var salida      = g('Salida');
  var horario     = g('Horario').toString();
  var localizador = g('Localizador').toString();

  var d1 = llegada instanceof Date ? llegada : new Date(llegada);
  if (isNaN(d1.getTime())) throw new Error('Fecha de llegada no válida para el evento');

  var d2 = (salida && salida !== '') ? (salida instanceof Date ? salida : new Date(salida)) : d1;
  if (isNaN(d2.getTime())) d2 = d1;

  var title = '⛺ Trueque: ' + nombre + ' · ' + camping;
  var desc  = 'Artista: ' + nombre + '\nActividad: ' + actividad +
              '\nCamping: ' + camping + '\nHorario: ' + horario +
              (localizador ? '\nLocalizador: ' + localizador : '');

  var opts = { description: desc };
  var artistEmail = (function(){ var i = headers.indexOf('Email'); return i >= 0 ? (rowData[i] || '').toString().trim() : ''; })();
  var guests = (campingEmail || '').split(/[\n,;]+/)
    .map(function(e){ return e.trim(); })
    .filter(function(e){ return e.indexOf('@') >= 0; });
  if (artistEmail && artistEmail.indexOf('@') >= 0) guests.unshift(artistEmail);
  if (guests.indexOf('noemirodriguez@kampaoh.com') < 0) guests.push('noemirodriguez@kampaoh.com');
  if (guests.length) { opts.guests = guests.join(','); opts.sendInvites = true; }

  var cal   = CalendarApp.getDefaultCalendar();

  // Evento de días completos de llegada a salida (el horario queda en la descripción)
  var d2end = new Date(d2);
  d2end.setDate(d2end.getDate() + 1); // Google Calendar: end date es exclusivo
  var event = (d2.getTime() > d1.getTime())
    ? cal.createAllDayEvent(title, d1, d2end, opts)
    : cal.createAllDayEvent(title, d1, opts);

  return event.getId() ? true : false;
}

function api_render_(template, vars) {
  if (!template) return { texto: '', varsSin: [] };
  var varsSin = [];
  var texto = template;
  texto = texto.replace(/\{\{([^}]+)\}\}/g, function(m, k){
    var key = k.trim();
    if (vars[key] !== undefined && vars[key] !== '') return vars[key];
    varsSin.push('{{' + key + '}}'); return m;
  });
  texto = texto.replace(/\{([^}]+)\}/g, function(m, k){
    var key = k.trim();
    if (vars[key] !== undefined && vars[key] !== '') return vars[key];
    varsSin.push('{' + key + '}'); return m;
  });
  return { texto: texto, varsSin: varsSin };
}

// ─────────────────────────────────────────────────────────────

function api_parseDate_(val) {
  if (!val) return null;
  var d = val instanceof Date ? val : new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function fmtFechaAPI_(val) {
  if (!val) return '';
  try {
    const d = val instanceof Date ? val : new Date(val);
    if (isNaN(d)) return val.toString();
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } catch (_) { return val ? val.toString() : ''; }
}

// ── Respuesta artista desde email (botones Sí / No) ──────────

function registrarRespuesta_(trqId, resp) {
  if (!trqId || (resp !== 'si' && resp !== 'no')) {
    return HtmlService.createHtmlOutput(paginaRespuesta_(false, '', 'Enlace no válido.')).setTitle('Kampaoh Trueque');
  }

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var bdd = ss.getSheetByName('BDD_Trueque');
  if (!bdd) return HtmlService.createHtmlOutput(paginaRespuesta_(false, '', 'Error interno.')).setTitle('Error');

  var headers = api_hdrs_(bdd);
  var idCol   = headers.indexOf('ID');
  var frmCol  = headers.indexOf('Formulario');  // Q
  var estCol  = headers.indexOf('Estado');       // R

  var allRows = bdd.getRange(2, 1, Math.max(bdd.getLastRow()-1,1), headers.length).getValues();
  var rowIdx  = -1;
  for (var i = 0; i < allRows.length; i++) {
    if (allRows[i][idCol] && allRows[i][idCol].toString() === trqId) { rowIdx = i; break; }
  }
  if (rowIdx < 0) return HtmlService.createHtmlOutput(paginaRespuesta_(false, '', 'Registro no encontrado.')).setTitle('Error');

  var nombre    = (allRows[rowIdx][headers.indexOf('Nombre')] || '').toString();
  var sheetRow  = rowIdx + 2;
  var frmVal    = resp === 'si' ? 'Acepta'     : 'Rechaza';
  var estVal    = resp === 'si' ? 'En proceso' : 'No acepta';

  if (frmCol >= 0) bdd.getRange(sheetRow, frmCol + 1).setValue(frmVal);
  if (estCol >= 0) bdd.getRange(sheetRow, estCol + 1).setValue(estVal);

  try {
    var log = ss.getSheetByName('Log_Acciones');
    if (log) log.appendRow([new Date(), nombre, 'Respuesta PL-02 (email)', 'BDD_Trueque', '', '', trqId, 'PL-02', '', 'OK',
      'Formulario → ' + frmVal + ' · Estado → ' + estVal, '', '', '']);
  } catch(e) {}

  return HtmlService.createHtmlOutput(paginaRespuesta_(resp === 'si', nombre, '')).setTitle('Kampaoh Trueque');
}

// ── Colorear conflictos en Google Sheets ─────────────────────

function api_updateCalGuests_(trqId) {
  if (!trqId) throw new Error('Falta parámetro id');
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var bdd     = ss.getSheetByName('BDD_Trueque');
  if (!bdd) throw new Error('Hoja BDD_Trueque no encontrada');
  var headers = api_hdrs_(bdd);
  var idCol   = headers.indexOf('ID');
  var allRows = bdd.getRange(2, 1, Math.max(bdd.getLastRow()-1,1), headers.length).getValues();
  var rowData = null;
  for (var i = 0; i < allRows.length; i++) {
    if (allRows[i][idCol] && allRows[i][idCol].toString() === trqId) { rowData = allRows[i]; break; }
  }
  if (!rowData) throw new Error('Registro ' + trqId + ' no encontrado');

  var g = function(n){ var idx = headers.indexOf(n); return idx >= 0 ? (rowData[idx] || '') : ''; };
  var nombre      = g('Nombre').toString();
  var campingNom  = g('Camping').toString();
  var llegada     = g('Llegada');
  var artistEmail = g('Email').toString().trim();

  // Construir lista de invitados: artista + camping
  var campingData   = api_camping_(campingNom);
  var campingEmails = campingData
    ? (campingData['Email'] || '').toString().split(/[\n,;]+/).map(function(e){ return e.trim(); }).filter(function(e){ return e.indexOf('@') >= 0; })
    : [];
  var guests = [];
  if (artistEmail && artistEmail.indexOf('@') >= 0) guests.push(artistEmail);
  guests = guests.concat(campingEmails);
  if (guests.indexOf('noemirodriguez@kampaoh.com') < 0) guests.push('noemirodriguez@kampaoh.com');
  if (!guests.length) throw new Error('No hay emails de artista ni de camping para invitar');

  // Buscar evento en Calendar por fecha y título exacto
  var d1 = llegada instanceof Date ? llegada : new Date(llegada);
  if (isNaN(d1.getTime())) throw new Error('Fecha de llegada no válida');
  var dayStart = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate());
  var dayEnd   = new Date(d1.getFullYear(), d1.getMonth(), d1.getDate() + 2);
  var titleSearch = '⛺ Trueque: ' + nombre + ' · ' + campingNom;
  var cal    = CalendarApp.getDefaultCalendar();
  var events = cal.getEvents(dayStart, dayEnd);
  var found  = null;
  for (var j = 0; j < events.length; j++) {
    if (events[j].getTitle() === titleSearch) { found = events[j]; break; }
  }
  if (!found) throw new Error('Evento no encontrado. Título buscado: "' + titleSearch + '"');

  // Añadir invitados que no estén ya
  var added = [];
  guests.forEach(function(email) {
    try { found.addGuest(email); added.push(email); } catch(e) {}
  });

  return { ok: true, added: added, eventTitle: found.getTitle() };
}

function colorConflictos_() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var bdd = ss.getSheetByName('BDD_Trueque');
  if (!bdd || bdd.getLastRow() <= 1) return { ok: true, colored: 0 };

  var headers = api_hdrs_(bdd);
  var ci      = function(n) { return headers.indexOf(n); };
  var numCols = headers.length;
  var data    = bdd.getRange(2, 1, bdd.getLastRow() - 1, numCols).getValues();

  var cancelados = ['Rechaza', 'No acepta', 'Cancelado', 'Cancelado STOP'];
  var conflicts = data.map(function(row, i) {
    var id         = row[ci('ID')];
    var camping    = (row[ci('Camping')] || '').toString();
    var llegadaFmt = api_fmt_(row[ci('Llegada')]);
    var estadoRaw2 = (row[ci('Estado')] || '').toString().trim();
    var estado     = estadoRaw2 === 'Cancelada' ? 'Cancelado' : estadoRaw2 === 'Cancelada STOP' ? 'Cancelado STOP' : estadoRaw2;
    if (!id || !camping || !llegadaFmt) return false;
    if (cancelados.indexOf(estado) >= 0) return false;
    for (var j = 0; j < data.length; j++) {
      if (j === i) continue;
      if (!data[j][ci('ID')]) continue;
      var estJ = (data[j][ci('Estado')] || '').toString().trim();
      estJ = estJ === 'Cancelada' ? 'Cancelado' : estJ === 'Cancelada STOP' ? 'Cancelado STOP' : estJ;
      if (cancelados.indexOf(estJ) >= 0) continue;
      if ((data[j][ci('Camping')] || '').toString() !== camping) continue;
      if (api_fmt_(data[j][ci('Llegada')]) === llegadaFmt) return true;
    }
    return false;
  });

  var colored = 0;
  conflicts.forEach(function(isConflict, i) {
    if (!data[i][ci('ID')]) return;
    var range = bdd.getRange(i + 2, 1, 1, numCols);
    if (isConflict) { range.setBackground('#FFF9C4'); colored++; }
    else { range.setBackground(null); }
  });

  return { ok: true, colored: colored };
}

function paginaRespuesta_(acepta, nombre, errorMsg) {
  if (errorMsg) return '<body style="font-family:sans-serif;text-align:center;padding:60px"><h2>⚠️ ' + errorMsg + '</h2></body>';
  var n = nombre ? ', ' + nombre : '';
  return acepta
    ? '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
      '<body style="font-family:sans-serif;background:#FFFBEB;text-align:center;padding:60px 20px">' +
      '<div style="max-width:480px;margin:0 auto">' +
      '<div style="font-size:3rem;margin-bottom:16px">💛</div>' +
      '<h1 style="color:#111;margin-bottom:12px">¡Genial' + n + '!</h1>' +
      '<p style="color:#555;font-size:1.05rem;line-height:1.7">Hemos registrado tu confirmación. En breve te enviamos todos los detalles del trueque.</p>' +
      '<p style="color:#aaa;font-size:.8rem;margin-top:48px">Kampaoh Trueque 2026</p>' +
      '</div></body></html>'
    : '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
      '<body style="font-family:sans-serif;background:#F9FAFB;text-align:center;padding:60px 20px">' +
      '<div style="max-width:480px;margin:0 auto">' +
      '<div style="font-size:3rem;margin-bottom:16px">😔</div>' +
      '<h1 style="color:#111;margin-bottom:12px">Oh' + n + '... ¡Qué pena!</h1>' +
      '<p style="color:#555;font-size:1.05rem;line-height:1.7">Nos pone muy tristes que no puedas participar esta vez, pero lo entendemos perfectamente. Si en algún momento te apetece, ¡escríbenos y buscamos la forma de coincidir!</p>' +
      '<p style="color:#aaa;font-size:.8rem;margin-top:48px">Kampaoh Trueque 2026</p>' +
      '</div></body></html>';
}

// ── Valoración post-trueque (PL-09) ──────────────────────────

function registrarValoracion_(trqId, stars) {
  if (!trqId || !stars) {
    return HtmlService.createHtmlOutput(paginaValoracion_(0, '', 'Enlace no válido.')).setTitle('Kampaoh Trueque');
  }
  var starsN = parseInt(stars);
  if (isNaN(starsN) || starsN < 1 || starsN > 5) {
    return HtmlService.createHtmlOutput(paginaValoracion_(0, '', 'Valoración no válida.')).setTitle('Kampaoh Trueque');
  }

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var bdd = ss.getSheetByName('BDD_Trueque');
  if (!bdd) return HtmlService.createHtmlOutput(paginaValoracion_(0, '', 'Error interno.')).setTitle('Error');

  var headers = api_hdrs_(bdd);
  var idCol   = headers.indexOf('ID');
  var valCol  = headers.indexOf('Valoracion');

  var allRows = bdd.getRange(2, 1, Math.max(bdd.getLastRow()-1,1), headers.length).getValues();
  var rowIdx  = -1;
  for (var i = 0; i < allRows.length; i++) {
    if (allRows[i][idCol] && allRows[i][idCol].toString() === trqId) { rowIdx = i; break; }
  }
  if (rowIdx < 0) return HtmlService.createHtmlOutput(paginaValoracion_(0, '', 'Registro no encontrado.')).setTitle('Error');

  var nombre = (allRows[rowIdx][headers.indexOf('Nombre')] || '').toString();

  if (valCol >= 0) bdd.getRange(rowIdx + 2, valCol + 1).setValue(starsN);
  // PropertiesService como fallback si la columna Valoracion no existe en la hoja
  try { PropertiesService.getScriptProperties().setProperty('val_' + trqId, starsN.toString()); } catch(ep) {}

  try {
    var log = ss.getSheetByName('Log_Acciones');
    if (log) log.appendRow([new Date(), nombre, 'Valoración PL-09 (email)', 'BDD_Trueque', '', '', trqId, 'PL-09', '', 'OK',
      'Valoración: ' + starsN + ' estrellas', '', '', '']);
  } catch(e) {}

  return HtmlService.createHtmlOutput(paginaValoracion_(starsN, nombre, '')).setTitle('Kampaoh Trueque');
}

function paginaValoracion_(stars, nombre, errorMsg) {
  if (errorMsg) return '<!DOCTYPE html><html><body style="font-family:sans-serif;text-align:center;padding:60px"><h2>⚠️ ' + errorMsg + '</h2></body></html>';
  var n = nombre ? ', ' + nombre : '';
  var filled = '', empty = '';
  for (var i = 0; i < stars; i++) filled += '★';
  for (var j = stars; j < 5; j++) empty += '☆';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="font-family:sans-serif;background:#FFFBEB;text-align:center;padding:60px 20px">' +
    '<div style="max-width:480px;margin:0 auto">' +
    '<div style="font-size:3rem;margin-bottom:16px">💛</div>' +
    '<h1 style="color:#111;margin-bottom:12px">¡Gracias' + n + '!</h1>' +
    '<div style="font-size:2.5rem;color:#F59E0B;margin:20px 0">' + filled + '<span style="color:#D1D5DB">' + empty + '</span></div>' +
    '<p style="color:#555;font-size:1.05rem;line-height:1.7">Hemos registrado tu valoración de <strong>' + stars + ' estrella' + (stars !== 1 ? 's' : '') + '</strong>. ¡Tu opinión nos ayuda a mejorar!</p>' +
    '<p style="color:#aaa;font-size:.8rem;margin-top:48px">Kampaoh Trueque 2026</p>' +
    '</div></body></html>';
}
