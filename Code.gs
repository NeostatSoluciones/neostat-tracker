/**
 * NEOSTAT · CENTRO DE CONTROL
 * Backend Apps Script vinculado a Google Sheets.
 *
 * Después de pegar este código:
 *  1. Habilita Calendar API: barra izquierda → Services (+) → Google Calendar API → Agregar
 *  2. Ejecuta la función instalar() una vez (Run → instalar)
 *  3. Llena la hoja "Agentes" con los emails de tu equipo
 *  4. Deploy → Nueva implementación → Aplicación web →
 *     Ejecutar como: yo · Acceso: cualquiera con el enlace
 *  5. Copia la URL del Web App y pégala en index.html como API_URL
 */

const SHEETS = {
  AGENTES:     'Agentes',
  CAMPANAS:    'Campañas',
  ESTADO:      'EstadoActual',
  EVENTOS:     'EventosEstado',
  ACTIVIDADES: 'Actividades',
  MENSAJES:    'Mensajes',
};
const STATES     = ['online','comida','bano','break','offline'];
const ACTIVITIES = ['llamada','cita','demo','cierre'];

/* ═════════════════════ ENTRY POINTS ═════════════════════ */

function doGet(e) {
  const action = (e.parameter.action || '').trim();
  return out(safe(() => {
    switch (action) {
      case 'status':       return api_status();
      case 'messages':     return api_messages(parseInt(e.parameter.since || 0));
      case 'agents':       return api_agents();
      case 'campaigns':    return api_campaigns();
      default: return { ok:false, error:'unknown_action' };
    }
  }));
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  return out(safe(() => {
    switch (body.action) {
      case 'changeState':    return api_changeState(body.email, body.newState);
      case 'logActivity':    return api_logActivity(body.email, body.type, body.campaign);
      case 'sendMessage':    return api_sendMessage(body.from, body.to, body.text);
      case 'createMeet':     return api_createMeet(body.from, body.to, body.title);
      case 'heartbeat':      return api_heartbeat(body.email);
      case 'overrideState':  return api_overrideState(body.supervisor, body.targetEmail, body.newState);
      // Admin (solo supervisor)
      case 'addAgent':       return api_addAgent(body.supervisor, body.email, body.nombre, body.rol, body.telefono, body.notas);
      case 'updateAgent':    return api_updateAgent(body.supervisor, body.targetEmail, body.nombre, body.rol, body.activo, body.telefono, body.notas);
      case 'removeAgent':    return api_removeAgent(body.supervisor, body.targetEmail);
      case 'addCampaign':    return api_addCampaign(body.supervisor, body.nombre, body.descripcion, body.fechaInicio, body.fechaFin);
      case 'updateCampaign': return api_updateCampaign(body.supervisor, body.id, body.nombre, body.descripcion, body.fechaInicio, body.fechaFin, body.activa);
      case 'removeCampaign': return api_removeCampaign(body.supervisor, body.id);
      default: return { ok:false, error:'unknown_action' };
    }
  }));
}

function safe(fn) {
  try { return fn(); }
  catch (err) { return { ok:false, error: err.toString() }; }
}
function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ═════════════════════ HELPERS ═════════════════════ */

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh(name) { return ss().getSheetByName(name); }
function tz() { return Session.getScriptTimeZone(); }
function todayStr() { return Utilities.formatDate(new Date(), tz(), 'yyyy-MM-dd'); }
function hourStr(d)  { return Utilities.formatDate(d, tz(), 'HH:mm:ss'); }

function findAgent(email) {
  if (!email) return null;
  const data = sh(SHEETS.AGENTES).getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (((data[i][0] || '') + '').toLowerCase() === email.toLowerCase()) {
      return {
        row: i+1,
        email: data[i][0],
        nombre: data[i][1] || data[i][0],
        rol: ((data[i][2] || 'agente')+'').toLowerCase(),
        activo: isTrue(data[i][3]),
        telefono: data[i][4] || '',
        notas: data[i][5] || ''
      };
    }
  }
  return null;
}
function isTrue(v) { return v===true || v===1 || v==='SI' || v==='TRUE' || v==='Sí' || v==='Si' || v==='si'; }
function isAllowed(email) { const a = findAgent(email); return a && a.activo; }
function isSupervisor(email) { const a = findAgent(email); return a && a.rol === 'supervisor' && a.activo; }
function requireSupervisor(email) {
  if (!isSupervisor(email)) throw new Error('Solo el supervisor puede ejecutar esta acción');
}

/* ═════════════════════ STATUS ═════════════════════ */

function api_status() {
  const agentes = sh(SHEETS.AGENTES).getDataRange().getValues();
  const estados = sh(SHEETS.ESTADO).getDataRange().getValues();
  const eventos = sh(SHEETS.EVENTOS).getDataRange().getValues();
  const acts    = sh(SHEETS.ACTIVIDADES).getDataRange().getValues();
  const hoy = todayStr();

  const estadoMap = {};
  for (let i=1; i<estados.length; i++) {
    const r = estados[i]; if (!r[0]) continue;
    estadoMap[(r[0]+'').toLowerCase()] = {
      estado: r[2] || 'offline',
      desde: r[3] ? new Date(r[3]).getTime() : Date.now(),
      ultimaSenal: r[4] ? new Date(r[4]).getTime() : 0,
    };
  }

  const evByEmail = {}, actByEmail = {};
  for (let j=1; j<eventos.length; j++) {
    const ev = eventos[j]; if (!ev[0]) continue;
    const em = (ev[1]+'').toLowerCase();
    const fecha = Utilities.formatDate(new Date(ev[0]), tz(), 'yyyy-MM-dd');
    if (fecha !== hoy) continue;
    (evByEmail[em] = evByEmail[em] || []).push(ev);
  }
  for (let j=1; j<acts.length; j++) {
    const a = acts[j]; if (!a[0]) continue;
    const em = (a[1]+'').toLowerCase();
    const fecha = Utilities.formatDate(new Date(a[0]), tz(), 'yyyy-MM-dd');
    if (fecha !== hoy) continue;
    (actByEmail[em] = actByEmail[em] || []).push(a);
  }

  const result = [];
  for (let i=1; i<agentes.length; i++) {
    const r = agentes[i];
    if (!r[0] || !isTrue(r[3])) continue;
    const rol = ((r[2] || 'agente')+'').toLowerCase();
    if (rol === 'supervisor') continue;

    const email = (r[0]+'').toLowerCase();
    const e = estadoMap[email] || { estado:'offline', desde: Date.now(), ultimaSenal: 0 };

    const tiempos = {online:0, comida:0, bano:0, break:0, offline:0};
    const evs = evByEmail[email] || [];
    evs.sort((a,b) => new Date(a[0]).getTime() - new Date(b[0]).getTime());
    for (let k=0; k<evs.length; k++) {
      const cur = evs[k];
      const next = evs[k+1];
      const start = new Date(cur[0]).getTime();
      const end = next ? new Date(next[0]).getTime() : Date.now();
      const estadoEv = cur[3];
      if (tiempos[estadoEv] !== undefined) tiempos[estadoEv] += Math.max(0, end - start);
    }

    const actividades = {llamada:0, cita:0, demo:0, cierre:0};
    (actByEmail[email] || []).forEach(a => { if (actividades[a[3]] !== undefined) actividades[a[3]]++; });

    let entrada = null, salida = null;
    evs.forEach(ev => {
      if (ev[3] === 'online' && !entrada) entrada = new Date(ev[0]).getTime();
      if (ev[3] === 'offline') salida = new Date(ev[0]).getTime();
    });

    result.push({
      email: r[0], nombre: r[1] || r[0], rol,
      estado: e.estado, desde: e.desde, ultimaSenal: e.ultimaSenal,
      tiempos, actividades, entrada, salida
    });
  }
  return { ok:true, agents: result, ts: Date.now() };
}

/* ═════════════════════ ESTADO ═════════════════════ */

function api_heartbeat(email) {
  if (!isAllowed(email)) return { ok:false, error:'not_allowed' };
  const sheet = sh(SHEETS.ESTADO);
  const data = sheet.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (((data[i][0]+'') || '').toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i+1, 5).setValue(new Date());
      return { ok:true };
    }
  }
  const a = findAgent(email);
  sheet.appendRow([email, a.nombre, 'offline', new Date(), new Date()]);
  return { ok:true };
}

function api_changeState(email, newState) {
  if (!STATES.includes(newState)) return { ok:false, error:'bad_state' };
  if (!isAllowed(email)) return { ok:false, error:'not_allowed' };
  return _doChangeState(email, newState);
}

function api_overrideState(supervisor, targetEmail, newState) {
  requireSupervisor(supervisor);
  if (!STATES.includes(newState)) return { ok:false, error:'bad_state' };
  if (!isAllowed(targetEmail))    return { ok:false, error:'target_not_allowed' };
  return _doChangeState(targetEmail, newState, supervisor);
}

function _doChangeState(email, newState, byEmail) {
  const a = findAgent(email);
  const estadoSheet = sh(SHEETS.ESTADO);
  const data = estadoSheet.getDataRange().getValues();
  let oldState = 'offline', row = -1;
  for (let i=1; i<data.length; i++) {
    if (((data[i][0]+'')||'').toLowerCase() === email.toLowerCase()) {
      oldState = data[i][2] || 'offline'; row = i+1; break;
    }
  }
  const ts = new Date();
  sh(SHEETS.EVENTOS).appendRow([
    ts, email, a.nombre, newState, oldState,
    todayStr(), hourStr(ts),
    byEmail ? ('override:' + byEmail) : ''
  ]);
  if (row === -1) estadoSheet.appendRow([email, a.nombre, newState, ts, ts]);
  else {
    estadoSheet.getRange(row, 3).setValue(newState);
    estadoSheet.getRange(row, 4).setValue(ts);
    estadoSheet.getRange(row, 5).setValue(ts);
  }
  return { ok:true, oldState, newState, ts: ts.getTime() };
}

/* ═════════════════════ ACTIVIDAD (con campaña) ═════════════════════ */

function api_logActivity(email, type, campaign) {
  if (!ACTIVITIES.includes(type)) return { ok:false, error:'bad_type' };
  if (!isAllowed(email)) return { ok:false, error:'not_allowed' };
  const a = findAgent(email);
  const ts = new Date();
  sh(SHEETS.ACTIVIDADES).appendRow([
    ts, email, a.nombre, type, todayStr(), hourStr(ts),
    campaign || ''
  ]);
  return { ok:true, ts: ts.getTime() };
}

/* ═════════════════════ MENSAJES ═════════════════════ */

function api_messages(since) {
  const data = sh(SHEETS.MENSAJES).getDataRange().getValues();
  const msgs = [];
  const start = Math.max(1, data.length - 200);
  for (let i=start; i<data.length; i++) {
    const r = data[i]; if (!r[0]) continue;
    const ts = new Date(r[0]).getTime();
    if (ts <= since) continue;
    msgs.push({ ts, from:r[1], to:r[2], text:r[3], tipo:r[4], meetLink:r[5] || null });
  }
  return { ok:true, messages: msgs };
}

function api_sendMessage(from, to, text) {
  if (!isAllowed(from)) return { ok:false, error:'not_allowed' };
  text = (text || '').toString().slice(0, 1000);
  if (!text.trim()) return { ok:false, error:'empty' };
  const ts = new Date();
  const tipo = to === 'all' ? 'grupal' : 'directo';
  sh(SHEETS.MENSAJES).appendRow([ts, from, to, text, tipo, '']);
  return { ok:true, ts: ts.getTime() };
}

/* ═════════════════════ MEET ═════════════════════ */

function api_createMeet(from, to, title) {
  if (!isAllowed(from)) return { ok:false, error:'not_allowed' };
  const start = new Date();
  const end = new Date(start.getTime() + 60*60*1000);
  const eventTitle = title || ('Neostat · ' + Utilities.formatDate(start, tz(), 'HH:mm'));

  let meetLink = '';
  try {
    const created = Calendar.Events.insert({
      summary: eventTitle,
      start: { dateTime: start.toISOString() },
      end:   { dateTime: end.toISOString() },
      conferenceData: {
        createRequest: {
          requestId: Utilities.getUuid(),
          conferenceSolutionKey: { type: 'hangoutsMeet' }
        }
      },
      attendees: to !== 'all' ? [{ email: to }] : []
    }, 'primary', { conferenceDataVersion: 1, sendUpdates: 'none' });

    if (created.conferenceData && created.conferenceData.entryPoints) {
      const video = created.conferenceData.entryPoints.find(e => e.entryPointType === 'video');
      if (video) meetLink = video.uri;
    }
  } catch (err) {
    return { ok:false, error: 'Calendar API no habilitado. Servicios → + → Google Calendar API.' };
  }
  if (!meetLink) return { ok:false, error:'no_meet_link' };

  const ts = new Date();
  const tipo = to === 'all' ? 'grupal' : 'directo';
  sh(SHEETS.MENSAJES).appendRow([ts, from, to, '📹 Videollamada iniciada', tipo, meetLink]);
  return { ok:true, meetLink, ts: ts.getTime() };
}

/* ═════════════════════ AGENTES (CRUD) ═════════════════════ */

function api_agents() {
  const data = sh(SHEETS.AGENTES).getDataRange().getValues();
  const list = [];
  for (let i=1; i<data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      email: data[i][0],
      nombre: data[i][1] || data[i][0],
      rol: ((data[i][2]||'agente')+'').toLowerCase(),
      activo: isTrue(data[i][3]),
      telefono: data[i][4] || '',
      notas: data[i][5] || ''
    });
  }
  return { ok:true, agents: list };
}

function api_addAgent(supervisor, email, nombre, rol, telefono, notas) {
  requireSupervisor(supervisor);
  if (!email || !nombre) return { ok:false, error:'datos_incompletos' };
  // Verifica que no exista ya
  if (findAgent(email)) return { ok:false, error:'agente_ya_existe' };
  const r = (rol || 'agente').toLowerCase();
  sh(SHEETS.AGENTES).appendRow([email, nombre, r, 'SI', telefono || '', notas || '']);
  return { ok:true };
}

function api_updateAgent(supervisor, targetEmail, nombre, rol, activo, telefono, notas) {
  requireSupervisor(supervisor);
  const sheet = sh(SHEETS.AGENTES);
  const data = sheet.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (((data[i][0]+'')||'').toLowerCase() === (targetEmail+'').toLowerCase()) {
      if (nombre !== undefined && nombre !== null) sheet.getRange(i+1, 2).setValue(nombre);
      if (rol !== undefined && rol !== null) sheet.getRange(i+1, 3).setValue(rol);
      if (activo !== undefined && activo !== null) sheet.getRange(i+1, 4).setValue(isTrue(activo) ? 'SI' : 'NO');
      if (telefono !== undefined && telefono !== null) sheet.getRange(i+1, 5).setValue(telefono);
      if (notas !== undefined && notas !== null) sheet.getRange(i+1, 6).setValue(notas);
      return { ok:true };
    }
  }
  return { ok:false, error:'agente_no_encontrado' };
}

function api_removeAgent(supervisor, targetEmail) {
  // No borra histórico, solo lo desactiva
  return api_updateAgent(supervisor, targetEmail, null, null, false, null, null);
}

/* ═════════════════════ CAMPAÑAS (CRUD) ═════════════════════ */

function api_campaigns() {
  const data = sh(SHEETS.CAMPANAS).getDataRange().getValues();
  const list = [];
  for (let i=1; i<data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      id: data[i][0],
      nombre: data[i][1],
      descripcion: data[i][2] || '',
      fechaInicio: data[i][3] ? new Date(data[i][3]).getTime() : null,
      fechaFin: data[i][4] ? new Date(data[i][4]).getTime() : null,
      activa: isTrue(data[i][5]),
      creadoPor: data[i][6] || '',
      creadoEn: data[i][7] ? new Date(data[i][7]).getTime() : null
    });
  }
  return { ok:true, campaigns: list };
}

function api_addCampaign(supervisor, nombre, descripcion, fechaInicio, fechaFin) {
  requireSupervisor(supervisor);
  if (!nombre) return { ok:false, error:'nombre_requerido' };
  const id = 'C' + (new Date().getTime().toString(36).toUpperCase());
  sh(SHEETS.CAMPANAS).appendRow([
    id, nombre, descripcion || '',
    fechaInicio ? new Date(fechaInicio) : '',
    fechaFin ? new Date(fechaFin) : '',
    'SI', supervisor, new Date()
  ]);
  return { ok:true, id };
}

function api_updateCampaign(supervisor, id, nombre, descripcion, fechaInicio, fechaFin, activa) {
  requireSupervisor(supervisor);
  const sheet = sh(SHEETS.CAMPANAS);
  const data = sheet.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (((data[i][0]+'')||'') === id) {
      if (nombre !== undefined && nombre !== null) sheet.getRange(i+1, 2).setValue(nombre);
      if (descripcion !== undefined && descripcion !== null) sheet.getRange(i+1, 3).setValue(descripcion);
      if (fechaInicio !== undefined && fechaInicio !== null) sheet.getRange(i+1, 4).setValue(fechaInicio ? new Date(fechaInicio) : '');
      if (fechaFin !== undefined && fechaFin !== null) sheet.getRange(i+1, 5).setValue(fechaFin ? new Date(fechaFin) : '');
      if (activa !== undefined && activa !== null) sheet.getRange(i+1, 6).setValue(isTrue(activa) ? 'SI' : 'NO');
      return { ok:true };
    }
  }
  return { ok:false, error:'campana_no_encontrada' };
}

function api_removeCampaign(supervisor, id) {
  // No borra: la desactiva (preserva histórico en Actividades)
  return api_updateCampaign(supervisor, id, null, null, null, null, false);
}

/* ═════════════════════ INSTALACIÓN ═════════════════════ */

function instalar() {
  const required = [
    { name: SHEETS.AGENTES,     headers: ['Email','Nombre','Rol','Activo','Teléfono','Notas'] },
    { name: SHEETS.CAMPANAS,    headers: ['ID','Nombre','Descripción','FechaInicio','FechaFin','Activa','CreadoPor','CreadoEn'] },
    { name: SHEETS.ESTADO,      headers: ['Email','Nombre','Estado','Desde','UltimaSenal'] },
    { name: SHEETS.EVENTOS,     headers: ['Timestamp','Email','Nombre','EstadoNuevo','EstadoAnterior','Fecha','Hora','Origen'] },
    { name: SHEETS.ACTIVIDADES, headers: ['Timestamp','Email','Nombre','Tipo','Fecha','Hora','Campaña'] },
    { name: SHEETS.MENSAJES,    headers: ['Timestamp','De','Para','Texto','Tipo','MeetLink'] },
  ];
  required.forEach(r => {
    let s = ss().getSheetByName(r.name);
    if (!s) s = ss().insertSheet(r.name);
    s.getRange(1, 1, 1, r.headers.length)
      .setValues([r.headers])
      .setFontWeight('bold')
      .setBackground('#0a0a0a')
      .setFontColor('#3b82f6');
    s.setFrozenRows(1);
    s.autoResizeColumns(1, r.headers.length);
  });

  // Pre-poblar agentes Neostat
  const agentesSheet = sh(SHEETS.AGENTES);
  if (agentesSheet.getLastRow() === 1) {
    agentesSheet.appendRow(['neostatsoluciones@gmail.com', 'Christian López',     'supervisor', 'SI', '', 'Cuenta maestra']);
    agentesSheet.appendRow(['christian.lopez@neo-stat.com', 'Christian López',    'supervisor', 'SI', '', 'Cuenta corporativa']);
    agentesSheet.appendRow(['alandgrave1000@gmail.com',    'Alfonso Landgrave',   'agente',     'SI', '', '']);
    agentesSheet.appendRow(['robertodaniel259@gmail.com',  'Roberto Santillán',   'agente',     'SI', '', '']);
    agentesSheet.appendRow(['jose8566619@gmail.com',       'Manuel Ramírez',      'agente',     'SI', '', '']);
    agentesSheet.appendRow(['ximenezmitzi2@gmail.com',     'Mitzi Jiménez',       'agente',     'SI', '', '']);
    agentesSheet.appendRow(['ventas@neo-stat.com',         'Teresa',              'agente',     'SI', '', 'Cuenta compartida ventas']);
    agentesSheet.appendRow(['',                            'Mario Martínez',      'agente',     'NO', '', 'Pendiente confirmar email']);
  }

  SpreadsheetApp.getUi().alert('✅ Listo. Hojas creadas y equipo Neostat pre-cargado. Revisa la hoja "Agentes" y completa el email de Mario cuando lo tengas.');
}

/* ═════════════════════ TRIGGERS OPCIONALES ═════════════════════ */

// Auto-offline tras 5 min sin heartbeat. Configura un trigger time-driven cada 5 min.
function autoOfflineIfStale() {
  const STALE_MS = 5 * 60 * 1000;
  const sheet = sh(SHEETS.ESTADO);
  const data = sheet.getDataRange().getValues();
  const now = Date.now();
  for (let i=1; i<data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if ((r[2] || 'offline') === 'offline') continue;
    const last = r[4] ? new Date(r[4]).getTime() : 0;
    if (now - last > STALE_MS) {
      _doChangeState(r[0], 'offline', 'system:stale');
    }
  }
}
