/**
 * ATS Saint Luke's — Fase 2 y 3
 * Entrevista por voz con IA (Retell AI) · Rúbricas adicionales · Depuración de datos
 * Creado por RHINOBUILD
 *
 * Este archivo se agrega junto a Código.gs (comparten el mismo espacio de nombres).
 *
 * Propiedades del script nuevas:
 *   RETELL_API_KEY      (obligatoria para entrevistas por voz)
 *   RETELL_AGENT_ID     (se llena sola con crearAgenteRetell)
 *   RETELL_VOICE_ID     (opcional: voz del asistente; si falta, se elige una voz en español)
 *   RETELL_FROM_NUMBER  (opcional: número importado a Retell, formato +52…, para llamadas telefónicas)
 *   URL_ENTREVISTA      (página pública de la entrevista; por defecto la de GitHub Pages)
 *   ENTREVISTA_HORAS    (vigencia de la invitación; por defecto 72)
 *   MESES_BANCO         (meses que se conserva el banco de talento; por defecto 24)
 */

// IMPORTANTE: en el editor, Fase2.gs debe ir DEBAJO de Código.gs (Apps Script carga los archivos en ese orden).
if (typeof HOJAS === 'undefined') throw new Error('Fase2.gs se cargó antes que Código.gs: muévelo abajo en la lista de archivos.');
HOJAS.Entrevistas = ['ID', 'Fecha', 'PostulacionID', 'Token', 'Estado', 'Modo', 'CallID', 'Intentos',
  'Inicio', 'DuracionSeg', 'Transcripcion', 'GrabacionURL', 'EvaluacionIA', 'PuntajeEntrevista',
  'Expira', 'CreadoPor', 'Actualizado'];

const RETELL = 'https://api.retellai.com';
const ENT = {
  ESTADOS: ['Invitado', 'EnCurso', 'Evaluada', 'NoRealizada', 'SinConsentimiento', 'Expirada', 'Error'],
  MAX_INTENTOS: 3,
  DURACION_MIN_SEG: 45,
  URL_DEFECTO: 'https://rhinobuild.github.io/BPH360/entrevista.html'
};

/* ============================================================
 *  INSTALACIÓN DE FASE 2
 * ============================================================ */

/** Ejecutar una vez: crea la hoja Entrevistas, rúbricas adicionales y el disparador de entrevistas. */
function setupFase2() {
  const ss = ss_();
  Object.keys(HOJAS).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    const h = HOJAS[n];
    sh.getRange(1, 1, 1, h.length).setValues([h])
      .setFontWeight('bold').setBackground('#0B2545').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  });
  sembrarRubricasAdicionales();
  if (!prop_('URL_ENTREVISTA')) setProp_('URL_ENTREVISTA', ENT.URL_DEFECTO);

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'procesarEntrevistas')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('procesarEntrevistas').timeBased().everyMinutes(5).create();
  Logger.log('Fase 2 lista. Siguiente paso: agrega RETELL_API_KEY y ejecuta crearAgenteRetell().');
}

/** Crea (o recrea) el asistente de voz en Retell y guarda su ID. Requiere RETELL_API_KEY. */
function crearAgenteRetell() {
  const llm = retell_('post', '/create-retell-llm', {
    start_speaker: 'agent',
    general_prompt: PROMPT_ENTREVISTA_,
    begin_message: 'Hola, {{nombre}}. Soy Sofía, asistente virtual de inteligencia artificial de Capital Humano de Saint Luke\'s Hospitals. ' +
      'Te contacto por tu postulación a {{puesto}}. Antes de empezar: esta entrevista dura unos diez minutos, se graba y se transcribe ' +
      'para que el equipo de Capital Humano la revise; ninguna decisión la toma la inteligencia artificial. Si prefieres, puedes pedir ' +
      'una entrevista con una persona sin afectar tu proceso. ¿Estás de acuerdo en continuar?',
    general_tools: [{
      type: 'end_call', name: 'end_call',
      description: 'Termina la llamada cuando la entrevista concluyó, cuando la persona no autoriza la grabación o pide hablar con una persona, o cuando pide terminar.'
    }],
    default_dynamic_variables: { nombre: 'candidato', puesto: 'la vacante', sede: '', preguntas: '' }
  });

  let voz = prop_('RETELL_VOICE_ID');
  if (!voz) {
    const voces = retell_('get', '/list-voices');
    const es = voces.filter(v => /spanish|mexic|latin|español|espanol/i.test((v.accent || '') + ' ' + (v.voice_name || '')));
    const elegida = es.find(v => v.gender === 'female') || es[0] || voces.find(v => v.gender === 'female') || voces[0];
    voz = elegida.voice_id;
    Logger.log('Voces en español disponibles: ' + es.map(v => v.voice_id + ' (' + v.voice_name + ', ' + v.gender + ')').join(' · '));
  }

  const body = {
    agent_name: "Entrevistadora ATS Saint Luke's",
    response_engine: { type: 'retell-llm', llm_id: llm.llm_id },
    voice_id: voz,
    language: 'es-419',
    max_call_duration_ms: 20 * 60 * 1000,
    end_call_after_silence_ms: 30000,
    data_storage_setting: 'everything'
  };
  const anterior = prop_('RETELL_AGENT_ID');
  const agente = retell_('post', '/create-agent', body);
  setProp_('RETELL_AGENT_ID', agente.agent_id);
  setProp_('RETELL_LLM_ID', llm.llm_id);
  setProp_('RETELL_VOICE_ID', voz);
  Logger.log('Asistente creado: ' + agente.agent_id + ' · voz ' + voz + (anterior ? ' (reemplaza a ' + anterior + ')' : ''));
}

const PROMPT_ENTREVISTA_ = [
  'Eres Sofía, asistente virtual de inteligencia artificial de Capital Humano de Saint Luke\'s Hospitals, un grupo hospitalario privado en México.',
  'Realizas una entrevista inicial por voz con {{nombre}}, quien se postuló a la vacante de {{puesto}} en {{sede}}.',
  '',
  '## Consentimiento (obligatorio)',
  '- Tu primer mensaje ya explicó que eres una IA y que la entrevista se graba. Espera la respuesta.',
  '- Si la persona NO acepta, o pide hablar con una persona: dile con amabilidad que Capital Humano la contactará directamente, que su proceso no se ve afectado, despídete y usa end_call.',
  '- Si acepta, confirma brevemente que hablas con {{nombre}} y continúa.',
  '',
  '## Entrevista',
  'Haz estas preguntas, una por una, en este orden:',
  '{{preguntas}}',
  '',
  'Después de las preguntas anteriores, si queda tiempo, pregunta también:',
  '- ¿Qué te interesa de trabajar en Saint Luke\'s?',
  '- ¿Cuál es tu disponibilidad de horario y a partir de cuándo podrías incorporarte?',
  '',
  '## Estilo',
  '- Español de México, cálido, profesional y breve. Una sola pregunta por turno.',
  '- Si una respuesta es vaga, pide UN ejemplo concreto ("¿Me podrías dar un ejemplo de una situación real?"). No insistas más de una vez.',
  '- No evalúes ni califiques en voz alta; no digas si una respuesta es buena o mala. Solo agradece y avanza.',
  '- Si la persona no entiende, reformula con palabras más sencillas.',
  '- Duración objetivo: 8 a 12 minutos.',
  '',
  '## Prohibido',
  '- No preguntes ni comentes edad, estado civil, hijos o planes de tenerlos, embarazo, salud, discapacidad, religión, origen, orientación sexual, afiliación política o sindical, ni apariencia.',
  '- Si la persona menciona por iniciativa propia alguno de esos temas, no profundices: agradece y regresa a la entrevista.',
  '- No prometas contratación, sueldo ni prestaciones. Si pregunta por sueldo o prestaciones, di que Capital Humano lo compartirá en la siguiente etapa.',
  '- No inventes información sobre el hospital. Si no sabes algo, di que Capital Humano dará esa información.',
  '',
  '## Cierre',
  'Agradece su tiempo, explica que Capital Humano revisará la entrevista y le dará seguimiento por correo o teléfono, despídete y usa end_call.'
].join('\n');

/* ============================================================
 *  PANEL: INVITACIONES Y ESTADO DE ENTREVISTAS
 * ============================================================ */

/** Crea la invitación a la entrevista por voz y (opcional) la envía por correo. */
function invitarEntrevista(tok, postId, enviarCorreo) {
  const u = usuario_(tok);
  configVozOk_();
  const p = buscar_('Postulaciones', postId);
  const c = buscar_('Candidatos', p.CandidatoID);
  const v = buscar_('Vacantes', p.VacanteID);

  // Una sola invitación vigente por postulación.
  leer_('Entrevistas').filter(e => e.PostulacionID === postId && (e.Estado === 'Invitado' || e.Estado === 'NoRealizada' || (e.Estado === 'Error' && !String(e.Transcripcion || '').trim())))
    .forEach(e => actualizar_('Entrevistas', e._row, { Estado: 'Expirada', Actualizado: new Date() }));

  const horas = Number(prop_('ENTREVISTA_HORAS', '72')) || 72;
  const expira = new Date(Date.now() + horas * 3600 * 1000);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  const ent = {
    ID: id_('I'), Fecha: new Date(), PostulacionID: postId, Token: token, Estado: 'Invitado', Modo: 'Web',
    CallID: '', Intentos: 0, Expira: expira, CreadoPor: u, Actualizado: new Date()
  };
  agregar_('Entrevistas', ent);

  const link = prop_('URL_ENTREVISTA', ENT.URL_DEFECTO) + '?t=' + token;
  const limite = Utilities.formatDate(expira, APP.TZ, "dd/MM/yyyy 'a las' HH:mm");
  if (enviarCorreo) {
    correo_(c.Email, 'Entrevista en línea — ' + v.Puesto,
      '<p>Hola ' + esc_(c.Nombre) + ':</p>' +
      '<p>Tu perfil avanzó en el proceso para <b>' + esc_(v.Puesto) + '</b>. El siguiente paso es una entrevista inicial por voz, de unos 10 minutos, ' +
      'con nuestra asistente virtual. Puedes hacerla cuando quieras desde tu celular o computadora, antes del <b>' + limite + '</b>.</p>' +
      '<p style="text-align:center;margin:22px 0"><a href="' + link + '" style="background:#C9A227;color:#0B2545;padding:12px 22px;border-radius:8px;text-decoration:none;font-weight:bold">Iniciar mi entrevista</a></p>' +
      '<p>Te recomendamos un lugar tranquilo y audífonos. La entrevista se graba para que la revise el equipo de Capital Humano; ' +
      'ninguna decisión la toma la inteligencia artificial. Si prefieres una entrevista con una persona, responde a este correo.</p>');
  }
  bitacora_(u, postId, 'Invitación a entrevista IA', (enviarCorreo ? 'Correo enviado · ' : '') + 'vence ' + limite);

  const nombre = String(c.Nombre || '').split(' ')[0];
  return {
    ok: true, link, limite,
    whatsapp: 'Hola ' + nombre + ', te saluda Capital Humano de Saint Luke\'s Hospitals. Tu perfil para ' + v.Puesto +
      ' avanzó. Te invitamos a una entrevista en línea de 10 minutos con nuestra asistente virtual, antes del ' + limite + ': ' + link
  };
}

/** Llamada telefónica saliente (solo si hay un número de origen importado a Retell). */
function llamarCandidato(tok, postId) {
  const u = usuario_(tok);
  configVozOk_();
  const from = prop_('RETELL_FROM_NUMBER');
  if (!from) throw new Error('No hay número de origen (RETELL_FROM_NUMBER). Usa la invitación por liga.');
  const h = Number(Utilities.formatDate(new Date(), APP.TZ, 'H'));
  if (h < 9 || h >= 19) throw new Error('Las llamadas solo se hacen de 9:00 a 19:00 (hora del Pacífico).');

  const p = buscar_('Postulaciones', postId);
  const c = buscar_('Candidatos', p.CandidatoID);
  const to = e164_(c.Telefono);
  if (!to) throw new Error('El teléfono del candidato no es válido.');

  const ent = {
    ID: id_('I'), Fecha: new Date(), PostulacionID: postId, Token: '', Estado: 'EnCurso', Modo: 'Teléfono',
    CallID: '', Intentos: 1, Expira: '', CreadoPor: u, Actualizado: new Date()
  };
  const r = retell_('post', '/v2/create-phone-call', {
    from_number: from, to_number: to, override_agent_id: prop_('RETELL_AGENT_ID'),
    metadata: { postulacion: postId, entrevista: ent.ID },
    retell_llm_dynamic_variables: variablesEntrevista_(p)
  });
  ent.CallID = r.call_id;
  agregar_('Entrevistas', ent);
  bitacora_(u, postId, 'Llamada IA iniciada', to);
  return { ok: true };
}

/** Datos de entrevistas para el panel. */
function panelEntrevistas(tok) {
  usuario_(tok);
  const porPost = {};
  leer_('Entrevistas').forEach(e => {
    let ev = null;
    try { ev = e.EvaluacionIA ? JSON.parse(e.EvaluacionIA) : null; } catch (x) { ev = null; }
    (porPost[e.PostulacionID] = porPost[e.PostulacionID] || []).push({
      ID: e.ID, Fecha: e.Fecha, Estado: e.Estado, Modo: e.Modo, Intentos: e.Intentos, Inicio: e.Inicio,
      DuracionSeg: e.DuracionSeg, Transcripcion: String(e.Transcripcion || ''), GrabacionURL: e.GrabacionURL,
      Evaluacion: ev, Puntaje: e.PuntajeEntrevista, Expira: e.Expira, CreadoPor: e.CreadoPor
    });
  });
  // google.script.run no acepta objetos Date en la respuesta (llega null): se serializa a ISO.
  return JSON.parse(JSON.stringify({
    porPost,
    config: {
      vozActiva: !!(prop_('RETELL_API_KEY') && prop_('RETELL_AGENT_ID')),
      telefono: !!prop_('RETELL_FROM_NUMBER')
    }
  }));
}

function reevaluarEntrevista(tok, entId) {
  const u = usuario_(tok);
  const e = buscar_('Entrevistas', entId);
  if (!String(e.Transcripcion || '').trim()) throw new Error('Esta entrevista no tiene transcripción.');
  evaluarEntrevista_(e);
  bitacora_(u, e.PostulacionID, 'Entrevista IA reevaluada', entId);
  return { ok: true };
}

/* ============================================================
 *  PÁGINA PÚBLICA DE ENTREVISTA (GitHub Pages → doPost)
 * ============================================================ */

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { req = {}; }
  let out;
  try {
    if (req.action === 'info') out = entrevistaInfo_(req.t);
    else if (req.action === 'webcall') out = entrevistaWebCall_(req.t);
    else out = { ok: false, error: 'Acción no válida.' };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function entrevistaPorToken_(t) {
  if (!t || !/^[a-f0-9]{40}$/.test(String(t))) throw new Error('El enlace no es válido.');
  const e = leer_('Entrevistas').find(x => x.Token === t);
  if (!e) throw new Error('El enlace no es válido.');
  return e;
}

function entrevistaInfo_(t) {
  const e = liberarSiNoConecto_(entrevistaPorToken_(t));
  const p = buscar_('Postulaciones', e.PostulacionID);
  const c = buscar_('Candidatos', p.CandidatoID);
  const v = buscar_('Vacantes', p.VacanteID);
  return {
    ok: true, estado: disponible_(e), nombre: String(c.Nombre || '').split(' ')[0],
    puesto: v.Puesto, sede: v.Sede,
    expira: e.Expira ? Utilities.formatDate(new Date(e.Expira), APP.TZ, "dd/MM/yyyy 'a las' HH:mm") : '',
    aviso: prop_('AVISO_PRIVACIDAD_URL')
  };
}

/** Si el intento anterior nunca conectó (micrófono negado, se cerró la página…), permite reintentar. */
function liberarSiNoConecto_(e) {
  if (e.Estado !== 'EnCurso' || !e.CallID || e.Modo !== 'Web') return e;
  try {
    const call = retell_('get', '/v2/get-call/' + encodeURIComponent(e.CallID));
    if (call.call_status === 'ongoing') return e;
    const corta = (call.duration_ms || 0) < ENT.DURACION_MIN_SEG * 1000;
    if (call.call_status !== 'ended' || corta) {
      actualizar_('Entrevistas', e._row, { Estado: 'NoRealizada', Actualizado: new Date() });
      e.Estado = 'NoRealizada';
    }
  } catch (x) { /* si Retell no responde, se queda como está */ }
  return e;
}

/** 'disponible' | 'realizada' | 'expirada' | 'agotada' */
function disponible_(e) {
  if (e.Estado === 'Evaluada' || e.Estado === 'EnCurso') return 'realizada';
  if (e.Estado === 'SinConsentimiento') return 'realizada';
  // Un error al EVALUAR (ya hay transcripción) no obliga a la persona a repetir la entrevista.
  if (e.Estado === 'Error' && String(e.Transcripcion || '').trim()) return 'realizada';
  if (e.Estado === 'Expirada' || (e.Expira && new Date(e.Expira) < new Date())) return 'expirada';
  if (Number(e.Intentos || 0) >= ENT.MAX_INTENTOS) return 'agotada';
  if (e.Estado === 'Invitado' || e.Estado === 'NoRealizada' || e.Estado === 'Error') return 'disponible';
  return 'expirada';
}

/** El navegador del candidato pide aquí la conexión; la API key nunca sale del servidor. */
function entrevistaWebCall_(t) {
  configVozOk_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let e = entrevistaPorToken_(t);
    e = liberarSiNoConecto_(e);
    const estado = disponible_(e);
    if (estado !== 'disponible') {
      return { ok: false, error: estado === 'realizada' ? 'Esta entrevista ya se realizó. ¡Gracias!' :
        estado === 'agotada' ? 'Se agotaron los intentos de esta invitación. Escríbenos para darte una nueva.' :
        'Esta invitación ya venció. Escríbenos para darte una nueva.' };
    }
    const p = buscar_('Postulaciones', e.PostulacionID);
    const res = UrlFetchApp.fetch(RETELL + '/v3/create-web-call', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + prop_('RETELL_API_KEY') },
      payload: JSON.stringify({
        agent_id: prop_('RETELL_AGENT_ID'),
        metadata: { postulacion: e.PostulacionID, entrevista: e.ID },
        retell_llm_dynamic_variables: variablesEntrevista_(p)
      })
    });
    const code = res.getResponseCode(), body = res.getContentText();
    if (code >= 200 && code < 300) {
      const j = JSON.parse(body);
      actualizar_('Entrevistas', e._row, {
        Estado: 'EnCurso', CallID: j.call_id, Intentos: Number(e.Intentos || 0) + 1, Inicio: new Date(), Actualizado: new Date()
      });
      bitacora_('Candidato', e.PostulacionID, 'Entrevista IA iniciada', e.ID);
    }
    return { ok: code >= 200 && code < 300, status: code, body };
  } finally {
    lock.releaseLock();
  }
}

function variablesEntrevista_(p) {
  const c = buscar_('Candidatos', p.CandidatoID);
  const v = buscar_('Vacantes', p.VacanteID);
  let preguntas = [];
  try { preguntas = (JSON.parse(p.FichaIA || '{}').preguntas || []).slice(0, 5); } catch (x) { preguntas = []; }
  if (!preguntas.length) {
    preguntas = [
      'Cuéntame brevemente tu experiencia más reciente relacionada con este puesto.',
      'Describe una situación difícil con un paciente o cliente y cómo la resolviste.',
      '¿Qué haces cuando tienes varias tareas urgentes al mismo tiempo?'
    ];
  }
  return {
    nombre: String(c.Nombre || '').split(' ')[0] || 'candidato',
    puesto: String(v.Puesto || ''), sede: String(v.Sede || ''),
    preguntas: preguntas.map((q, i) => (i + 1) + '. ' + q).join('\n')
  };
}

/* ============================================================
 *  PROCESAMIENTO (disparador cada 5 minutos)
 * ============================================================ */

function procesarEntrevistas() {
  if (!prop_('RETELL_API_KEY')) return;
  const t0 = Date.now();
  leer_('Entrevistas').forEach(e => {
    if (Date.now() - t0 > 4.5 * 60 * 1000) return;
    try {
      if (e.Estado === 'Invitado' && e.Expira && new Date(e.Expira) < new Date()) {
        actualizar_('Entrevistas', e._row, { Estado: 'Expirada', Actualizado: new Date() });
      } else if (e.Estado === 'EnCurso' && e.CallID) {
        revisarLlamada_(e);
      } else if (e.Estado === 'Error' && String(e.Transcripcion || '').trim() && reintentosEval_(e) < 3) {
        // La llamada sí se hizo pero falló la evaluación (p. ej. la API de IA no respondió): se reintenta.
        evaluarEntrevista_(e);
      }
    } catch (err) {
      actualizar_('Entrevistas', e._row, {
        Estado: 'Error',
        EvaluacionIA: JSON.stringify({ error: String(err.message || err), reintentos: reintentosEval_(e) + 1 }),
        Actualizado: new Date()
      });
    }
  });
}

function revisarLlamada_(e) {
  const call = retell_('get', '/v2/get-call/' + encodeURIComponent(e.CallID));
  if (call.call_status === 'registered' || call.call_status === 'ongoing') {
    // Una sesión web que nunca conectó se libera después de 30 minutos.
    if (e.Inicio && Date.now() - new Date(e.Inicio).getTime() > 30 * 60 * 1000 && call.call_status === 'registered') {
      actualizar_('Entrevistas', e._row, { Estado: 'NoRealizada', Actualizado: new Date() });
    }
    return;
  }
  const dur = Math.round((call.duration_ms || 0) / 1000);
  const transcript = String(call.transcript || '');
  if (call.call_status !== 'ended' || dur < ENT.DURACION_MIN_SEG || transcript.length < 200) {
    actualizar_('Entrevistas', e._row, { Estado: 'NoRealizada', DuracionSeg: dur, Actualizado: new Date() });
    return;
  }
  let grabacion = '';
  try { grabacion = guardarGrabacion_(e, call.recording_url); } catch (x) { grabacion = ''; }
  actualizar_('Entrevistas', e._row, {
    DuracionSeg: dur, Transcripcion: transcript.slice(0, 45000), GrabacionURL: grabacion, Actualizado: new Date()
  });
  evaluarEntrevista_(buscar_('Entrevistas', e.ID));
}

function guardarGrabacion_(e, url) {
  if (!url) return '';
  const blob = UrlFetchApp.fetch(url).getBlob();
  const raiz = DriveApp.getFolderById(prop_('CV_FOLDER_ID'));
  const it = raiz.getFoldersByName('Entrevistas IA');
  const carpeta = it.hasNext() ? it.next() : raiz.createFolder('Entrevistas IA');
  const ext = (String(url).split('?')[0].match(/\.(wav|mp3|ogg|m4a)$/i) || [, 'wav'])[1];
  return carpeta.createFile(blob.setName(e.PostulacionID + ' — ' + e.ID + '.' + ext)).getUrl();
}

function evaluarEntrevista_(e) {
  const p = buscar_('Postulaciones', e.PostulacionID);
  const v = buscar_('Vacantes', p.VacanteID);
  const crit = leer_('Rubricas').filter(r => r.RubricaID === v.RubricaID);
  const prompt =
    'Eres analista de selección de un hospital privado en México. Abajo está la transcripción de una entrevista inicial por voz ' +
    'entre una asistente de IA ("Agent") y una persona candidata ("User") para el puesto de ' + v.Puesto + '.\n\n' +
    'REGLAS:\n' +
    '- Primero determina si la persona AUTORIZÓ continuar con la entrevista grabada.\n' +
    '- Evalúa solo con base en lo que la persona dijo. Si algo no se abordó, indica "Sin evidencia".\n' +
    '- No consideres ni menciones edad, sexo, estado civil, embarazo, salud, religión, origen, acento ni forma de hablar.\n' +
    '- No recomiendes contratar ni descartar: resume para que una persona decida.\n\n' +
    'CRITERIOS DEL PUESTO:\n' + crit.map(c => '- ' + c.Criterio + ': ' + c.Descripcion).join('\n') + '\n\n' +
    'COMPETENCIAS A VALORAR: comunicación, orientación al paciente/servicio, manejo de presión, trabajo en equipo, ' +
    'conocimiento técnico del puesto, disponibilidad.\n\n' +
    'Responde SOLO con JSON válido:\n' +
    '{"consentimiento":true,"resumen":"3-4 líneas",' +
    '"competencias":[{"competencia":"...","valoracion":"Sólida|Adecuada|Débil|Sin evidencia","evidencia":"cita o paráfrasis breve"}],' +
    '"respuestas":[{"pregunta":"...","sintesis":"..."}],"fortalezas":["..."],"a_profundizar":["temas para la entrevista presencial"],' +
    '"alertas":["contradicciones con el CV, respuestas evasivas, datos por verificar"],"disponibilidad":"lo que dijo sobre horario e incorporación"}\n\n' +
    'TRANSCRIPCIÓN:\n' + String(e.Transcripcion).slice(0, 40000);

  const r = extraerJSON_(llamarClaude_([{ type: 'text', text: prompt }], 3000));
  if (r.consentimiento === false) {
    actualizar_('Entrevistas', e._row, {
      Estado: 'SinConsentimiento', Transcripcion: '', GrabacionURL: '', EvaluacionIA: JSON.stringify({ consentimiento: false }),
      PuntajeEntrevista: '', Actualizado: new Date()
    });
    borrarGrabacionDrive_(e.GrabacionURL);
    try { retell_('delete', '/v2/delete-call/' + encodeURIComponent(e.CallID)); } catch (x) { /* best effort */ }
    bitacora_('Sistema', e.PostulacionID, 'Entrevista IA sin consentimiento', 'Grabación y transcripción eliminadas');
    return;
  }
  const valor = { 'Sólida': 1, 'Adecuada': 0.66, 'Débil': 0.33, 'Sin evidencia': 0 };
  const comps = (r.competencias || []).filter(c => valor.hasOwnProperty(c.valoracion));
  const conEvid = comps.filter(c => c.valoracion !== 'Sin evidencia');
  const puntaje = conEvid.length ? Math.round(conEvid.reduce((a, c) => a + valor[c.valoracion], 0) / conEvid.length * 100) : '';
  r.generada = Utilities.formatDate(new Date(), APP.TZ, 'dd/MM/yyyy HH:mm');
  actualizar_('Entrevistas', e._row, {
    Estado: 'Evaluada', EvaluacionIA: JSON.stringify(r).slice(0, 45000), PuntajeEntrevista: puntaje, Actualizado: new Date()
  });
  bitacora_('Sistema', e.PostulacionID, 'Entrevista IA evaluada', 'Puntaje ' + puntaje);
}

/* ============================================================
 *  RÚBRICAS ADICIONALES
 * ============================================================ */

function sembrarRubricasAdicionales() {
  const sh = sh_('Rubricas');
  const existentes = new Set(leer_('Rubricas').map(r => r.RubricaID));
  const R = {
    'RUB-TUM': ['Técnico en Urgencias Médicas (TUM) / Paramédico', [
      ['Formación TUM', 'Técnico en Urgencias Médicas (básico, intermedio o avanzado) concluido, con constancia o título.', 25, 'Sí'],
      ['Experiencia prehospitalaria', 'Experiencia en ambulancia, rescate o atención prehospitalaria; más valor en servicios privados o con turistas.', 20, 'No'],
      ['Certificaciones', 'BLS vigente; PHTLS, ACLS, PALS o AMLS deseables.', 15, 'No'],
      ['Normatividad y protocolos', 'NOM-034-SSA3-2013 (atención médica prehospitalaria), triage, manejo de trauma, traslados.', 10, 'No'],
      ['Licencia de conducir', 'Licencia de chofer vigente o experiencia como operador de ambulancia (si el puesto lo requiere).', 10, 'No'],
      ['Inglés', 'Inglés conversacional para atender turistas.', 10, 'No'],
      ['Salvamento acuático', 'Curso de salvamento acuático o guardavidas (deseable para Beach Patrol).', 5, 'No'],
      ['Disponibilidad', 'Disponibilidad para guardias, turnos rotativos y trabajo en playa o exteriores.', 5, 'No']
    ]],
    'RUB-MED-GRAL': ['Médico General / Urgencias', [
      ['Título y cédula', 'Médico Cirujano titulado con cédula profesional.', 25, 'Sí'],
      ['Experiencia clínica', 'Experiencia en urgencias, hospitalización o consulta; más valor en hospital privado.', 20, 'No'],
      ['Soporte vital', 'ACLS y BLS vigentes; ATLS o PALS deseables.', 15, 'No'],
      ['Inglés', 'Inglés avanzado para atender pacientes extranjeros y seguros internacionales.', 15, 'No'],
      ['Expediente y normatividad', 'NOM-004-SSA3-2012 (expediente clínico), consentimiento informado, notas médicas.', 10, 'No'],
      ['Calidad y seguridad del paciente', 'Metas Internacionales de Seguridad del Paciente, hospitales certificados (CSG/JCI).', 10, 'No'],
      ['Disponibilidad', 'Disponibilidad para guardias y turnos.', 5, 'No']
    ]],
    'RUB-RECEP': ['Recepción / Admisión Hospitalaria Bilingüe', [
      ['Inglés', 'Inglés avanzado (conversación y correo) para atender pacientes extranjeros.', 25, 'Sí'],
      ['Atención al cliente', 'Experiencia en recepción, hotelería, atención a pacientes o servicio al cliente.', 20, 'No'],
      ['Admisión y seguros', 'Admisión hospitalaria, aseguradoras nacionales/internacionales, cartas de garantía.', 15, 'No'],
      ['Herramientas', 'Manejo de computadora, sistemas hospitalarios o de reservas, Office/Google Workspace.', 15, 'No'],
      ['Escolaridad', 'Bachillerato concluido; carrera en turismo, administración o afín deseable.', 10, 'No'],
      ['Caja y facturación', 'Cobros, cortes de caja, facturación.', 10, 'No'],
      ['Disponibilidad', 'Disponibilidad para turnos rotativos, fines de semana y días festivos.', 5, 'No']
    ]],
    'RUB-INTEND': ['Auxiliar de Intendencia Hospitalaria', [
      ['Experiencia en limpieza', 'Experiencia en limpieza; más valor en hospitales, clínicas u hoteles.', 30, 'No'],
      ['RPBI y bioseguridad', 'Manejo de residuos peligrosos biológico-infecciosos (NOM-087-SEMARNAT-SSA1-2002), áreas críticas.', 25, 'No'],
      ['Químicos y técnicas', 'Uso seguro de químicos, diluciones, técnicas de limpieza y desinfección.', 20, 'No'],
      ['Estabilidad', 'Permanencia en empleos previos.', 15, 'No'],
      ['Disponibilidad', 'Disponibilidad para turnos rotativos y fines de semana.', 10, 'No']
    ]],
    'RUB-LAB': ['Químico / Laboratorista Clínico', [
      ['Título y cédula', 'QFB, QBP, QC o Lic. en Laboratorio Clínico titulado con cédula profesional.', 25, 'Sí'],
      ['Experiencia en laboratorio clínico', 'Hematología, química clínica, urianálisis, microbiología, banco de sangre.', 25, 'No'],
      ['Normatividad y calidad', 'NOM-007-SSA3-2011 (laboratorios clínicos), control de calidad interno/externo.', 15, 'No'],
      ['Equipos y sistemas', 'Equipos automatizados, LIS.', 15, 'No'],
      ['Toma de muestras', 'Flebotomía y manejo de muestras.', 10, 'No'],
      ['Disponibilidad', 'Disponibilidad para guardias y turnos.', 10, 'No']
    ]]
  };
  const filas = [];
  Object.keys(R).forEach(id => {
    if (existentes.has(id)) return;
    R[id][1].forEach(c => filas.push([id, R[id][0], c[0], c[1], c[2], c[3]]));
  });
  if (filas.length) sh.getRange(sh.getLastRow() + 1, 1, filas.length, 6).setValues(filas);
  Logger.log('Rúbricas agregadas: ' + filas.length + ' criterios.');
}

/* ============================================================
 *  DEPURACIÓN DE DATOS (conforme al aviso de privacidad)
 * ============================================================ */

/**
 * Lista (o elimina, con ejecutar=true) los datos que ya cumplieron su plazo de conservación:
 *  - Sin autorización de banco de talento: 90 días después de cerrar su proceso (Descartado).
 *  - Con banco de talento: MESES_BANCO meses (24 por defecto) después de su última postulación.
 *  Nunca toca candidatos con un proceso abierto ni a los Seleccionados (pasan a expediente laboral).
 *  Los archivos van a la papelera de Drive (recuperables 30 días).
 */
function depurarDatos(ejecutar) {
  const meses = Number(prop_('MESES_BANCO', '24')) || 24;
  const ahora = Date.now(), DIA = 86400000;
  const posts = leer_('Postulaciones'), ents = leer_('Entrevistas');
  const elegibles = [];
  leer_('Candidatos').forEach(c => {
    if (c.Nombre === '[Datos eliminados]') return;
    const ps = posts.filter(p => p.CandidatoID === c.ID);
    if (!ps.length) return;
    if (ps.some(p => p.Etapa !== 'Descartado')) return;
    const ultima = Math.max.apply(null, ps.map(p => new Date(p.Actualizado || p.Fecha).getTime()));
    const dias = (ahora - ultima) / DIA;
    const limite = c.BancoTalento === 'No' ? 90 : meses * 30.4;
    if (dias >= limite) elegibles.push({ c, ps, dias: Math.round(dias) });
  });
  Logger.log((ejecutar ? 'Depurando ' : 'Vista previa — se depurarían ') + elegibles.length + ' candidato(s).');
  elegibles.forEach(x => Logger.log('· ' + x.c.ID + ' · ' + x.dias + ' días · banco ' + (x.c.BancoTalento || 'Sí')));
  if (ejecutar) elegibles.forEach(x => anonimizar_(x.c, x.ps, ents, 'Plazo de conservación'));
  return elegibles.length;
}

/** Activa la depuración automática el día 1 de cada mes a las 3:00. */
function activarDepuracionMensual() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'depuracionMensual_')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('depuracionMensual_').timeBased().onMonthDay(1).atHour(3).create();
  Logger.log('Depuración mensual activada.');
}
function depuracionMensual_() { depurarDatos(true); }

/** Derecho de cancelación (ARCO) desde el panel. */
function cancelarDatosCandidato(tok, candId) {
  const u = usuario_(tok);
  const c = buscar_('Candidatos', candId);
  const ps = leer_('Postulaciones').filter(p => p.CandidatoID === candId);
  anonimizar_(c, ps, leer_('Entrevistas'), 'Solicitud del titular · ' + u);
  return { ok: true };
}

function anonimizar_(c, ps, ents, motivo) {
  const archivos = new Set([c.CV_FileId].concat(ps.map(p => p.CV_FileId)).filter(Boolean));
  archivos.forEach(id => { try { DriveApp.getFileById(id).setTrashed(true); } catch (x) { } });
  actualizar_('Candidatos', c._row, {
    Nombre: '[Datos eliminados]', Email: '', Telefono: '', Ciudad: '', Cedula: '', CV_URL: '', CV_FileId: '', ResumenIA: ''
  });
  ps.forEach(p => {
    actualizar_('Postulaciones', p._row, { FichaIA: '', CV_URL: '', CV_FileId: '', Notas: '', EstadoFicha: 'Eliminada', Actualizado: new Date() });
    ents.filter(e => e.PostulacionID === p.ID).forEach(e => {
      borrarGrabacionDrive_(e.GrabacionURL);
      if (e.CallID) { try { retell_('delete', '/v2/delete-call/' + encodeURIComponent(e.CallID)); } catch (x) { } }
      actualizar_('Entrevistas', e._row, { Token: '', Transcripcion: '', GrabacionURL: '', EvaluacionIA: '', Actualizado: new Date() });
    });
    bitacora_('Sistema', p.ID, 'Datos eliminados', motivo);
  });
}

function borrarGrabacionDrive_(url) {
  const m = String(url || '').match(/\/d\/([\w-]+)/);
  if (m) { try { DriveApp.getFileById(m[1]).setTrashed(true); } catch (x) { } }
}

/* ============================================================
 *  UTILIDADES
 * ============================================================ */

function retell_(method, path, body) {
  const key = prop_('RETELL_API_KEY');
  if (!key) throw new Error('Falta RETELL_API_KEY en las Propiedades del script.');
  const opt = { method, muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + key } };
  if (body !== undefined) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(body); }
  const res = UrlFetchApp.fetch(RETELL + path, opt);
  const code = res.getResponseCode(), txt = res.getContentText();
  if (code < 200 || code >= 300) throw new Error('Retell (' + code + '): ' + txt.slice(0, 300));
  return txt ? JSON.parse(txt) : {};
}

function reintentosEval_(e) {
  try { return Number((JSON.parse(e.EvaluacionIA || '{}').reintentos) || 0); } catch (x) { return 0; }
}

/**
 * Diagnóstico: ejecútalo desde el editor y revisa el registro.
 * Comprueba propiedades, hoja Entrevistas, disparadores, el asistente en Retell y la URL pública.
 */
function verificarFase2() {
  const ok = [], mal = [];
  const chk = (cond, txt, ayuda) => (cond ? ok : mal).push(cond ? txt : txt + ' → ' + ayuda);

  ['ANTHROPIC_API_KEY', 'RETELL_API_KEY', 'RETELL_AGENT_ID', 'CV_FOLDER_ID'].forEach(k =>
    chk(!!prop_(k), 'Propiedad ' + k, 'agrégala en Configuración del proyecto › Propiedades del script'));
  chk(!!prop_('AVISO_PRIVACIDAD_URL'), 'Propiedad AVISO_PRIVACIDAD_URL', 'recomendada: la página de entrevista enlaza el aviso de privacidad');

  const sh = ss_().getSheetByName('Entrevistas');
  chk(!!sh, 'Hoja Entrevistas', 'ejecuta setupFase2()');
  if (sh) {
    const enc = sh.getRange(1, 1, 1, HOJAS.Entrevistas.length).getValues()[0];
    chk(enc.join('|') === HOJAS.Entrevistas.join('|'), 'Encabezados de Entrevistas', 'ejecuta setupFase2() para reescribirlos');
  }
  const trig = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  chk(trig.indexOf('procesarEntrevistas') >= 0, 'Disparador procesarEntrevistas (cada 5 min)', 'ejecuta setupFase2()');
  chk(trig.indexOf('depuracionMensual_') >= 0, 'Depuración mensual', 'opcional: ejecuta activarDepuracionMensual()');

  if (prop_('RETELL_API_KEY') && prop_('RETELL_AGENT_ID')) {
    try {
      const a = retell_('get', '/get-agent/' + encodeURIComponent(prop_('RETELL_AGENT_ID')));
      chk(true, 'Asistente Retell: ' + (a.agent_name || a.agent_id) + ' · voz ' + a.voice_id + ' · idioma ' + a.language);
    } catch (x) { chk(false, 'Asistente Retell', 'no se encontró (' + x.message + '); ejecuta crearAgenteRetell()'); }
  }

  const url = prop_('URL_ENTREVISTA', ENT.URL_DEFECTO);
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const html = r.getContentText();
    chk(r.getResponseCode() === 200, 'Página pública ' + url, 'no responde (' + r.getResponseCode() + '); súbela a GitHub Pages');
    if (r.getResponseCode() === 200) {
      const api = (html.match(/const API = '([^']+)'/) || [])[1];
      chk(!!api && api === ScriptApp.getService().getUrl(), 'La página apunta a esta app web',
        'en entrevista.html, la constante API debe ser ' + ScriptApp.getService().getUrl() + ' (actual: ' + api + ')');
      const sdk = (html.match(/<script src="([^"]*retell[^"]*)"/i) || [])[1];
      if (sdk) {
        const s = UrlFetchApp.fetch(URL_(url, sdk), { muteHttpExceptions: true });
        chk(s.getResponseCode() === 200, 'SDK de Retell (' + sdk + ')', 'no está publicado junto a entrevista.html');
      }
    }
  } catch (x) { chk(false, 'Página pública', x.message); }

  Logger.log('✅ ' + ok.length + ' correcto(s):\n  ' + ok.join('\n  '));
  Logger.log(mal.length ? '⚠️ ' + mal.length + ' pendiente(s):\n  ' + mal.join('\n  ') : 'Todo listo para entrevistas por voz.');
  return mal.length;
}

/** Une una URL base con una ruta relativa (Apps Script no tiene la clase URL). */
function URL_(base, rel) {
  if (/^https?:\/\//i.test(rel)) return rel;
  return base.replace(/[?#].*$/, '').replace(/[^\/]*$/, '') + rel.replace(/^\.\//, '');
}

function configVozOk_() {
  if (!prop_('RETELL_API_KEY')) throw new Error('La entrevista por voz no está activa: falta RETELL_API_KEY.');
  if (!prop_('RETELL_AGENT_ID')) throw new Error('Falta crear el asistente de voz: ejecuta crearAgenteRetell().');
}

function e164_(tel) {
  const d = String(tel || '').replace(/\D/g, '');
  if (d.length === 10) return '+52' + d;
  if (d.length === 12 && d.indexOf('52') === 0) return '+' + d;
  if (d.length === 13 && d.indexOf('521') === 0) return '+52' + d.slice(3);
  return d.length >= 11 ? '+' + d : '';
}
