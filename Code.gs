// ============================================================
//  Google Apps Script — Cruz Verde Visitas
//  Hoja: "Visitas Cruz Verde" (o el nombre de tu Spreadsheet)
//  Pestañas requeridas: "Visitas" e "Incidencias"
// ============================================================

// ★ CAMBIA ESTE ID por el de tu Google Sheet ★
const SHEET_ID = 'TU_SHEET_ID_AQUI';

// ─── Helpers ───
function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function jsonOk(data)    { return ContentService.createTextOutput(JSON.stringify({ status: 'ok', ...data })).setMimeType(ContentService.MimeType.JSON); }
function jsonError(msg)  { return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: msg })).setMimeType(ContentService.MimeType.JSON); }

// ─── Punto de entrada GET ───
function doGet(e) {
  try {
    const raw  = e.parameter.data || '{}';
    const data = JSON.parse(raw);
    const type = data.type || '';

    switch (type) {

      // ★ 1. Guardar nueva visita
      case 'visita':
        return guardarVisita(data);

      // ★ 2. Guardar incidencia
      case 'incidencia':
        return guardarIncidencia(data);

      // ★ 3. Actualizar URLs de fotos de visita (Cloudinary)
      case 'update_fotos':
        return updateFotosVisita(data);

      // ★ 4. Actualizar URLs de fotos de incidencia
      case 'update_inc_fotos':
        return updateFotosIncidencia(data);

      // ★ 5. Listar todas las visitas (para dashboard y revisitas)
      case 'visitas':
        return listarVisitas();

      // ★ 6. Listar incidencias
      case 'incidencias':
        return listarIncidencias();

      default:
        return jsonError('Tipo no reconocido: ' + type);
    }
  } catch (err) {
    return jsonError(err.message);
  }
}

// También aceptar POST por si acaso
function doPost(e) {
  return doGet(e);
}


// ============================================================
//  1. GUARDAR VISITA
// ============================================================
function guardarVisita(d) {
  const sh = getSheet('Visitas');
  if (!sh) return jsonError('No se encontró la pestaña "Visitas"');

  // Columnas de la pestaña Visitas (43 columnas)
  const row = [
    d.id || Date.now().toString(),         // A - ID
    d.timestamp || new Date().toISOString(),// B - Timestamp
    d.auditor || '',                        // C - Auditor
    d.codigo || '',                         // D - Código local
    (d.nombre || '').trim(),                // E - Nombre droguería
    (d.ciudad || '').trim(),                // F - Ciudad
    d.lider || '',                          // G - Líder
    d.fecha_visita || '',                   // H - Fecha visita
    d.hora_visita || '',                    // I - Hora visita
    // Sala de venta
    d.sala_piso || '',                      // J
    d.sala_limpieza || '',                  // K
    d.sala_exhibicion || '',                // L
    d.sala_pop || '',                       // M
    d.sala_obs || '',                       // N
    // Bodegas
    d.bod_org || '',                        // O
    d.bod_inv || '',                        // P
    d.bod_cond || '',                       // Q
    d.bod_seg || '',                        // R
    d.bod_obs || '',                        // S
    // Comercial
    d.com_metas || '',                      // T
    d.com_flujo || '',                      // U
    d.com_servicio || '',                   // V
    d.com_estrategias || '',                // W
    d.com_est_estado || '',                 // X  ← NUEVO
    d.com_dinamica || '',                   // Y  ← NUEVO
    d.com_medios || '',                     // Z  ← NUEVO
    d.com_obs || '',                        // AA
    // General
    d.gen_infra || '',                      // AB
    d.gen_personal || '',                   // AC
    d.gen_protocolos || '',                 // AD
    d.gen_calificacion || '',               // AE
    d.gen_obs || '',                        // AF
    // Protocolo
    d.prot_saludo || '',                    // AG
    d.prot_ofrecimiento || '',              // AH
    d.prot_bienvenida || '',                // AI
    d.prot_calificacion || '',              // AJ
    d.prot_obs || '',                       // AK
    // Cierre
    d.compromisos || '',                    // AL
    d.proxima_visita || '',                 // AM
    d.firma || '',                          // AN
    // Métricas
    d.num_incidencias || 0,                 // AO
    d.incidencias_alta || 0,                // AP
    d.fotos_urls || '',                     // AQ - URLs de fotos (se llena después)
  ];

  sh.appendRow(row);
  return jsonOk({ id: row[0] });
}


// ============================================================
//  2. GUARDAR INCIDENCIA
// ============================================================
function guardarIncidencia(d) {
  const sh = getSheet('Incidencias');
  if (!sh) return jsonError('No se encontró la pestaña "Incidencias"');

  // Columnas de la pestaña Incidencias (13 columnas)
  const row = [
    d.visita_id || '',        // A - ID de visita padre
    d.tab || 'general',       // B - Tab/sección
    d.codigo || '',            // C - Código local
    d.nombre || '',            // D - Nombre droguería
    (d.ciudad || '').trim(),   // E - Ciudad
    d.tipo || '',              // F - Tipo incidencia
    d.descripcion || '',       // G - Descripción
    d.criticidad || '',        // H - Criticidad (Alta/Media/Baja)
    d.responsable || '',       // I - Responsable
    d.area || '',              // J - Área
    d.fecha_seguimiento || '', // K - Fecha límite
    d.estado || 'Pendiente',   // L - Estado
    '',                        // M - Fotos URLs (se llena después)
  ];

  sh.appendRow(row);
  return jsonOk({});
}


// ============================================================
//  3. ACTUALIZAR FOTOS VISITA (después de subir a Cloudinary)
// ============================================================
function updateFotosVisita(d) {
  const sh = getSheet('Visitas');
  if (!sh) return jsonError('No se encontró la pestaña "Visitas"');

  const visitaId = String(d.visita_id);
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === visitaId) {
      // Columna AQ (índice 42) = fotos_urls
      const colFotos = 43; // columna 43 en base-1
      const existing = sh.getRange(i + 1, colFotos).getValue() || '';
      const newUrls = existing ? existing + '|' + d.fotos_urls : d.fotos_urls;
      sh.getRange(i + 1, colFotos).setValue(newUrls);
      return jsonOk({});
    }
  }
  return jsonError('Visita no encontrada: ' + visitaId);
}


// ============================================================
//  4. ACTUALIZAR FOTOS INCIDENCIA
// ============================================================
function updateFotosIncidencia(d) {
  const sh = getSheet('Incidencias');
  if (!sh) return jsonError('No se encontró la pestaña "Incidencias"');

  const visitaId = String(d.visita_id);
  const data = sh.getDataRange().getValues();

  // Buscar la incidencia por visita_id (col A) - actualizamos la primera que no tenga fotos
  // o la que coincida con el inc_n
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === visitaId) {
      const colFotos = 13; // columna M en base-1
      const existing = sh.getRange(i + 1, colFotos).getValue() || '';
      if (!existing) {
        sh.getRange(i + 1, colFotos).setValue(d.fotos_urls);
        return jsonOk({});
      }
    }
  }
  // Si todas ya tienen fotos, agregar a la última
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === visitaId) {
      const colFotos = 13;
      const existing = sh.getRange(i + 1, colFotos).getValue() || '';
      const newUrls = existing ? existing + '|' + d.fotos_urls : d.fotos_urls;
      sh.getRange(i + 1, colFotos).setValue(newUrls);
      return jsonOk({});
    }
  }
  return jsonError('Incidencia no encontrada para visita: ' + visitaId);
}


// ============================================================
//  5. LISTAR VISITAS (dashboard + revisitas)
// ============================================================
function listarVisitas() {
  const sh = getSheet('Visitas');
  if (!sh) return jsonError('No se encontró la pestaña "Visitas"');

  const data = sh.getDataRange().getValues();
  // Retornar todo excepto la fila de encabezado
  return jsonOk({ data: data.slice(1) });
}


// ============================================================
//  6. LISTAR INCIDENCIAS
// ============================================================
function listarIncidencias() {
  const sh = getSheet('Incidencias');
  if (!sh) return jsonError('No se encontró la pestaña "Incidencias"');

  const data = sh.getDataRange().getValues();
  return jsonOk({ data: data.slice(1) });
}
