// ============================================================
//  Google Apps Script — Cruz Verde Visitas
//  Hoja: "Visitas Cruz Verde" (o el nombre de tu Spreadsheet)
//  Pestañas requeridas: "Visitas" e "Incidencias"
// ============================================================

// ID del Google Sheet de Cruz Verde
const SHEET_ID = '1-f9BSx6ZNUlY-5e1FsY_a1gWAmDnmv51JsFC7Mu8gls';

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

      // ★ 7. Actualizar estado de incidencia
      case 'update_estado':
        return updateEstadoIncidencia(data);

      // ★ 8. Inicializar encabezados si faltan
      case 'init_headers':
        return initHeaders();

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
    d.com_est_estado || '',                 // X
    d.com_dinamica || '',                   // Y
    d.com_medios || '',                     // Z
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
  if (data.length === 0) return jsonOk({ data: [] });

  // Detectar si la primera fila es encabezado o datos
  // Si la primera celda es un número (timestamp ID), es datos, no encabezado
  const firstCell = data[0][0];
  const hasHeader = (typeof firstCell === 'string' && isNaN(Number(firstCell)));
  return jsonOk({ data: hasHeader ? data.slice(1) : data });
}


// ============================================================
//  6. LISTAR INCIDENCIAS
// ============================================================
function listarIncidencias() {
  const sh = getSheet('Incidencias');
  if (!sh) return jsonError('No se encontró la pestaña "Incidencias"');

  const data = sh.getDataRange().getValues();
  if (data.length === 0) return jsonOk({ data: [] });

  // Detectar si la primera fila es encabezado o datos
  const firstCell = data[0][0];
  const hasHeader = (typeof firstCell === 'string' && isNaN(Number(firstCell)));
  return jsonOk({ data: hasHeader ? data.slice(1) : data });
}


// ============================================================
//  7. ACTUALIZAR ESTADO DE INCIDENCIA
// ============================================================
function updateEstadoIncidencia(d) {
  const sh = getSheet('Incidencias');
  if (!sh) return jsonError('No se encontró la pestaña "Incidencias"');

  const visitaId = String(d.visita_id);
  const tipo = (d.tipo || '').trim();
  const descripcion = (d.descripcion || '').trim();
  const nuevoEstado = d.nuevo_estado || 'Pendiente';
  const data = sh.getDataRange().getValues();

  // Detectar formato de columnas (tab en col B vs col E)
  const _tabs = ['sala','bodegas','comercial','general','protocolo'];

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) !== visitaId) continue;

    // Detectar si esta fila tiene tab en posición 1 o 4
    const tabInB = _tabs.includes(String(data[i][1]).toLowerCase());
    const colTipo = tabInB ? 5 : 5;       // F en ambos formatos
    const colDesc = tabInB ? 6 : 6;       // G en ambos formatos
    const colEstado = tabInB ? 11 : 11;   // L en ambos formatos

    const rowTipo = (String(data[i][colTipo]) || '').trim();
    const rowDesc = (String(data[i][colDesc]) || '').trim();

    if (rowTipo === tipo && rowDesc.substring(0, 50) === descripcion.substring(0, 50)) {
      sh.getRange(i + 1, colEstado + 1).setValue(nuevoEstado);
      return jsonOk({});
    }
  }
  return jsonError('Incidencia no encontrada');
}


// ============================================================
//  8. INICIALIZAR ENCABEZADOS (ejecutar una sola vez)
// ============================================================
function initHeaders() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Encabezados de Visitas
  const shV = ss.getSheetByName('Visitas');
  if (shV) {
    const data = shV.getDataRange().getValues();
    const firstCell = data.length > 0 ? data[0][0] : '';
    // Solo agregar headers si no existen
    if (typeof firstCell !== 'string' || !isNaN(Number(firstCell))) {
      const headers = [
        'ID','Timestamp','Auditor','Codigo','Nombre','Ciudad','Lider',
        'Fecha_Visita','Hora_Visita',
        'Sala_Piso','Sala_Limpieza','Sala_Exhibicion','Sala_POP','Sala_Obs',
        'Bod_Org','Bod_Inv','Bod_Cond','Bod_Seg','Bod_Obs',
        'Com_Metas','Com_Flujo','Com_Servicio','Com_Estrategias',
        'Com_Est_Estado','Com_Dinamica','Com_Medios','Com_Obs',
        'Gen_Infra','Gen_Personal','Gen_Protocolos','Gen_Calificacion','Gen_Obs',
        'Prot_Saludo','Prot_Ofrecimiento','Prot_Bienvenida','Prot_Calificacion','Prot_Obs',
        'Compromisos','Proxima_Visita','Firma',
        'Num_Incidencias','Incidencias_Alta','Fotos_URLs'
      ];
      shV.insertRowBefore(1);
      shV.getRange(1, 1, 1, headers.length).setValues([headers]);
      shV.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }

  // Encabezados de Incidencias
  const shI = ss.getSheetByName('Incidencias');
  if (shI) {
    const data = shI.getDataRange().getValues();
    const firstCell = data.length > 0 ? data[0][0] : '';
    if (typeof firstCell !== 'string' || !isNaN(Number(firstCell))) {
      const headers = [
        'Visita_ID','Tab','Codigo','Nombre','Ciudad',
        'Tipo','Descripcion','Criticidad','Responsable','Area',
        'Fecha_Seguimiento','Estado','Fotos_URLs'
      ];
      shI.insertRowBefore(1);
      shI.getRange(1, 1, 1, headers.length).setValues([headers]);
      shI.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }

  return jsonOk({ message: 'Headers inicializados correctamente' });
}
