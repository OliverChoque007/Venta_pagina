// ─── CONFIGURACIÓN ───────────────────────────────────────────────────
const CONFIG = {
  SHEET_ID:         '1ZvvTMBnOB_85QkURzBijESdg_ZbdyDXw-fBzotJA_no',       
  SHEET_NAME:       'Solicitudes',
  DRIVE_FOLDER:     'DevStudio_Solicitudes',          // Carpeta principal en tu Drive
  DRIVE_QR_FOLDER:  'DevStudio_QR_Pagos',            // Carpeta donde subirás el QR de pago
  ADMIN_EMAIL:      'oliverkhan709@gmail.com',
  ADMIN_WHATSAPP:   '59174050023',                    // Sin el +
  NOMBRE_EMPRESA:   'DevStudio'
};

// ═══════════════════════════════════════════════════════════════════════
//  PUNTO DE ENTRADA — GET (prueba rápida en el navegador)
// ═══════════════════════════════════════════════════════════════════════
function doGet() {
  return ContentService
    .createTextOutput('✅ DevStudio Apps Script activo — Sistema de Solicitudes')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ═══════════════════════════════════════════════════════════════════════
//  PUNTO DE ENTRADA — POST (recibe datos del HTML)
// ═══════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 1. Guardar archivos del usuario en Drive y obtener URL de la carpeta
    const folderUrl = guardarArchivosEnDrive(data);

    // 2. Registrar en Google Sheets
    const nSolicitud = registrarEnSheet(data, folderUrl);

    // 3. Notificar al admin
    notificarAdmin(data, nSolicitud, folderUrl);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, n: nSolicitud }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('[ERROR doPost] ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════════════════════════════════════
//  GOOGLE DRIVE — Guardar archivos del usuario
// ═══════════════════════════════════════════════════════════════════════
function guardarArchivosEnDrive(data) {
  // Obtener/crear carpeta principal
  let carpetaPrincipal = obtenerOCrearCarpeta(CONFIG.DRIVE_FOLDER, null);

  // Crear subcarpeta para esta solicitud: "Solicitud_NombreUsuario_Fecha"
  const fecha    = new Date();
  const fechaStr = Utilities.formatDate(fecha, 'America/La_Paz', 'yyyyMMdd_HHmmss');
  const subName  = 'SOL_' + fechaStr + '_' + data.nombre.replace(/\s+/g, '_').substring(0, 20);
  let subcarpeta = carpetaPrincipal.createFolder(subName);

  // Guardar archivos si los hay
  if (data.archivos && data.archivos.length > 0) {
    data.archivos.forEach(arch => {
      try {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(arch.datos),
          arch.tipo || 'application/octet-stream',
          arch.nombre
        );
        subcarpeta.createFile(blob);
        Logger.log('[DRIVE] Archivo guardado: ' + arch.nombre);
      } catch (e) {
        Logger.log('[DRIVE] Error guardando ' + arch.nombre + ': ' + e.toString());
      }
    });
  }

  // Crear un archivo TXT resumen de requerimientos en la carpeta
  const resumen = buildResumenTxt(data);
  subcarpeta.createFile('_RESUMEN_SOLICITUD.txt', resumen, MimeType.PLAIN_TEXT);

  Logger.log('[DRIVE] Carpeta creada: ' + subcarpeta.getUrl());
  return subcarpeta.getUrl();
}

// ─── Auxiliar: obtener o crear carpeta por nombre ─────────────────────
function obtenerOCrearCarpeta(nombre, padre) {
  const buscador = padre
    ? padre.getFoldersByName(nombre)
    : DriveApp.getFoldersByName(nombre);
  return buscador.hasNext()
    ? buscador.next()
    : (padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre));
}

// ─── Resumen TXT del pedido ───────────────────────────────────────────
function buildResumenTxt(data) {
  const ts = Utilities.formatDate(new Date(), 'America/La_Paz', 'dd/MM/yyyy HH:mm');
  return [
    '════════════════════════════════════════',
    '  DEVSTUDIO — RESUMEN DE SOLICITUD',
    '════════════════════════════════════════',
    'Fecha          : ' + ts,
    'Nombre         : ' + data.nombre,
    'Correo         : ' + data.correo,
    'WhatsApp       : ' + data.whatsapp,
    'Proyecto       : ' + data.proyecto,
    'Servicio       : ' + data.servicio,
    'Plataforma     : ' + (data.plataforma || 'No especificado'),
    'Diseño         : ' + (data.diseno || 'No especificado'),
    'Funcionalidades: ' + (data.funciones || 'No especificado'),
    'Presupuesto    : ' + (data.presupuesto || 'A consultar'),
    'Plazo          : ' + (data.plazo || 'Flexible'),
    '────────────────────────────────────────',
    'REQUERIMIENTOS:',
    (data.requerimientos || '(Sin requerimientos adicionales)'),
    '════════════════════════════════════════'
  ].join('\n');
}

// ═══════════════════════════════════════════════════════════════════════
//  GOOGLE SHEETS — Registrar solicitud
// ═══════════════════════════════════════════════════════════════════════
function registrarEnSheet(data, folderUrl) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // ── Crear hoja con encabezados si no existe ──────────────────────────
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);

    const headers = [
      'N',
      'TIMESTAMP',
      'NOMBRE_USUARIO',
      'CORREO_USUARIO',
      'WHATSAPP',
      'URL',
      'REQUERIMIENTOS',
      'ESTADO',
      'OBS_EMPLEADO',
      'MONTO',
      'QR',
      'ENVIO_WHATSAPP'
    ];

    // Estilizar encabezados
    const hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setValues([headers]);
    hr.setFontWeight('bold');
    hr.setFontColor('#1e1e2e');
    hr.setBackground('#e0e7ff');
    hr.setFontSize(10);
    hr.setHorizontalAlignment('center');
    hr.setBorder(true, true, true, true, true, true,
                 '#c7d2fe', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setFrozenRows(1);

    // Anchos de columna
    sheet.setColumnWidth(1,  60);   // N
    sheet.setColumnWidth(2,  160);  // TIMESTAMP
    sheet.setColumnWidth(3,  180);  // NOMBRE
    sheet.setColumnWidth(4,  200);  // CORREO
    sheet.setColumnWidth(5,  130);  // WHATSAPP
    sheet.setColumnWidth(6,  220);  // URL
    sheet.setColumnWidth(7,  300);  // REQUERIMIENTOS
    sheet.setColumnWidth(8,  110);  // ESTADO
    sheet.setColumnWidth(9,  200);  // OBS_EMPLEADO
    sheet.setColumnWidth(10, 100);  // MONTO
    sheet.setColumnWidth(11, 220);  // QR
    sheet.setColumnWidth(12, 180);  // ENVIO_WHATSAPP

    // Validación de datos para columna ESTADO (col 8)
    // Empieza desde fila 2 hasta fila 1000
    const estadoRange = sheet.getRange(2, 8, 999, 1);
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(['PENDIENTE', 'ACEPTADO', 'AGENDAR CITA'], true)
      .setAllowInvalid(false)
      .build();
    estadoRange.setDataValidation(regla);
  }

  // ── Calcular N autoincremental ───────────────────────────────────────
  const lastRow = sheet.getLastRow();
  const n = lastRow >= 1 ? lastRow : 1; // fila 1 = encabezado, entonces N = lastRow (que ya cuenta encabezado)

  // ── Construir texto de requerimientos completo ───────────────────────
  const reqTexto = [
    data.plataforma    ? '▸ Plataforma: '     + data.plataforma    : '',
    data.diseno        ? '▸ Diseño: '         + data.diseno        : '',
    data.funciones     ? '▸ Funciones: '      + data.funciones     : '',
    data.presupuesto   ? '▸ Presupuesto: '    + data.presupuesto   : '',
    data.plazo         ? '▸ Plazo: '          + data.plazo         : '',
    data.requerimientos? '▸ Notas: '          + data.requerimientos : '',
  ].filter(Boolean).join('\n');

  // ── Fórmula ENVIO_WHATSAPP dinámica ─────────────────────────────────
  // Construye el enlace wa.me usando los valores de la misma fila
  // Columnas: N=A, TIMESTAMP=B, NOMBRE=C, CORREO=D, WHATSAPP=E,
  //           URL=F, REQS=G, ESTADO=H, OBS=I, MONTO=J, QR=K, WSP=L
  const fila = lastRow + 1;

  // Fórmula que genera el link de WhatsApp dinámicamente.
  // Detecta si el ESTADO es "AGENDAR CITA" para enviar mensaje diferente.
  const formula = `=IF(E${fila}="","",`
    + `IF(H${fila}="AGENDAR CITA",`
    // Mensaje AGENDAR CITA
    + `HYPERLINK("https://wa.me/"&SUBSTITUTE(E${fila},"+","")&"?text="&ENCODEURL(`
    + `"Hola " & C${fila} & ", tu solicitud de desarrollo está en estado de '" & H${fila} & "'. ")`
    + `&ENCODEURL("Por favor, indícanos tu disponibilidad respondiendo: ")`
    + `&ENCODEURL("📅 Fecha: (escribe la fecha) | 🕐 Hora: (escribe la hora) | 💻 Modalidad: Virtual o Presencial"),`
    + `"📅 Agendar Cita"),`
    // Mensaje NORMAL (PENDIENTE / ACEPTADO)
    + `HYPERLINK("https://wa.me/"&SUBSTITUTE(E${fila},"+","")&"?text="&ENCODEURL(`
    + `"Hola " & C${fila} & ", tu solicitud de desarrollo está en estado de '" & H${fila} & "'. ")`
    + `&IF(I${fila}<>"",ENCODEURL("Nota: " & I${fila} & ". "),"")` 
    + `&IF(J${fila}<>"",ENCODEURL("La cotización es de Bs. " & J${fila} & ". "),"")` 
    + `&IF(K${fila}<>"",ENCODEURL("El QR de pago: " & K${fila}),"")`
    + `,"💬 Enviar WhatsApp")))`;

  // ── Insertar fila ─────────────────────────────────────────────────────
  const newRow = [
    n,                          // A: N autoincremental
    new Date(),                 // B: TIMESTAMP
    data.nombre,                // C: NOMBRE_USUARIO
    data.correo,                // D: CORREO_USUARIO
    data.whatsapp,              // E: WHATSAPP
    folderUrl,                  // F: URL carpeta Drive
    reqTexto,                   // G: REQUERIMIENTOS
    '',                         // H: ESTADO (vacío por defecto, dropdown)
    '',                         // I: OBS_EMPLEADO (vacío, lo llena el empleado)
    '',                         // J: MONTO (vacío, lo llena el empleado)
    '',                         // K: QR (vacío, lo llena el empleado)
    formula                     // L: ENVIO_WHATSAPP (hipervínculo dinámico)
  ];

  sheet.appendRow(newRow);

  // ── Estilo de la nueva fila ───────────────────────────────────────────
  const rowRange = sheet.getRange(fila, 1, 1, 12);
  rowRange.setBackground('#ffffff');
  rowRange.setFontColor('#1e1e2e');
  rowRange.setFontSize(10);
  rowRange.setBorder(false, false, true, false, false, false,
                     '#e5e7eb', SpreadsheetApp.BorderStyle.SOLID);

  // Columna N en negrita azul
  sheet.getRange(fila, 1).setFontWeight('bold').setFontColor('#4f46e5');

  // Columna URL como hipervínculo
  const urlCell = sheet.getRange(fila, 6);
  urlCell.setFormula(`=HYPERLINK("${folderUrl}","📁 Ver archivos")`);
  urlCell.setFontColor('#4f46e5');

  // Columna ESTADO — fondo amarillo suave para recordar que está vacía
  sheet.getRange(fila, 8).setBackground('#fefce8').setFontColor('#92400e');

  // Wrap text en columna REQUERIMIENTOS
  sheet.getRange(fila, 7).setWrap(true);

  Logger.log('[SHEETS] Fila registrada N°' + n);
  return n;
}

// ═══════════════════════════════════════════════════════════════════════
//  GMAIL — Notificar al administrador
// ═══════════════════════════════════════════════════════════════════════
function notificarAdmin(data, n, folderUrl) {
  const subject = `🆕 Nueva solicitud #${n} — ${data.nombre} | ${data.servicio}`;

  const html = `
<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="background:#0c0c0c;font-family:Courier New,monospace;padding:32px;">
<div style="max-width:580px;margin:0 auto;background:#141414;border:1px solid #333;padding:32px;">
  <div style="border-bottom:2px solid #e8ff47;padding-bottom:16px;margin-bottom:24px;">
    <h1 style="color:#e8ff47;font-size:1.4rem;margin:0;">DevStudio</h1>
    <p style="color:#888;font-size:11px;margin:4px 0 0;">Nueva solicitud registrada</p>
  </div>
  <table style="width:100%;border-collapse:collapse;">
    ${fila('N° Solicitud',  '#' + n)}
    ${fila('Nombre',        data.nombre)}
    ${fila('Correo',        data.correo)}
    ${fila('WhatsApp',      data.whatsapp)}
    ${fila('Proyecto',      data.proyecto)}
    ${fila('Servicio',      data.servicio)}
    ${fila('Plataforma',    data.plataforma || '—')}
    ${fila('Presupuesto',   data.presupuesto || 'A consultar')}
    ${fila('Plazo',         data.plazo || 'Flexible')}
    ${fila('Archivos',      data.archivos && data.archivos.length > 0 ? data.archivos.length + ' archivo(s)' : 'Ninguno')}
  </table>
  <div style="margin-top:20px;background:#0c0c0c;padding:16px;border-left:3px solid #e8ff47;">
    <p style="color:#888;font-size:10px;text-transform:uppercase;letter-spacing:2px;margin:0 0 8px;">Requerimientos</p>
    <p style="color:#f5f5f0;font-size:12px;line-height:1.8;margin:0;">${(data.requerimientos || '(Sin requerimientos adicionales)').replace(/\n/g,'<br>')}</p>
  </div>
  <div style="margin-top:20px;">
    <a href="${folderUrl}" style="display:inline-block;background:#e8ff47;color:#0c0c0c;padding:12px 24px;font-weight:bold;font-size:12px;text-decoration:none;">
      📁 Ver archivos en Drive
    </a>
  </div>
  <p style="color:#444;font-size:10px;margin-top:24px;">DevStudio — Sistema de Solicitudes</p>
</div>
</body></html>`;

  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, '', { htmlBody: html, name: 'DevStudio Sistema' });
  Logger.log('[GMAIL] Notificación enviada al admin para solicitud #' + n);
}

// ─── Auxiliar HTML tabla ──────────────────────────────────────────────
function fila(label, valor) {
  return `<tr>
    <td style="color:#888;font-size:10px;text-transform:uppercase;letter-spacing:1px;padding:8px 0;border-bottom:1px solid #222;width:130px;">${label}</td>
    <td style="color:#f5f5f0;font-size:12px;padding:8px 0;border-bottom:1px solid #222;">${valor}</td>
  </tr>`;
}