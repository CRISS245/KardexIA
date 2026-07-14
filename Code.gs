function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('KardexIA Lab - Sistema de Kardex Digital')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==========================================
// OBTENER DATOS DEL KARDEX
// ==========================================
function getKardexData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventario");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const rows = data.slice(1);
  const formattedData = rows.map(row => {
    let fecha = row[0];
    if (fecha instanceof Date) fecha = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    let hora = row[1];
    if (hora instanceof Date) hora = Utilities.formatDate(hora, Session.getScriptTimeZone(), "hh:mm a");
    
    return {
      fecha: fecha || '',
      hora: hora || '',
      tipo: String(row[2] || ''),
      codigo: String(row[3] || ''),
      nombre: String(row[4] || ''),
      tipo2: String(row[5] || ''),
      lote: String(row[6] || ''),
      entrada: row[7] === "" ? null : row[7],
      salida: row[8] === "" ? null : row[8],
      saldo: row[9] || 0,
      usuario: String(row[10] || ''),
      observacion: String(row[11] || ''),
      area: String(row[12] || ''),
      comprobante: String(row[13] || ''),
      transcrito: String(row[14] || 'NO')
    };
  });
  
  return formattedData.reverse();
}

// ==========================================
// REGISTRO SEGURO DE MOVIMIENTOS (CON SEMÁFORO)
// ==========================================
function registrarMovimiento(datos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // Semáforo de 15 segundos para evitar choques
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Validar PIN
    const sheetUsers = ss.getSheetByName("Usuarios");
    if (!sheetUsers) throw new Error("La pestaña 'Usuarios' no existe.");
    const usersData = sheetUsers.getDataRange().getValues();
    let pinValido = false;
    for (let i = 1; i < usersData.length; i++) {
      if (String(usersData[i][0]).trim() === String(datos.usuario).trim()) {
        if (String(usersData[i][1]).trim() === String(datos.pin).trim()) {
          pinValido = true;
        }
        break;
      }
    }
    if (!pinValido) throw new Error("Contraseña incorrecta o usuario no encontrado.");

    // 2. Abrir hoja inventario
    const sheet = ss.getSheetByName("Inventario");
    if (!sheet) throw new Error("No se encontró la pestaña 'Inventario'");

    // 3. Buscar último saldo de ese código y lote específico
    const data = sheet.getDataRange().getValues();
    let ultimoSaldo = 0;
    for (let i = data.length - 1; i > 0; i--) {
      // Ignorar anulados al buscar saldo
      if (String(data[i][11] || "").includes("[ANULADO]")) continue;
      
      if (String(data[i][3]).trim() === String(datos.codigo).trim() &&
          String(data[i][6] || "").trim() === String(datos.lote || "").trim()) {
        ultimoSaldo = Number(data[i][9]) || 0;
        break;
      }
    }

    // 4. Calcular nuevo saldo y proteger contra stock negativo
    const cantidad = Number(datos.cantidad) || 0;
    let nuevoSaldo = ultimoSaldo;
    let entrada = "";
    let salida = "";

    if (datos.tipo === "Entrada") {
      entrada = cantidad;
      nuevoSaldo += cantidad;
    } else if (datos.tipo === "Salida") {
      salida = cantidad;
      if (cantidad > ultimoSaldo) {
        throw new Error("Stock insuficiente. Disponible: " + ultimoSaldo + ", solicitado: " + cantidad);
      }
      nuevoSaldo -= cantidad;
    }

    // 5. Generar Identificador Único Universal (UUID)
    const movId = Utilities.getUuid();

    // 6. Fecha y Hora exactas desde el servidor (no manipulables)
    const now = new Date();
    const fecha = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const hora = Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm:ss a");

    // 7. Insertar fila en Google Sheets
    sheet.appendRow([
      fecha, hora, datos.tipo, datos.codigo, datos.nombre, datos.tipo2,
      datos.lote || "", entrada, salida, nuevoSaldo, datos.usuario,
      datos.observacion || "", datos.area || "", datos.comprobante || "",
      "NO", movId
    ]);

    SpreadsheetApp.flush(); // Forzar guardado físico inmediato
    return { success: true, nuevoSaldo: nuevoSaldo, id: movId };

  } finally {
    lock.releaseLock(); // Apagar semáforo para el siguiente usuario
  }
}

// ==========================================
// MARCAR TRANSCRITOS FÍSICAMENTE
// ==========================================
function marcarTranscrito(fecha, hora, codigo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inventario");
  if (!sheet) throw new Error("No se encontró la pestaña 'Inventario'");
  
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    const sheetFecha = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(data[i][0]);
    const sheetHora = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "hh:mm a") : String(data[i][1]);
    
    if (sheetFecha === String(fecha) && 
        sheetHora === String(hora) && 
        String(data[i][3]) === String(codigo)) {
      sheet.getRange(i + 1, 15).setValue("SI");
      return { success: true };
    }
  }
  return { success: false, message: "Movimiento no encontrado" };
}

function marcarMultiplesTranscritos(lista) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inventario");
  if (!sheet) throw new Error("No se encontró la pestaña 'Inventario'");
  
  const data = sheet.getDataRange().getValues();
  let actualizados = 0;
  
  for (let idx = 0; idx < lista.length; idx++) {
    const item = lista[idx];
    for (let i = data.length - 1; i > 0; i--) {
      const sheetFecha = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(data[i][0]);
      const sheetHora = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], Session.getScriptTimeZone(), "hh:mm a") : String(data[i][1]);
      
      if (sheetFecha === String(item.fecha) && 
          sheetHora === String(item.hora) && 
          String(data[i][3]) === String(item.codigo)) {
        sheet.getRange(i + 1, 15).setValue("SI");
        actualizados++;
        break; 
      }
    }
  }
  return { success: true, count: actualizados };
}

// ==========================================
// ANULACIÓN DE MOVIMIENTOS SEGURA
// ==========================================
function anularMovimiento(datos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetUsers = ss.getSheetByName("Usuarios");
    if (!sheetUsers) throw new Error("La pestaña 'Usuarios' no existe.");
    const usersData = sheetUsers.getDataRange().getValues();
    let pinValido = false;
    for (let i = 1; i < usersData.length; i++) {
      if (String(usersData[i][0]).trim() === String(datos.usuario).trim()) {
        if (String(usersData[i][1]).trim() === String(datos.pin).trim()) {
          pinValido = true;
        }
        break;
      }
    }
    if (!pinValido) throw new Error("Contraseña incorrecta o usuario no encontrado.");

    const sheetInv = ss.getSheetByName("Inventario");
    if (!sheetInv) throw new Error("No se encontró la pestaña 'Inventario'");

    const invData = sheetInv.getDataRange().getValues();
    let rowOriginalIndex = -1;
    let rowOriginal = null;

    for (let i = invData.length - 1; i > 0; i--) {
      if (datos.movId) {
        if (String(invData[i][15] || "").trim() === String(datos.movId).trim()) {
          rowOriginalIndex = i + 1;
          rowOriginal = invData[i];
          break;
        }
      } else {
        const sheetFecha = invData[i][0] instanceof Date ? Utilities.formatDate(invData[i][0], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(invData[i][0]);
        const sheetHora = invData[i][1] instanceof Date ? Utilities.formatDate(invData[i][1], Session.getScriptTimeZone(), "hh:mm:ss a") : String(invData[i][1]);
        if (sheetFecha === String(datos.fecha_orig) &&
            sheetHora === String(datos.hora_orig) &&
            String(invData[i][3]).trim() === String(datos.codigo).trim()) {
          rowOriginalIndex = i + 1;
          rowOriginal = invData[i];
          break;
        }
      }
    }
    if (!rowOriginal) throw new Error("Movimiento original no encontrado para anular.");

    const obsOriginal = String(rowOriginal[11] || "");
    if (obsOriginal.includes("[ANULADO]")) throw new Error("Este movimiento ya fue anulado.");

    sheetInv.getRange(rowOriginalIndex, 12).setValue(obsOriginal + " [ANULADO]");

    let ultimoSaldo = 0;
    for (let i = invData.length - 1; i > 0; i--) {
      if (String(invData[i][3]).trim() === String(rowOriginal[3]).trim() &&
          String(invData[i][6] || "").trim() === String(rowOriginal[6] || "").trim() &&
          !String(invData[i][11] || "").includes("[ANULADO]")) {
        ultimoSaldo = Number(invData[i][9]) || 0;
        break;
      }
    }

    const originalTipo = rowOriginal[2];
    let entradaInversa = "";
    let salidaInversa = "";
    let nuevoSaldo = ultimoSaldo;

    if (originalTipo === "Entrada") {
      const qty = Number(rowOriginal[7]) || 0;
      salidaInversa = qty;
      nuevoSaldo -= qty;
    } else {
      const qty = Number(rowOriginal[8]) || 0;
      entradaInversa = qty;
      nuevoSaldo += qty;
    }

    const now = new Date();
    const fechaExt = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const horaExt = Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm:ss a");
    const extId = Utilities.getUuid();

    sheetInv.appendRow([
      fechaExt, horaExt, "Ajuste (Anulación)", rowOriginal[3], rowOriginal[4],
      rowOriginal[5], rowOriginal[6], entradaInversa, salidaInversa, nuevo/Saldo,
      datos.usuario,
      "AJUSTE: Anulación de movimiento del " + datos.fecha_orig + ". Motivo: " + datos.motivo,
      rowOriginal[12], rowOriginal[13], "SI", extId
    ]);

    let sheetAudit = ss.getSheetByName("Auditoria_Logs");
    if (!sheetAudit) {
      sheetAudit = ss.insertSheet("Auditoria_Logs");
      sheetAudit.appendRow(["FechaHora", "Usuario", "Acción", "Detalle", "MovID_Original"]);
      sheetAudit.getRange("A1:E1").setFontWeight("bold").setBackground("#f4cccc");
    }
    const fechaHoraFull = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    const detalle = "Anuló movimiento de " + originalTipo + " del producto " + rowOriginal[4] +
                    " (Cód: " + rowOriginal[3] + "). Motivo: " + datos.motivo;
    sheetAudit.appendRow([fechaHoraFull, datos.usuario, "ANULACIÓN", detalle, datos.movId || "legacy"]);

    SpreadsheetApp.flush();
    return { success: true };

  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// OBTENER USUARIOS DEL SISTEMA
// ==========================================
function getUsuarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Usuarios");
  
  if (!sheet) {
    sheet = ss.insertSheet("Usuarios");
    sheet.appendRow(["Usuario", "PIN", "Rol"]);
    sheet.appendRow(["Christian R.", "12345", "Administrador"]);
    sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#d9ead3");
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  return data.slice(1).map(row => row[0]).filter(name => name !== "");
}
