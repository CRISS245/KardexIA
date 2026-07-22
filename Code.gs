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
      transcrito: String(row[14] || 'NO'),
      movId: String(row[15] || '')
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
        // En Descarte el PIN se envía como 0000 por defecto. Si el PIN coincide, o si es un Descarte (bypass), es válido.
        if (String(usersData[i][1]).trim() === String(datos.pin).trim() || datos.tipo === 'Descarte') {
            pinValido = true;
        } else {
            throw new Error("La contraseña está mal escrita.");
        }
        break;
      }
    }
    if (!pinValido) throw new Error("Usuario no encontrado.");

    // 2. Abrir hoja inventario
    const sheet = ss.getSheetByName("Inventario");
    if (!sheet) throw new Error("No se encontró la pestaña 'Inventario'");

    // 3. Buscar saldo (Global y por Lote) sumando entradas y restando salidas
    const data = sheet.getDataRange().getValues();
    let ultimoSaldoLote = 0;
    let ultimoSaldoGlobal = 0;
    
    // Normalizar: quitar caracteres especiales, espacios, mayúsculas
    const normalize = (str) => String(str || "").toLowerCase().replace(/[^a-z0-9]/g, '');
    // Para el lote usamos SOLO la primera parte (antes del |) para no fallar por diferencias en Marca/Vence
    const extractLoteNum = (str) => normalize(String(str || "").split("|")[0]);
    
    const targetCodigo = normalize(datos.codigo);
    const targetLote = extractLoteNum(datos.lote);

    for (let i = 1; i < data.length; i++) {
      // Ignorar anulados
      if (String(data[i][11] || "").includes("[ANULADO]")) continue;
      
      const rowCodigo = normalize(data[i][3]);
      const rowLote = extractLoteNum(data[i][6]);
      
      if (rowCodigo === targetCodigo) {
         const rowTipo = String(data[i][2]).trim();
         const rowEntrada = Number(data[i][7]) || 0;
         const rowSalida = Number(data[i][8]) || 0;
         
         // Suma Global (todos los lotes)
         if (rowTipo === 'Entrada' || rowTipo === 'Prestamo') ultimoSaldoGlobal += rowEntrada;
         if (rowTipo === 'Salida' || rowTipo === 'Descarte' || rowTipo === 'Devolucion') ultimoSaldoGlobal -= rowSalida;
         if (rowTipo === 'Ajuste') {
            ultimoSaldoGlobal += rowEntrada;
            ultimoSaldoGlobal -= rowSalida;
         }

         // Suma Específica del Lote afectado (para el registro)
         if (rowLote === targetLote) {
             if (rowTipo === 'Entrada' || rowTipo === 'Prestamo') ultimoSaldoLote += rowEntrada;
             if (rowTipo === 'Salida' || rowTipo === 'Descarte' || rowTipo === 'Devolucion') ultimoSaldoLote -= rowSalida;
             if (rowTipo === 'Ajuste') {
                ultimoSaldoLote += rowEntrada;
                ultimoSaldoLote -= rowSalida;
             }
         }
      }
    }

    // 4. Calcular nuevo saldo (del Lote y Global) y proteger contra stock negativo
    const cantidad = Number(datos.cantidad) || 0;
    let nuevoSaldoLote = Number(ultimoSaldoLote);
    let nuevoSaldoGlobal = Number(ultimoSaldoGlobal);
    let entrada = "";
    let salida = "";

    if (datos.tipo === "Entrada") {
      entrada = cantidad;
      nuevoSaldoLote += cantidad;
      nuevoSaldoGlobal += cantidad;
    } else if (datos.tipo === "Salida" || datos.tipo === "Descarte") {
      salida = cantidad;
      if (datos.tipo === "Salida" && cantidad > ultimoSaldoLote) {
        throw new Error("Stock insuficiente en este lote. Disponible: " + ultimoSaldoLote + ", solicitado: " + cantidad);
      }
      nuevoSaldoLote -= cantidad;
      nuevoSaldoGlobal -= cantidad;
    } else if (datos.tipo === "Ajuste") {
      if (cantidad > ultimoSaldoLote) {
        entrada = cantidad - ultimoSaldoLote;
        nuevoSaldoLote = cantidad;
        nuevoSaldoGlobal += entrada;
      } else if (cantidad < ultimoSaldoLote) {
        salida = ultimoSaldoLote - cantidad;
        nuevoSaldoLote = cantidad;
        nuevoSaldoGlobal -= salida;
      } else {
        throw new Error("La cantidad ingresada es igual al stock actual de este lote (" + ultimoSaldoLote + "). No hay diferencia para ajustar.");
      }
    } else if (datos.tipo === "Prestamo") {
      entrada = cantidad;
      nuevoSaldoLote += cantidad;
      nuevoSaldoGlobal += cantidad;
    } else if (datos.tipo === "Devolucion") {
      salida = cantidad;
      nuevoSaldoLote -= cantidad;
      nuevoSaldoGlobal -= cantidad;
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
      datos.lote || "", entrada, salida, nuevoSaldoLote, datos.usuario,
      datos.observacion || "", datos.area || "", datos.comprobante || "",
      "NO", movId
    ]);

    // 8. Alerta en tiempo real por stock crítico GLOBAL (solo si baja a 5 o menos en total)
    if (ultimoSaldoGlobal > 5 && nuevoSaldoGlobal <= 5) {
      try {
        enviarAlertaCriticaRealTime(datos.nombre, "TODOS LOS LOTES", nuevoSaldoGlobal);
      } catch(e) {
        // Ignorar errores de envío para no trabar el sistema
      }
    }

    SpreadsheetApp.flush(); // Forzar guardado físico inmediato
    return { success: true, nuevoSaldo: nuevoSaldoLote, id: movId };

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
        const sheetHora = invData[i][1] instanceof Date ? Utilities.formatDate(invData[i][1], Session.getScriptTimeZone(), "hh:mm a") : String(invData[i][1]);
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

    // Calcular saldo real sumando entradas y restando salidas (igual que registrarMovimiento)
    const normL = (str) => String(str || "").split("|")[0].toLowerCase().replace(/[^a-z0-9]/g, '');
    const targetCodAnul = normL(rowOriginal[3]);
    const targetLoteAnul = normL(rowOriginal[6]);
    let ultimoSaldo = 0;
    for (let i = 1; i < invData.length; i++) {
      if (String(invData[i][11] || "").includes("[ANULADO]")) continue;
      if (normL(invData[i][3]) === targetCodAnul && normL(invData[i][6]) === targetLoteAnul) {
        const tipo = String(invData[i][2]).trim();
        if (tipo === 'Entrada') ultimoSaldo += Number(invData[i][7]) || 0;
        if (tipo === 'Salida' || tipo === 'Descarte') ultimoSaldo -= Number(invData[i][8]) || 0;
        if (tipo === 'Ajuste' || tipo === 'Ajuste (Anulación)') {
          ultimoSaldo += Number(invData[i][7]) || 0;
          ultimoSaldo -= Number(invData[i][8]) || 0;
        }
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
      rowOriginal[5], rowOriginal[6], entradaInversa, salidaInversa, nuevoSaldo,
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
    sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#d9ead3");
  }
  
  const data = sheet.getDataRange().getValues();
  const existingUsers = data.length > 1 ? data.slice(1).map(row => String(row[0]).trim()) : [];
  
  const requiredUsers = [
    "Christian R.", "Maria S.", "Violeta G.", "Junior E.", "Juan G.", 
    "Eliana S.", "Evelyn L.", "Clau A.", "Susan D.", "Yesenia R.", "Dr. Aarón"
  ];
  
  let added = false;
  requiredUsers.forEach(u => {
    if (!existingUsers.includes(u) && !existingUsers.includes("María S.") && !(u === "Maria S." && existingUsers.includes("María S."))) {
      let rol = "Técnico de Laboratorio";
      if (u === "Christian R.") rol = "Administrador";
      if (u === "Maria S." || u === "María S." || u === "Junior E." || u === "Dr. Aarón") rol = "Personal de Almacén";
      
      const pin = (u === "Christian R.") ? "12345" : "0000";
      sheet.appendRow([u, pin, rol]);
      added = true;
    }
  });

  // Forzar actualización de roles en la hoja por si ya estaban creados
  let changed = false;
  const newData = sheet.getDataRange().getValues();
  for (let i = 1; i < newData.length; i++) {
     let u = String(newData[i][0]).trim();
     let currentRol = String(newData[i][2]).trim();
     let desiredRol = currentRol;
     
     if (u === "Maria S." || u === "María S." || u === "Junior E." || u === "Dr. Aarón" || u === "Dr. Aaron") {
         desiredRol = "Personal de Almacén";
     } else if (u === "Christian R.") {
         desiredRol = "Administrador";
     } else {
         desiredRol = "Técnico de Laboratorio";
     }
     
     if (currentRol !== desiredRol) {
         sheet.getRange(i + 1, 3).setValue(desiredRol);
         changed = true;
     }
  }
  
  if (added || changed) {
    const finalData = sheet.getDataRange().getValues();
    return finalData.slice(1).map(row => ({ nombre: String(row[0]).trim(), rol: String(row[2]).trim() })).filter(u => u.nombre !== "");
  }
  
  if (data.length <= 1) return [];
  return data.slice(1).map(row => ({ nombre: String(row[0]).trim(), rol: String(row[2]).trim() })).filter(u => u.nombre !== "");
}

// ==========================================
// POBLAR DATOS DE SIMULACIÓN REAL (ESCENARIO COMPLETO)
// ==========================================
function poblarDatosDePrueba() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inventario");
  const logSheet = ss.getSheetByName("Auditoria_Logs");
  
  if (!sheet) return;
  
  // 1. Limpiar las hojas de datos anteriores
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  if (logSheet) {
    const logLastRow = logSheet.getLastRow();
    if (logLastRow > 1) {
      logSheet.getRange(2, 1, logLastRow - 1, logSheet.getLastColumn()).clearContent();
    }
  }
  
  // 2. Datos de prueba exactos (Catálogo completo de las 3 imágenes)
  const productosBase = [
    // --- MATERIAL DE LABORATORIO (1 - 40) ---
    { nombre: "AGUJAS DESCARTABLES 20 X 1 [CAJA X (100 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 01 | Vence: 06/03/2028" },
    { nombre: "AGUJAS VACUTAINER 21 X 1/2 [CAJA X (100 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 02 | Vence: 01/01/2028" },
    { nombre: "ALITAS VACUTAINER SAFETY-LOK 23 X 3/4 [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 03 | Vence: 30/04/2026" },
    { nombre: "GASA FRACCIONADA ESTERIL 5X5 cm [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 04 | Vence: 01/12/2025" },
    { nombre: "ALGODÓN x 500 GR [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 05 | Vence: 01/06/2028" },
    { nombre: "ALGODÓN x 100 GR [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 06 | Vence: 01/08/2026" },
    { nombre: "ESPARADRAPO HOSPITALARIO 4 CORTES [CAJA X (4 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 07 | Vence: 11/12/2025" },
    { nombre: "GUANTES DESCARTABLES S CAJA X 100 UNIDADES [CAJA X (100 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 08 | Vence: 12/07/2028" },
    { nombre: "GUANTES DESCARTABLES L CAJA X 100 UNIDADES [CAJA X (100 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 09 | Vence: 01/02/2027" },
    { nombre: "GUANTES DE NITRILO DESCART. S CAJA X 100 UNIDADES [CAJA X (100 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 10 | Vence: 11/07/2026" },
    { nombre: "GUANTES DE NITRILO DESCART. M CAJA X 100 UNIDADES [CAJA X (100 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 11 | Vence: 19/06/2028" },
    { nombre: "JERINGA DESC. 10 ML C/AGUJA DE 21x1 1/2 [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 12 | Vence: 19/07/2028" },
    { nombre: "MASCARILLAS DESCARTABLES x 50 und [CAJA X (50 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 13 | Vence: 30/07/2028" },
    { nombre: "MANDILON DESCARTABLE T/M [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 14 | Vence: 01/02/2028" },
    { nombre: "MAMELUCO DESCARTABLE T/M [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 15 | Vence: 10/01/2027" },
    { nombre: "MAMELUCO DESCARTABLE T/L [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 16 | Vence: 10/01/2027" },
    { nombre: "KIT CHAQUETA Y PANTALON T/L [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 17 | Vence: 01/04/2025" },
    { nombre: "KIT CHAQUETA Y PANTALON T/M [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 18 | Vence: 01/04/2026" },
    { nombre: "LANCETAS 1.8 DESCARTABLES (200 UNIDADES) [CAJA X (200 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 19 | Vence: 30/04/2025" },
    { nombre: "PROTECTORES OCULARES [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 20 | Vence: S/N" },
    { nombre: "GORRO DESCARTABLE DE ENFERMERA [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 21 | Vence: 26/12/2026" },
    { nombre: "BOTAS PROTECTORAS DE CALZADO [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 22 | Vence: 26/12/2026" },
    { nombre: "ALCOHOL PURO 70° [LITROS]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 23 | Vence: 30/04/2026" },
    { nombre: "LAPIZ DE CERA PARA VIDRIO ROJO CAJA X 10 UNIDADES [CAJA X (10 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 24 | Vence: S/N" },
    { nombre: "LAPIZ DE CERA PARA VIDRIO AZUL CAJA X 10 UNIDADES [CAJA X (10 UNIDADES)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 25 | Vence: S/N" },
    { nombre: "TUBO CON EDTA X 0.5 ML. [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 26 | Vence: 31/10/2026" },
    { nombre: "PAPEL DE FILTRO CORRIENTE [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 27 | Vence: S/N" },
    { nombre: "CAPILAR PARA HEMATOCRITO CON HEPARINA X 100 [CAJA X (10 FRASCOS)]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 28 | Vence: 01/07/2026" },
    { nombre: "TUBOS DE PLASTICO CON EDTA X 6ML (INFECTOLOGIA) [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 29 | Vence: 06/06/2026" },
    { nombre: "TUBOS DE PLASTICO CON EDTA X 3ML [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 30 | Vence: 01/05/2026" },
    { nombre: "TUBOS DE PLASTICO CON EDTA X 2ML [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 31 | Vence: 30/11/2026" },
    { nombre: "TUBOS DE PLASTICO CON GEL SEPARADOR X 3.5 ML [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 32 | Vence: 30/11/2026" },
    { nombre: "TUBO DE PLASTICO CON CITRATO [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 33 | Vence: 29/02/2026" },
    { nombre: "TUBO DE PLASTICO CON CITRATO (HEMATOLOGIA 249) [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 34 | Vence: 01/03/2026" },
    { nombre: "FRASCO BACTEC PARA HEMOCULTIVOS [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 35 | Vence: 20/05/2026" },
    { nombre: "LAMINAS PORTA OBJETO DE VIDRIO (HEMATOLOGIA) [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 36 | Vence: S/N" },
    { nombre: "LAMINAS PORTA OBJETO DE VIDRIO CITOPLUS (DEIDAD) [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 37 | Vence: S/N" },
    { nombre: "ETIQUETERAS PARA CODIGO DE BARRA (HEMATOLOGIA) [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 38 | Vence: S/N" },
    { nombre: "LAMINAS PARA TEST DE GRAHAM (MICROBIOLOGIA) [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 39 | Vence: S/N" },
    { nombre: "CAJAS PARA TRANSPORTE DE MUESTRAS [UNIDAD]", cat: "MATERIAL DE LABORATORIO", loteInicial: "Lote 40 | Vence: S/N" },

    // --- MATERIAL DE ESCRITORIO (41 - 55) ---
    { nombre: "GRAPAS [CAJA X (5000 UNIDADES)]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 41 | Vence: S/N" },
    { nombre: "ROLLO PARA PEDESTAL DE RECEPCION SANGRE (DB) [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 42 | Vence: S/N" },
    { nombre: "ROLLO PARA PEDESTAL DE RECEPCION ORINA (BD) [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 43 | Vence: S/N" },
    { nombre: "CORRECTOR COLOR BLANCO [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 44 | Vence: S/N" },
    { nombre: "PAPEL BOND 80 gr TAMAÑO A4 [PAQUETE X (500 UNIDADES)]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 45 | Vence: S/N" },
    { nombre: "CINTA ADHESIVA DELGADA [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 46 | Vence: S/N" },
    { nombre: "TINTA AZUL/ROJO PARA TAMPON [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 47 | Vence: S/N" },
    { nombre: "TAMPOM PARA SELLO ROJO [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 48 | Vence: S/N" },
    { nombre: "CUADERNO DE CARGO [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 49 | Vence: S/N" },
    { nombre: "FASTENERS [CAJA X (50 UNIDADES)]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 50 | Vence: S/N" },
    { nombre: "GOMA EN BARRA [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 51 | Vence: S/N" },
    { nombre: "ARCHIVADOR PLASTIFICADO [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 52 | Vence: S/N" },
    { nombre: "FORMATO DE REGISTRO DE MATERIALES KARDEX [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 53 | Vence: S/N" },
    { nombre: "ENGRAPADOR DE METAL [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 54 | Vence: S/N" },
    { nombre: "BOLIGRAFO TINTA FINA COLOR AZUL [UNIDAD]", cat: "MATERIAL DE ESCRITORIO", loteInicial: "Lote 55 | Vence: S/N" },

    // --- MATERIAL DE LIMPIEZA (56 - 60) ---
    { nombre: "BOLSA DE PLASTICO NEGRA x 100 [PAQUETE X (100 UNIDADES)]", cat: "MATERIAL DE LIMPIEZA", loteInicial: "Lote 56 | Vence: S/N" },
    { nombre: "BOLSA DE PLASTICO ROJA x 100 [PAQUETE X (100 UNIDADES)]", cat: "MATERIAL DE LIMPIEZA", loteInicial: "Lote 57 | Vence: S/N" },
    { nombre: "GLUCONATO DE CLORHEXIDINA 2% EN ESPUMA X 1 Lt. [UNIDAD]", cat: "MATERIAL DE LIMPIEZA", loteInicial: "Lote 58 | Vence: 01/10/2026" },
    { nombre: "PAPEL TOALLA DOBLADO PQT. x 200 HOJAS [UNIDAD]", cat: "MATERIAL DE LIMPIEZA", loteInicial: "Lote 59 | Vence: S/N" },
    { nombre: "LEJIA X 500 Ml [UNIDAD]", cat: "MATERIAL DE LIMPIEZA", loteInicial: "Lote 60 | Vence: 02/05/2026" }
  ];

  let eventos = [];
  let logs = [];
  let saldosGlobales = {};
  
  const usuariosTomaMuestra = ["Dr. Aarón", "Junior E.", "María S."];
  const areaUnica = "TOMA DE MUESTRA";
  
  // Asignar códigos automáticos y saldos iniciales 0
  productosBase.forEach((p, idx) => {
    p.codigo = "INSN-" + ("00" + (idx + 1)).slice(-3);
    saldosGlobales[p.codigo] = 0;
  });
  
  // -- FASE 1: ENERO 2026 (INVENTARIO INICIAL) --
  productosBase.forEach(p => {
    saldosGlobales[p.codigo] += 200; // Todos empiezan con 200 unidades
    eventos.push([
      "05/01/2026", "08:00 AM", "Entrada", p.codigo, p.nombre, p.cat, p.loteInicial, 200, "", saldosGlobales[p.codigo], "Administrador", "Apertura de inventario anual", "ALMACÉN CENTRAL", "PECOSA-001", "SI"
    ]);
  });
  
  // -- FASE 2: FEBRERO Y MARZO 2026 (CONSUMO RUTINARIO) --
  // Simulamos consumos en la mitad de todos los productos (los primeros 30)
  productosBase.slice(0, 30).forEach((p, idx) => {
    let usu1 = usuariosTomaMuestra[idx % 3];
    let usu2 = usuariosTomaMuestra[(idx + 1) % 3];
    
    saldosGlobales[p.codigo] -= 25;
    eventos.push([
      "14/02/2026", "10:30 AM", "Salida", p.codigo, p.nombre, p.cat, p.loteInicial, "", 25, saldosGlobales[p.codigo], usu1, "Consumo regular", areaUnica, "VAL-101", "NO"
    ]);
    
    saldosGlobales[p.codigo] -= 30;
    eventos.push([
      "22/03/2026", "09:15 AM", "Salida", p.codigo, p.nombre, p.cat, p.loteInicial, "", 30, saldosGlobales[p.codigo], usu2, "Material para toma de muestras", areaUnica, "VAL-102", "NO"
    ]);
  });
  
  // -- FASE 3: ABRIL 2026 (NUEVOS INGRESOS CON NUEVOS LOTES) --
  // Ingresa nuevo stock para Gasa, Algodón, y Tubos EDTA
  let prodGasa = productosBase.find(p => p.nombre.includes("GASA"));
  saldosGlobales[prodGasa.codigo] += 500;
  eventos.push(["10/04/2026", "11:00 AM", "Entrada", prodGasa.codigo, prodGasa.nombre, prodGasa.cat, "Lote NUEVO | Vence: 15/10/2029", 500, "", saldosGlobales[prodGasa.codigo], "Administrador", "Compra mensual", "PROVEEDOR EXT", "FAC-992", "NO"]);
  
  let prodTubo = productosBase.find(p => p.nombre.includes("EDTA"));
  saldosGlobales[prodTubo.codigo] += 1000;
  eventos.push(["15/04/2026", "15:20 PM", "Entrada", prodTubo.codigo, prodTubo.nombre, prodTubo.cat, "Lote EDTA-B | Vence: 20/05/2030", 1000, "", saldosGlobales[prodTubo.codigo], "Administrador", "Compra trimestral", "PROVEEDOR EXT", "FAC-993", "NO"]);

  // -- FASE 4: MAYO 2026 (SIMULACIÓN DE ERROR Y ANULACIÓN) --
  let prodGuantes = productosBase.find(p => p.nombre.includes("GUANTES DESCARTABLES S"));
  let usuError = usuariosTomaMuestra[2]; // María S.
  
  // 1. Sale por error una cantidad inmensa (5000 en vez de 50)
  eventos.push(["10/05/2026", "08:15 AM", "Salida", prodGuantes.codigo, prodGuantes.nombre, prodGuantes.cat, prodGuantes.loteInicial, "", 5000, saldosGlobales[prodGuantes.codigo] - 5000, usuError, "[ANULADO] Error de tipeo", areaUnica, "VAL-150", "NO"]);
  logs.push(["10/05/2026 08:30 AM", "Administrador", "ANULACIÓN", `Anuló movimiento de Salida del producto ${prodGuantes.nombre}. Motivo: Se tecleó 5000 por error, eran 50.`]);
  
  // 2. Se anula devolviendo los 5000 al kardex
  eventos.push(["10/05/2026", "08:30 AM", "Ajuste (Anulación)", prodGuantes.codigo, prodGuantes.nombre, prodGuantes.cat, prodGuantes.loteInicial, 5000, "", saldosGlobales[prodGuantes.codigo], "Administrador", "Ajuste: Anulación por error de tipeo", "SISTEMA", "VAL-150", "SI"]);
  
  // 3. Se ingresa la salida correcta
  saldosGlobales[prodGuantes.codigo] -= 50;
  eventos.push(["10/05/2026", "08:35 AM", "Salida", prodGuantes.codigo, prodGuantes.nombre, prodGuantes.cat, prodGuantes.loteInicial, "", 50, saldosGlobales[prodGuantes.codigo], usuError, "Consumo correcto", areaUnica, "VAL-151", "NO"]);

  // -- FASE 5: JUNIO/JULIO 2026 (CONSUMOS DIVERSOS RECIENTES) --
  productosBase.forEach((p, idx) => {
    if (idx % 3 === 0) { // Consumo de algunos productos al azar
      let cant = 15;
      let usuRand = usuariosTomaMuestra[idx % 3];
      saldosGlobales[p.codigo] -= cant;
      eventos.push([
        "20/06/2026", "14:10 PM", "Salida", p.codigo, p.nombre, p.cat, p.loteInicial, "", cant, saldosGlobales[p.codigo], usuRand, "Uso general", areaUnica, "VAL-200", "NO"
      ]);
    }
  });

  // Guardar en la hoja principal
  if (eventos.length > 0) {
    sheet.getRange(2, 1, eventos.length, eventos[0].length).setValues(eventos);
  }
  
  // Guardar logs
  if (logSheet && logs.length > 0) {
    logSheet.getRange(2, 1, logs.length, logs[0].length).setValues(logs);
  }
}

// ==========================================
// ALERTAS AUTOMÁTICAS POR CORREO
// ==========================================
function revisarAlertasYEnviarCorreo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inventario");
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; // Sin datos
  
  // 1. Agrupar saldos por lote y extraer fechas de vencimiento
  const lotes = {}; 
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[11] || "").includes("[ANULADO]")) continue;
    
    const codigo = String(row[3]);
    const nombre = String(row[4]);
    const cat = String(row[5]);
    const loteRaw = String(row[6]);
    if (!loteRaw) continue;
    
    // Extraer lote corto y fecha
    let loteCorto = loteRaw;
    let fechaVenceStr = null;
    const partes = loteRaw.split('|');
    loteCorto = partes[0].trim();
    
    for (let p of partes) {
      if (p.includes("Vence:")) {
        fechaVenceStr = p.replace("Vence:", "").trim(); // DD/MM/YY
      }
    }
    
    if (!lotes[loteRaw]) {
      lotes[loteRaw] = { codigo: codigo, nombre: nombre, cat: cat, loteCorto: loteCorto, stock: 0, fechaVenceStr: fechaVenceStr };
    }
    
    const entrada = Number(row[7]) || 0;
    const salida = Number(row[8]) || 0;
    
    if (row[2] === 'Entrada' || row[2] === 'Ajuste' || row[2] === 'Prestamo') lotes[loteRaw].stock += entrada;
    if (row[2] === 'Salida' || row[2] === 'Ajuste' || row[2] === 'Devolucion' || row[2] === 'Descarte') lotes[loteRaw].stock -= salida;
  }
  
  // 2. Revisar umbrales y fechas
  const alertasVencimiento = [];
  const alertasStockBajo = [];
  
  const today = new Date();
  today.setHours(0,0,0,0);
  
  for (const l in lotes) {
    const obj = lotes[l];
    if (obj.stock > 0) { // Solo evaluar si hay stock físico
      
      // Revisar Stock Crítico (5 o menos unidades, ajústalo según necesites)
      if (obj.stock <= 5) { 
        alertasStockBajo.push(obj);
      }
      
      // Revisar Vencimiento
      if (obj.fechaVenceStr) {
        const parts = obj.fechaVenceStr.split('/');
        if (parts.length === 3) {
           let yy = parts[2].length === 2 ? '20' + parts[2] : parts[2];
           const expDate = new Date(parseInt(yy), parseInt(parts[1]) - 1, parseInt(parts[0]));
           const diffDays = Math.ceil((expDate - today) / (1000 * 60 * 60 * 24));
           
           if (diffDays <= 90) { // Menos de 3 meses para vencer
             obj.diasRestantes = diffDays;
             alertasVencimiento.push(obj);
           }
        }
      }
    }
  }
  
  // 3. Preparar correo si hay alertas
  if (alertasVencimiento.length === 0 && alertasStockBajo.length === 0) {
    return; // Nada que notificar hoy
  }
  
  let htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 800px; margin: 0 auto; border: 1px solid #ddd; padding: 20px; border-radius: 8px;">
      <h2 style="color: #be123c; margin-top:0; border-bottom: 2px solid #be123c; padding-bottom:10px;">⚠️ Alertas del Almacén (KardexIA)</h2>
      <p>Hola, el sistema ha detectado alertas críticas en tu inventario hoy <b>${today.toLocaleDateString('es-PE')}</b>.</p>
  `;
  
  if (alertasVencimiento.length > 0) {
    alertasVencimiento.sort((a,b) => a.diasRestantes - b.diasRestantes);
    htmlBody += `
      <h3 style="background-color: #fef2f2; border-left: 4px solid #ef4444; padding: 10px; color: #b91c1c; margin-top: 25px;">⏳ Productos por vencer (Próx. 90 días)</h3>
      <table style="border-collapse: collapse; width: 100%; font-size: 13px; margin-bottom: 20px;">
        <tr style="background-color: #f3f4f6; text-align: left;">
          <th style="padding: 10px; border: 1px solid #ddd;">Producto</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Lote</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Vence</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Días Rest.</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Stock Físico</th>
        </tr>
    `;
    alertasVencimiento.forEach(a => {
      let colorDias = a.diasRestantes <= 30 ? "color: #b91c1c; font-weight: bold; background: #fee2e2;" : "";
      htmlBody += `
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd; font-weight:bold;">${a.nombre}</td>
          <td style="padding: 8px; border: 1px solid #ddd;">${a.loteCorto}</td>
          <td style="padding: 8px; border: 1px solid #ddd; text-align:center;">${a.fechaVenceStr}</td>
          <td style="padding: 8px; border: 1px solid #ddd; text-align:center; ${colorDias}">${a.diasRestantes} días</td>
          <td style="padding: 8px; border: 1px solid #ddd; text-align:center;">${a.stock}</td>
        </tr>
      `;
    });
    htmlBody += `</table>`;
  }
  
  if (alertasStockBajo.length > 0) {
    htmlBody += `
      <h3 style="background-color: #fffbeb; border-left: 4px solid #f59e0b; padding: 10px; color: #b45309; margin-top: 25px;">📦 Stock Crítico (5 unidades o menos)</h3>
      <table style="border-collapse: collapse; width: 100%; font-size: 13px;">
        <tr style="background-color: #f3f4f6; text-align: left;">
          <th style="padding: 10px; border: 1px solid #ddd;">Producto</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Lote Activo</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Stock Actual</th>
        </tr>
    `;
    alertasStockBajo.forEach(a => {
      htmlBody += `
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd; font-weight:bold;">${a.nombre}</td>
          <td style="padding: 8px; border: 1px solid #ddd;">${a.loteCorto}</td>
          <td style="padding: 8px; border: 1px solid #ddd; color: #b91c1c; font-weight: bold; text-align:center; background:#fee2e2;">${a.stock}</td>
        </tr>
      `;
    });
    htmlBody += `</table>`;
  }
  
  htmlBody += `
      <br>
      <hr style="border:0; border-top:1px solid #eee; margin-top:20px;">
      <p style="font-size: 11px; color: #888; text-align:center;">
        Generado automáticamente por tu sistema de Inventarios <b>KardexIA</b>.<br>
        <i>No respondas a este correo.</i>
      </p>
    </div>
  `;
  
  // 4. Enviar Correo a todos los usuarios con correo registrado
  const sheetUsers = ss.getSheetByName("Usuarios");
  let destinatarios = "";
  if (sheetUsers) {
    const usersData = sheetUsers.getDataRange().getValues();
    const correos = [];
    for (let i = 1; i < usersData.length; i++) {
      const email = String(usersData[i][2] || "").trim();
      if (email.includes("@")) correos.push(email);
    }
    destinatarios = correos.join(",");
  }
  
  if (!destinatarios) return; // Nadie a quien notificar
  
  MailApp.sendEmail({
    to: destinatarios,
    subject: "🚨 ALERTA DIARIA: Vencimientos y Stock Crítico - KardexIA",
    htmlBody: htmlBody
  });
}

// ==========================================
// MÓDULO DE CONFIGURACIÓN Y USUARIOS
// ==========================================

function getUsuariosConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Usuarios");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const usuarios = [];
  for (let i = 1; i < data.length; i++) {
    const nombre = String(data[i][0] || "").trim();
    if (nombre !== "") {
      usuarios.push({
        nombre: nombre,
        pin: String(data[i][1] || "").trim(),
        correo: String(data[i][2] || "").trim(),
        rol: String(data[i][3] || "Personal de Almacén").trim()
      });
    }
  }
  return usuarios;
}

function agregarUsuarioConfig(nombre, pin, correo, rol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Usuarios");
  if (!sheet) {
    sheet = ss.insertSheet("Usuarios");
    sheet.appendRow(["Nombre", "PIN", "Correo", "Rol"]);
  }
  
  const data = sheet.getDataRange().getValues();
  let firstEmptyRow = 0;
  for (let i = 1; i < data.length; i++) {
    const rowNombre = String(data[i][0] || "").trim();
    if (rowNombre.toUpperCase() === String(nombre).trim().toUpperCase()) {
      throw new Error("El usuario ya existe.");
    }
    if (rowNombre === "" && firstEmptyRow === 0) {
      firstEmptyRow = i + 1; // 1-indexed row number
    }
  }
  
  if (firstEmptyRow > 0) {
    sheet.getRange(firstEmptyRow, 1, 1, 4).setValues([[nombre, pin, correo, rol]]);
  } else {
    sheet.appendRow([nombre, pin, correo, rol]);
  }
  SpreadsheetApp.flush();
  return { success: true };
}

function editarUsuarioConfig(nombreAntiguo, nombreNuevo, pinNuevo, correoNuevo, rolNuevo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Usuarios");
  if (!sheet) throw new Error("No existe la hoja de usuarios.");
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === String(nombreAntiguo).trim().toUpperCase()) {
      sheet.getRange(i + 1, 1, 1, 4).setValues([[nombreNuevo, pinNuevo, correoNuevo, rolNuevo]]);
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  throw new Error("Usuario no encontrado para editar.");
}

function eliminarUsuarioConfig(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Usuarios");
  if (!sheet) throw new Error("No existe la hoja de usuarios.");
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === String(nombre).trim().toUpperCase()) {
      sheet.getRange(i + 1, 1, 1, 4).clearContent();
      SpreadsheetApp.flush();
      return { success: true };
    }
  }
  throw new Error("Usuario no encontrado para eliminar.");
}

// ==========================================
// ALERTA EN TIEMPO REAL
// ==========================================
function enviarAlertaCriticaRealTime(nombre, lote, saldoActual) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUsers = ss.getSheetByName("Usuarios");
  let destinatarios = "";
  if (sheetUsers) {
    const usersData = sheetUsers.getDataRange().getValues();
    const correos = [];
    for (let i = 1; i < usersData.length; i++) {
      const email = String(usersData[i][2] || "").trim();
      if (email.includes("@")) correos.push(email);
    }
    destinatarios = correos.join(",");
  }
  
  if (!destinatarios) return; 
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #ddd; padding: 20px; border-radius: 8px;">
      <h2 style="color: #f59e0b; margin-top:0; border-bottom: 2px solid #f59e0b; padding-bottom:10px;">⚡ Alerta de Stock en Tiempo Real</h2>
      <p>Hola, se acaba de registrar un movimiento en el sistema que ha dejado un producto en <b>estado crítico</b>.</p>
      
      <table style="border-collapse: collapse; width: 100%; font-size: 14px; margin-top: 15px;">
        <tr style="background-color: #f3f4f6; text-align: left;">
          <th style="padding: 10px; border: 1px solid #ddd;">Producto</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Lote</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Stock Físico Restante</th>
        </tr>
        <tr>
          <td style="padding: 10px; border: 1px solid #ddd; font-weight:bold;">${nombre}</td>
          <td style="padding: 10px; border: 1px solid #ddd;">${lote}</td>
          <td style="padding: 10px; border: 1px solid #ddd; color: #b91c1c; font-weight: bold; text-align:center; background:#fee2e2; font-size:16px;">${saldoActual}</td>
        </tr>
      </table>
      
      <p style="font-size: 11px; color: #888; text-align:center; margin-top: 20px;">
        Esta es una notificación automática generada al instante por <b>KardexIA</b>.
      </p>
    </div>
  `;
  
  MailApp.sendEmail({
    to: destinatarios,
    subject: "⚡ URGENTE: Stock Agotándose (" + nombre + ")",
    htmlBody: htmlBody
  });
}

// ==========================================
// LIMPIEZA DE CATÁLOGO (ELIMINAR DUPLICADOS Y RESETEAR A 0)
// ==========================================
function limpiarCatalogoInicial() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Inventario");
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  
  const productosVistos = {};
  const filasConservadas = [];
  
  // Guardar cabecera (Fila 0)
  filasConservadas.push(data[0]);
  
  // Recorrer de arriba a abajo
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const codigo = String(row[3] || "").trim(); // Columna D (Código)
    
    // Si la fila tiene código y NO lo hemos visto antes, lo guardamos
    if (codigo !== "" && !productosVistos[codigo]) {
      productosVistos[codigo] = true;
      
      // Vaciamos Entrada (H), Salida (I) y Saldo (J)
      row[7] = ""; // Entrada
      row[8] = ""; // Salida
      row[9] = 0;  // Saldo
      
      filasConservadas.push(row);
    }
  }
  
  // Borrar toda la hoja de Excel
  sheet.clearContents();
  
  // Escribir de vuelta solo los productos únicos con stock en cero
  sheet.getRange(1, 1, filasConservadas.length, filasConservadas[0].length).setValues(filasConservadas);
  
  SpreadsheetApp.flush();
}

// ==========================================
// MÓDULO PECOSA: CONFIGURAR PESTAÑA HISTÓRICO
// ==========================================
function setupPestanaHistorico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Histórico Consumo");
  if (!sheet) {
    sheet = ss.insertSheet("Histórico Consumo");
    sheet.appendRow(["Código", "Producto", "Consumo Mes 1", "Consumo Mes 2", "Consumo Mes 3", "Consumo Mes 4"]);
    sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#3b82f6").setFontColor("#ffffff");
    sheet.setColumnWidth(2, 300);
  }
  return "Pestaña 'Histórico Consumo' creada o ya existe.";
}

// ==========================================
// MÓDULO PECOSA: OBTENER DATOS HISTÓRICOS
// ==========================================
function getHistoricoData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Histórico Consumo");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0].map(h => String(h).toLowerCase());
  const isConsolidado = headers.length <= 3 || headers.join('').includes('consolidado');
  
  const rows = data.slice(1);
  return rows.map(row => {
    let promedio = 0;
    if (isConsolidado) {
      promedio = Number(row[2]) || 0;
    } else {
      let m1 = Number(row[2]) || 0;
      let m2 = Number(row[3]) || 0;
      let m3 = Number(row[4]) || 0;
      let m4 = Number(row[5]) || 0;
      promedio = (m1 + m2 + m3 + m4) / 4;
    }
    
    return {
      codigo: String(row[0] || '').trim(),
      producto: String(row[1] || '').trim(),
      promedio_mensual: promedio,
      presentacion: (isConsolidado && row[3]) ? String(row[3]).trim() : 'UNIDADES'
    };
  });
}

// ==========================================
// MÓDULO PECOSA: GUARDAR DATOS HISTÓRICOS DESDE UPLOAD
// ==========================================
function guardarHistoricoData(datosAgrupados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Histórico Consumo");
  if (!sheet) {
    sheet = ss.insertSheet("Histórico Consumo");
  }
  
  sheet.clearContents();
  sheet.appendRow(["Código", "Producto", "Promedio Mensual (Consolidado IA)", "Presentación"]);
  sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#3b82f6").setFontColor("#ffffff");
  sheet.setColumnWidth(2, 300);
  
  if (datosAgrupados && datosAgrupados.length > 0) {
    const rows = datosAgrupados.map(d => [d.codigo, d.producto, d.promedio_mensual, d.presentacion || 'UNIDADES']);
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
  
  SpreadsheetApp.flush();
  return "Datos históricos guardados exitosamente (" + datosAgrupados.length + " productos consolidados).";
}
