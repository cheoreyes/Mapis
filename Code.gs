/**
 * MAPIS MVP - BACKEND GOOGLE APPS SCRIPT
 * Versión: 2.0 (Auth Docente + Datos Reales + Fotos)
 */

// ============================================================================
// 1. SERVICIO WEB (DOGET)
// ============================================================================
function doGet(e) {
  const page = e.parameter.page || 'docente';
  const metaViewport = 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no';
  
  if (page === 'estudiante') {
    return HtmlService.createHtmlOutputFromFile('EstudianteUI')
      .setTitle('MAPIS - Mi Perfil')
      .addMetaTag('viewport', metaViewport)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } 
  
  return HtmlService.createHtmlOutputFromFile('DocenteUI')
    .setTitle('MAPIS - Acceso Docente')
    .addMetaTag('viewport', metaViewport)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================================
// 2. API DOCENTE: VALIDAR PIN DE ACCESO
// Busca en la hoja 'DOCENTES' si el PIN es válido
// ============================================================================
function validarPinDocente(pinIngresado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetDoc = ss.getSheetByName('DOCENTES');
    
    if (!sheetDoc) return { success: false, error: "Hoja DOCENTES no encontrada" };
    
    const docentes = sheetDoc.getDataRange().getValues();
    
    // Recorremos buscando coincidencia (Columna C = índice 2 es el PIN)
    for (let i = 1; i < docentes.length; i++) {
      // Convertimos a string para evitar errores de formato numérico
      if (String(docentes[i][2]).trim() === String(pinIngresado).trim()) {
        return { 
          success: true, 
          idDocente: docentes[i][0], // Columna A
          nombreDocente: docentes[i][1] // Columna B
        };
      }
    }
    
    return { success: false, error: "PIN incorrecto o no registrado" };
    
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================================
// 3. API DOCENTE: BUSCAR ESTUDIANTE (DATOS REALES + FOTO)
// ============================================================================
function buscarEstudianteReal(hashQR) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    
    if (!sheetEst) return { success: false, error: "Hoja ESTUDIANTES no encontrada" };
    
    const estudiantes = sheetEst.getDataRange().getValues();
    
    for (let i = 1; i < estudiantes.length; i++) {
      // Columna F (índice 5) es el Hash QR
      if (String(estudiantes[i][5]).trim() === String(hashQR).trim()) { 
        
        // Obtener foto (Columna G, índice 6)
        let fotoUrl = estudiantes[i][6];
        if (!fotoUrl || String(fotoUrl).trim() === "") {
           // Avatar por defecto basado en nombre si no hay foto
           fotoUrl = "https://api.dicebear.com/7.x/avataaars/svg?seed=" + encodeURIComponent(estudiantes[i][1]) + "&backgroundColor=c0c0c0";
        }

        return {
            success: true,
            idEstudiante: estudiantes[i][0], // Columna A
            nombre: estudiantes[i][1],       // Columna B
            grado: estudiantes[i][2],        // Columna C
            saldo: Number(estudiantes[i][3]),// Columna D
            rango: estudiantes[i][4],        // Columna E
            foto: fotoUrl
        };
      }
    }
    
    return { success: false, error: "Estudiante no encontrado con ese ID/QR" };
    
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================================
// 4. API DOCENTE: REGISTRAR TRANSACCIÓN (CON AUTH)
// ============================================================================
function registrarTransaccion(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTrans = ss.getSheetByName('TRANSACCIONES');
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    
    if (!sheetTrans || !sheetEst) {
      return { success: false, error: "Error BD: Faltan hojas" };
    }
    
    // Buscar estudiante para obtener fila y saldo actual
    const estudiantes = sheetEst.getDataRange().getValues();
    let filaEstudiante = -1;
    let saldoActual = 0;
    let nombreEstudiante = "";
    
    for (let i = 1; i < estudiantes.length; i++) {
      if (String(estudiantes[i][5]).trim() === String(data.hashQR).trim()) { 
        filaEstudiante = i + 1;
        saldoActual = Number(estudiantes[i][3]);
        nombreEstudiante = estudiantes[i][1];
        break;
      }
    }
    
    if (filaEstudiante === -1) {
      return { success: false, error: "Estudiante no encontrado (posible cambio de ID)" };
    }
    
    // Calcular y actualizar saldo
    const puntos = Number(data.puntos);
    const nuevoSaldo = saldoActual + puntos;
    sheetEst.getRange(filaEstudiante, 4).setValue(nuevoSaldo);
    
    // Registrar en Ledger
    sheetTrans.appendRow([
      new Date(),
      nombreEstudiante,
      data.idDocente || 'DOC-UNKNOWN', // Usamos el ID real del docente logueado
      data.tipoConducta,
      puntos,
      data.nota || "",
      data.categoria || "General"
    ]);
    
    return { success: true, nuevoSaldo: nuevoSaldo };
    
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 5. API ESTUDIANTE: CONSULTAR DATOS (Sin cambios mayores)
// ============================================================================
function obtenerDatosEstudiante(hashQR) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    const sheetTrans = ss.getSheetByName('TRANSACCIONES');
    
    if (!sheetEst || !sheetTrans) return { error: "BD no configurada" };
    
    const estudiantes = sheetEst.getDataRange().getValues();
    let datosEst = null;
    
    for (let i = 1; i < estudiantes.length; i++) {
      if (String(estudiantes[i][5]).trim() === String(hashQR).trim()) {
        let fotoUrl = estudiantes[i][6]; 
        if (!fotoUrl || String(fotoUrl).trim() === "") {
           fotoUrl = "https://api.dicebear.com/7.x/avataaars/svg?seed=" + encodeURIComponent(estudiantes[i][1]) + "&backgroundColor=c0c0c0";
        }

        datosEst = {
          nombre: estudiantes[i][1],
          grado: estudiantes[i][2],
          saldo: Number(estudiantes[i][3]),
          rango: estudiantes[i][4],
          foto: fotoUrl
        };
        break;
      }
    }
    
    if (!datosEst) return { error: "Estudiante no encontrado" };
    
    const transacciones = sheetTrans.getDataRange().getValues();
    const historial = [];
    
    for (let i = transacciones.length - 1; i >= 1 && historial.length < 10; i--) {
      if (String(transacciones[i][1]) === String(datosEst.nombre)) {
        historial.push({
          fecha: Utilities.formatDate(new Date(transacciones[i][0]), Session.getScriptTimeZone(), "dd/MM hh:mm"),
          descripcion: transacciones[i][3],
          puntos: Number(transacciones[i][4]),
          responsable: transacciones[i][2]
        });
      }
    }
    
    return { ...datosEst, historial: historial };
    
  } catch (err) {
    return { error: err.toString() };
  }
}
