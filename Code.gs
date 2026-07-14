/**
 * MAPIS MVP - BACKEND COMPLETO V2.5 (Con Toast Logic)
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
// 2. OBTENER LISTA DE ALUMNOS (Para el buscador por nombre)
// ============================================================================
function obtenerListaAlumnos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    if (!sheetEst) return [];
    
    const datos = sheetEst.getDataRange().getValues();
    const lista = [];
    
    // Saltamos encabezado (i=1)
    for (let i = 1; i < datos.length; i++) {
      // Solo agregamos si tiene nombre y hash
      if(datos[i][1] && datos[i][5]) {
        lista.push({
          id: datos[i][0],       // Columna A
          nombre: datos[i][1],   // Columna B (Nombre Completo)
          hash: datos[i][5]      // Columna F (Hash QR)
        });
      }
    }
    return lista;
  } catch (e) {
    Logger.log("Error obteniendo lista: " + e.toString());
    return [];
  }
}

// ============================================================================
// 3. VALIDAR PIN DOCENTE
// ============================================================================
function validarPinDocente(pinIngresado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetDoc = ss.getSheetByName('DOCENTES');
    if (!sheetDoc) return { success: false, error: "Hoja DOCENTES no encontrada" };
    
    const docentes = sheetDoc.getDataRange().getValues();
    for (let i = 1; i < docentes.length; i++) {
      if (String(docentes[i][2]).trim() === String(pinIngresado).trim()) {
        return { success: true, idDocente: docentes[i][0], nombreDocente: docentes[i][1] };
      }
    }
    return { success: false, error: "PIN incorrecto o no registrado" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================================
// 4. BUSCAR ESTUDIANTE POR HASH (Detalles completos)
// ============================================================================
function buscarEstudiantePorHash(hashQR) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    if (!sheetEst) return { success: false, error: "Hoja ESTUDIANTES no encontrada" };
    
    const estudiantes = sheetEst.getDataRange().getValues();
    for (let i = 1; i < estudiantes.length; i++) {
      if (String(estudiantes[i][5]).trim() === String(hashQR).trim()) { 
        let fotoUrl = estudiantes[i][6]; // Columna G
        if (!fotoUrl || String(fotoUrl).trim() === "") {
           fotoUrl = "https://api.dicebear.com/7.x/avataaars/svg?seed=" + encodeURIComponent(estudiantes[i][1]) + "&backgroundColor=c0c0c0";
        }
        return {
            success: true,
            idEstudiante: estudiantes[i][0],
            nombre: estudiantes[i][1],
            grado: estudiantes[i][2],
            saldo: Number(estudiantes[i][3]),
            rango: estudiantes[i][4],
            foto: fotoUrl,
            hash: estudiantes[i][5] // Importante devolver el hash para usarlo al registrar
        };
      }
    }
    return { success: false, error: "Alumno no encontrado" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================================
// 5. REGISTRAR TRANSACCIÓN
// ============================================================================
function registrarTransaccion(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTrans = ss.getSheetByName('TRANSACCIONES');
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    
    if (!sheetTrans || !sheetEst) return { success: false, error: "Error de configuración en BD" };
    
    // Buscar fila del estudiante por Hash
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
    
    if (filaEstudiante === -1) return { success: false, error: "Alumno no encontrado (posible cambio de ID)" };
    
    const puntos = Number(data.puntos);
    const nuevoSaldo = saldoActual + puntos;
    
    // Actualizar saldo
    sheetEst.getRange(filaEstudiante, 4).setValue(nuevoSaldo);
    
    // Registrar en Ledger (Append-only)
    sheetTrans.appendRow([
      new Date(), 
      nombreEstudiante, 
      data.idDocente || 'DOC-UNKNOWN',
      data.tipoConducta, 
      puntos, 
      "", // Nota vacía por defecto en esta versión rápida
      data.categoria || "General"
    ]);
    
    return { success: true, nuevoSaldo: nuevoSaldo, nombre: nombreEstudiante };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 6. OBTENER DATOS ESTUDIANTE (Con estadísticas dinámicas)
// ============================================================================
function obtenerDatosEstudiante(hashQR) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    const sheetTrans = ss.getSheetByName('TRANSACCIONES');
    
    if (!sheetEst || !sheetTrans) return { error: "BD no configurada" };
    
    // Buscar estudiante
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
    
    if (!datosEst) return { error: "No encontrado" };
    
    // Calcular Estadísticas
    const transacciones = sheetTrans.getDataRange().getValues();
    const historial = [];
    let puntosSemana = 0;
    let conductasPositivas = 0;
    
    const hoy = new Date();
    const hace7Dias = new Date(hoy.getTime() - (7 * 24 * 60 * 60 * 1000));
    
    // Recorrer desde el final (más reciente)
    for (let i = transacciones.length - 1; i >= 1; i--) {
      const row = transacciones[i];
      const fechaTrans = new Date(row[0]);
      const nombreTrans = String(row[1]);
      const puntosTrans = Number(row[4]);
      
      if (nombreTrans === datosEst.nombre) {
        // Historial (limitado a 10)
        if (historial.length < 10) {
            historial.push({
              fecha: Utilities.formatDate(fechaTrans, Session.getScriptTimeZone(), "dd/MM hh:mm"),
              descripcion: row[3],
              puntos: puntosTrans,
              responsable: row[2]
            });
        }
        
        // Puntos últimos 7 días
        if (fechaTrans >= hace7Dias) {
            puntosSemana += puntosTrans;
        }
        
        // Conteo total de conductas positivas
        if (puntosTrans > 0) {
            conductasPositivas++;
        }
      }
    }
    
    return { 
        ...datosEst, 
        historial: historial,
        puntosSemana: puntosSemana,
        conductasPositivas: conductasPositivas
    };
    
  } catch (err) { 
      return { error: err.toString() }; 
  }
}
