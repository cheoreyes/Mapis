/**
 * MAPIS v7.5 - BACKEND INSTITUCIONAL
 * U.E. Mariano Picón Salas - Los Robles, Nueva Esparta
 * 
 * CAMBIOS v7.5:
 * - Notificación al estudiante cuando un docente registra una incidencia
 * - Conteo correcto de incidencias (incluye Faltas Graves)
 * - Ya no se cuentan conductas positivas/bonos
 */

// ============================================================================
// 1. SERVICIO WEB
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
  
  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('AdminUI')
      .setTitle('MAPIS - Control de Estudio')
      .addMetaTag('viewport', metaViewport)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('DocenteUI')
    .setTitle('MAPIS - Acceso Docente')
    .addMetaTag('viewport', metaViewport)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================================
// 2. ASEGURAR HOJAS DE CONFIGURACIÓN
// ============================================================================
function asegurarHojasConfiguracion() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    let sheetCond = ss.getSheetByName('CONFIG_CONDUCTAS');
    if (!sheetCond) {
      sheetCond = ss.insertSheet('CONFIG_CONDUCTAS');
      sheetCond.appendRow(['ID_Conducta', 'Nombre_Display', 'Puntos', 'Categoria']);
      sheetCond.appendRow(['COND01', 'Falta de Responsabilidad', -50, 'Responsabilidad']);
      sheetCond.appendRow(['COND02', 'Uso indebido de Tecnología', -30, 'Ciberética']);
      sheetCond.appendRow(['COND03', 'Falta de Respeto', -50, 'Empatía y Respeto']);
      sheetCond.appendRow(['COND04', 'Impuntualidad', -20, 'Puntualidad']);
      sheetCond.appendRow(['COND05', 'Desregulación Emocional', -30, 'Autorregulación']);
      sheetCond.appendRow(['COND06', 'Falta de 3er Nivel', -9999, 'Faltas Graves']);
      sheetCond.appendRow(['COND07', 'Otros', -25, 'General']);
    }
    
    let sheetBonos = ss.getSheetByName('CONFIG_BONOS');
    if (!sheetBonos) {
      sheetBonos = ss.insertSheet('CONFIG_BONOS');
      sheetBonos.appendRow(['Incidencias_Min', 'Incidencias_Max', 'Puntos_Bono']);
      sheetBonos.appendRow([0, 0, 500]);
      sheetBonos.appendRow([1, 2, 100]);
      sheetBonos.appendRow([3, 4, 200]);
      sheetBonos.appendRow([5, 99, 0]);
    }
    
    let sheetPremios = ss.getSheetByName('CONFIG_PREMIOS');
    if (!sheetPremios) {
      sheetPremios = ss.insertSheet('CONFIG_PREMIOS');
      sheetPremios.appendRow(['ID_Premio', 'Nombre', 'Costo', 'Descripcion']);
      sheetPremios.appendRow(['PREM01', 'Tiempo Extra de Receso', 50, '+10 min en el receso']);
      sheetPremios.appendRow(['PREM02', 'Día sin Uniforme', 100, 'Vestir ropa particular']);
      sheetPremios.appendRow(['PREM03', 'Extensión de Entrega 24h', 75, '+24 horas para una tarea']);
      sheetPremios.appendRow(['PREM04', 'Extensión de Entrega 48h', 150, '+48 horas para una tarea']);
      sheetPremios.appendRow(['PREM05', 'Compra en Tienda', 200, 'Artículo de la tienda escolar']);
    }
    
    let sheetSol = ss.getSheetByName('SOLICITUDES_CANJE');
    if (!sheetSol) {
      sheetSol = ss.insertSheet('SOLICITUDES_CANJE');
      sheetSol.appendRow(['Fecha', 'ID_Estudiante', 'Nombre_Estudiante', 'ID_Premio', 'Nombre_Premio', 'Costo', 'Estado', 'ID_Docente_Aprobador', 'Fecha_Aprobacion']);
    }
    
    let sheetGuia = ss.getSheetByName('CONFIG_GUIA');
    if (!sheetGuia) {
      sheetGuia = ss.insertSheet('CONFIG_GUIA');
      sheetGuia.appendRow(['ID_Docente', 'Grado_Asignado']);
    }
    
    let sheetRoles = ss.getSheetByName('CONFIG_ROLES');
    if (!sheetRoles) {
      sheetRoles = ss.insertSheet('CONFIG_ROLES');
      sheetRoles.appendRow(['ID_Docente', 'Rol']);
    }
    
    let sheetEst = ss.getSheetByName('ESTUDIANTES');
    if (!sheetEst) {
      sheetEst = ss.insertSheet('ESTUDIANTES');
      sheetEst.appendRow(['ID_Unico', 'Nombre_Completo', 'Grado', 'Saldo_Actual', 'Rango', 'Email']);
    }
    
    let sheetSug = ss.getSheetByName('SUGERENCIAS');
    if (!sheetSug) {
      sheetSug = ss.insertSheet('SUGERENCIAS');
      sheetSug.appendRow(['Fecha', 'Autor', 'Rol_Autor', 'Tipo', 'Mensaje', 'Estado']);
    }
    
    let sheetTienda = ss.getSheetByName('TIENDA_PRODUCTOS');
    if (!sheetTienda) {
      sheetTienda = ss.insertSheet('TIENDA_PRODUCTOS');
      sheetTienda.appendRow(['ID_Producto', 'Nombre', 'Costo', 'Descripcion', 'Foto_URL', 'Stock']);
    }
    
    let sheetNotif = ss.getSheetByName('NOTIFICACIONES');
    if (!sheetNotif) {
      sheetNotif = ss.insertSheet('NOTIFICACIONES');
      sheetNotif.appendRow(['Fecha', 'Titulo', 'Mensaje', 'Fecha_Expiracion']);
    }
    
    let sheetNotifEst = ss.getSheetByName('NOTIFICACIONES_ESTUDIANTE');
    if (!sheetNotifEst) {
      sheetNotifEst = ss.insertSheet('NOTIFICACIONES_ESTUDIANTE');
      sheetNotifEst.appendRow(['Timestamp', 'ID_Estudiante', 'Titulo', 'Mensaje', 'Leida']);
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================================
// 3. OBTENER CONFIGURACIONES
// ============================================================================
function obtenerConfiguracion() {
  try {
    asegurarHojasConfiguracion();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sheetCond = ss.getSheetByName('CONFIG_CONDUCTAS');
    const conductas = sheetCond ? sheetCond.getDataRange().getValues().slice(1).map(r => ({
      id: r[0], nombre: r[1], puntos: Number(r[2]), categoria: r[3]
    })).filter(c => c.id && c.nombre) : [];

    const sheetBonos = ss.getSheetByName('CONFIG_BONOS');
    const bonos = sheetBonos ? sheetBonos.getDataRange().getValues().slice(1).map(r => ({
      min: Number(r[0]), max: Number(r[1]), puntos: Number(r[2])
    })).filter(b => !isNaN(b.min)) : [];

    const sheetPremios = ss.getSheetByName('CONFIG_PREMIOS');
    const premios = sheetPremios ? sheetPremios.getDataRange().getValues().slice(1).map(r => ({
      id: r[0], nombre: r[1], costo: Number(r[2]), descripcion: r[3]
    })).filter(p => p.id && p.nombre) : [];

    return { success: true, conductas, bonos, premios };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================================
// 4. OBTENER LISTA DE ALUMNOS
// ============================================================================
function obtenerListaAlumnos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    if (!sheetEst) return [];
    
    const datos = sheetEst.getDataRange().getValues();
    const lista = [];
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][1]) {
        lista.push({
          id: String(datos[i][0]).trim().toUpperCase(),
          nombre: String(datos[i][1]).trim(),
          grado: String(datos[i][2] || '').trim()
        });
      }
    }
    return lista;
  } catch (e) {
    return [];
  }
}

// ============================================================================
// 5. VALIDAR PIN DOCENTE
// ============================================================================
function validarPinDocente(pinIngresado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetDoc = ss.getSheetByName('DOCENTES');
    
    if (!sheetDoc) {
      const nuevaHoja = ss.insertSheet('DOCENTES');
      nuevaHoja.appendRow(['ID_Docente', 'Nombre', 'PIN_Acceso']);
      nuevaHoja.appendRow(['DOC01', 'Prof. Demo', '1234']);
      return { success: true, idDocente: 'DOC01', nombreDocente: 'Prof. Demo', gradosAsignados: [], rol: null };
    }
    
    const docentes = sheetDoc.getDataRange().getValues();
    
    if (docentes.length <= 1) {
      return { success: false, error: "No hay docentes registrados." };
    }
    
    for (let i = 1; i < docentes.length; i++) {
      if (String(docentes[i][2]).trim() === String(pinIngresado).trim()) {
        const idDoc = String(docentes[i][0]);
        const gradosAsignados = obtenerGradosDelDocente(idDoc).grados || [];
        const rol = obtenerRolDelDocente(idDoc);
        return { 
          success: true, 
          idDocente: idDoc, 
          nombreDocente: String(docentes[i][1]),
          gradosAsignados: gradosAsignados,
          rol: rol
        };
      }
    }
    return { success: false, error: "PIN incorrecto" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================================
// 6. BUSCAR ESTUDIANTE POR ID_UNICO
// ============================================================================
function buscarEstudiantePorHash(codigo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    if (!sheetEst) return { success: false, error: "Hoja ESTUDIANTES no encontrada" };
    
    const estudiantes = sheetEst.getDataRange().getValues();
    const busqueda = String(codigo).trim().toLowerCase();
    
    for (let i = 1; i < estudiantes.length; i++) {
      const idEst = String(estudiantes[i][0]).trim();
      if (idEst.toLowerCase() === busqueda) {
        return {
            success: true,
            idEstudiante: idEst.toUpperCase(),
            nombre: String(estudiantes[i][1]),
            grado: String(estudiantes[i][2] || ''),
            saldo: Number(estudiantes[i][3]) || 0,
            rango: String(estudiantes[i][4] || 'Explorador')
        };
      }
    }
    return { success: false, error: "Alumno no encontrado (" + codigo + ")" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================================
// 7. REGISTRAR INCIDENCIA (con notificación al estudiante)
// ============================================================================
function registrarIncidencia(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTrans = ss.getSheetByName('TRANSACCIONES');
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    
    if (!sheetTrans || !sheetEst) return { success: false, error: "Error BD" };
    
    if (data.categoria === 'Faltas Graves') {
      const rol = obtenerRolDelDocente(data.idDocente);
      if (rol !== 'coordinador' && rol !== 'director') {
        return { success: false, error: 'Solo el coordinador o director puede aplicar la Falta de 3er Nivel.' };
      }
    }
    
    const estudiantes = sheetEst.getDataRange().getValues();
    let filaEstudiante = -1;
    let saldoActual = 0;
    let idEstudiante = "";
    
    const busqueda = String(data.idEstudiante).trim().toLowerCase();
    
    for (let i = 1; i < estudiantes.length; i++) {
      if (String(estudiantes[i][0]).trim().toLowerCase() === busqueda) { 
        filaEstudiante = i + 1;
        idEstudiante = String(estudiantes[i][0]).trim().toUpperCase();
        saldoActual = Number(estudiantes[i][3]);
        break;
      }
    }
    
    if (filaEstudiante === -1) return { success: false, error: "Alumno no encontrado" };
    
    const puntos = Number(data.puntos);
    let nuevoSaldo;
    let puntosRegistrados = puntos;
    
    if (data.categoria === 'Faltas Graves') {
      const transacciones = sheetTrans.getDataRange().getValues();
      const hace3Meses = new Date(new Date().getTime() - (90 * 24 * 60 * 60 * 1000));
      let faltasGraves = 0;
      
      for (let j = 1; j < transacciones.length; j++) {
        const fechaT = new Date(transacciones[j][0]);
        const idT = String(transacciones[j][1]).trim().toLowerCase();
        const categoriaT = String(transacciones[j][6] || '');
        if (fechaT >= hace3Meses && idT === busqueda && categoriaT === 'Faltas Graves') {
          faltasGraves++;
        }
      }
      
      let porcentaje = 0.25;
      if (faltasGraves === 1) porcentaje = 0.50;
      if (faltasGraves >= 2) porcentaje = 1.00;
      
      const descuento = Math.round(saldoActual * porcentaje);
      nuevoSaldo = saldoActual - descuento;
      puntosRegistrados = -descuento;
    } else {
      nuevoSaldo = saldoActual + puntos;
    }
    
    sheetEst.getRange(filaEstudiante, 4).setValue(nuevoSaldo);
    
    sheetTrans.appendRow([
      new Date(), 
      idEstudiante,
      data.idDocente || 'DOC-UNKNOWN',
      data.tipoConducta, 
      puntosRegistrados, 
      data.nota || "", 
      data.categoria || "General"
    ]);
    
    // CREAR NOTIFICACIÓN PARA EL ESTUDIANTE
    try {
      let nombreDocente = 'Un docente';
      const sheetDoc = ss.getSheetByName('DOCENTES');
      if (sheetDoc) {
        const docentes = sheetDoc.getDataRange().getValues();
        for (let k = 1; k < docentes.length; k++) {
          if (String(docentes[k][0]).trim() === String(data.idDocente).trim()) {
            nombreDocente = String(docentes[k][1]);
            break;
          }
        }
      }
      
      const categoria = data.categoria || 'General';
      let titulo = '';
      let mensaje = '';
      
      if (categoria === 'Faltas Graves') {
        titulo = '⚠️ Falta de 3er Nivel';
        mensaje = `${nombreDocente} registró una Falta de 3er Nivel en tu expediente. Descuento aplicado: ${Math.abs(puntosRegistrados)} pts. Nuevo saldo: ${nuevoSaldo} pts.${data.nota ? ' Observación: ' + data.nota : ''}`;
      } else {
        titulo = '📋 Nueva incidencia';
        mensaje = `${nombreDocente} registró la incidencia "${data.tipoConducta}" (${puntosRegistrados} pts). Nuevo saldo: ${nuevoSaldo} pts.${data.nota ? ' Observación: ' + data.nota : ''}`;
      }
      
      crearNotificacionEstudiante(idEstudiante, titulo, mensaje);
    } catch (notifErr) {
      console.error('Error creando notificación:', notifErr);
    }
    
    return { success: true, nuevoSaldo: nuevoSaldo };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 8. PROCESAR CARGA SEMANAL
// ============================================================================
function procesarCargaSemanal() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    const sheetTrans = ss.getSheetByName('TRANSACCIONES');
    const sheetBonos = ss.getSheetByName('CONFIG_BONOS');
    
    if (!sheetEst || !sheetTrans || !sheetBonos) return { success: false, error: "Faltan hojas" };
    
    const bonosConfig = sheetBonos.getDataRange().getValues().slice(1).map(r => ({
      min: Number(r[0]), max: Number(r[1]), puntos: Number(r[2])
    })).sort((a,b) => a.min - b.min);
    
    const estudiantes = sheetEst.getDataRange().getValues();
    const hoy = new Date();
    const hace7Dias = new Date(hoy.getTime() - (7 * 24 * 60 * 60 * 1000));
    
    const transacciones = sheetTrans.getDataRange().getValues();
    const incidenciasPorEstudiante = {};
    
    for (let i = 1; i < transacciones.length; i++) {
      const fecha = new Date(transacciones[i][0]);
      if (fecha >= hace7Dias && Number(transacciones[i][4]) < 0) { 
        const idEst = String(transacciones[i][1]).trim().toUpperCase();
        incidenciasPorEstudiante[idEst] = (incidenciasPorEstudiante[idEst] || 0) + 1;
      }
    }
    
    let actualizados = 0;
    for (let i = 1; i < estudiantes.length; i++) {
      const idEst = String(estudiantes[i][0]).trim().toUpperCase();
      const saldoActual = Number(estudiantes[i][3]);
      const incidencias = incidenciasPorEstudiante[idEst] || 0;
      
      let bonoAplicable = bonosConfig.find(b => incidencias >= b.min && incidencias <= b.max);
      if (!bonoAplicable) bonoAplicable = { puntos: 0 };
      
      const nuevoSaldo = saldoActual + bonoAplicable.puntos;
      
      sheetEst.getRange(i + 1, 4).setValue(nuevoSaldo);
      
      sheetTrans.appendRow([
        new Date(), idEst, "SISTEMA-MAPIS", "BONO SEMANAL",
        bonoAplicable.puntos, `${incidencias} incidencias registradas`, "Bono"
      ]);
      
      actualizados++;
    }
    
    return { success: true, message: `Carga semanal completada. ${actualizados} estudiantes actualizados.` };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 9. OBTENER DATOS ESTUDIANTE (con conteo correcto de incidencias)
// ============================================================================
function obtenerDatosEstudiante(identificador) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    const sheetTrans = ss.getSheetByName('TRANSACCIONES');
    
    if (!sheetEst || !sheetTrans) return { error: "BD no configurada" };
    
    const estudiantes = sheetEst.getDataRange().getValues();
    const busqueda = String(identificador).trim().toLowerCase();
    let datosEst = null;
    let idEstudiante = null;
    
    for (let i = 1; i < estudiantes.length; i++) {
      const idRow = String(estudiantes[i][0]).trim().toLowerCase();
      
      if (idRow === busqueda) {
        idEstudiante = String(estudiantes[i][0]).trim().toUpperCase();
        datosEst = { 
          id: idEstudiante,
          nombre: estudiantes[i][1], 
          grado: estudiantes[i][2], 
          saldo: Number(estudiantes[i][3]), 
          rango: estudiantes[i][4]
        };
        break;
      }
    }
    
    if (!datosEst) return { error: "Estudiante no encontrado. Verifica tu código." };
    
    const transacciones = sheetTrans.getDataRange().getValues();
    const historial = [];
    let puntosSemana = 0;
    let totalIncidencias = 0;
    
    const hoy = new Date();
    const hace7Dias = new Date(hoy.getTime() - (7 * 24 * 60 * 60 * 1000));
    const idEstLower = idEstudiante.toLowerCase();
    
    for (let i = transacciones.length - 1; i >= 1; i--) {
      const row = transacciones[i];
      const fechaTrans = new Date(row[0]);
      const idTrans = String(row[1]).trim().toLowerCase();
      const puntosTrans = Number(row[4]);
      const tipo = String(row[3] || '');
      
      if (idTrans === idEstLower) {
        if (historial.length < 10) {
            historial.push({
              fecha: Utilities.formatDate(fechaTrans, Session.getScriptTimeZone(), "dd/MM HH:mm"),
              descripcion: row[3],
              puntos: puntosTrans,
              responsable: row[2],
              nota: row[5] || ''
            });
        }
        
        if (fechaTrans >= hace7Dias) puntosSemana += puntosTrans;
        
        const esCanje = tipo.includes('CANJE');
        const esBono = tipo.includes('BONO');
        const esIncidencia = !esCanje && !esBono && puntosTrans < 0;
        
        if (esIncidencia) {
          totalIncidencias++;
        }
      }
    }
    
    return { 
        ...datosEst, 
        historial: historial,
        puntosSemana: puntosSemana,
        totalIncidencias: totalIncidencias
    };
    
  } catch (err) { 
      return { error: err.toString() }; 
  }
}

// ============================================================================
// 10. CREAR ESTUDIANTE
// ============================================================================
function crearEstudiante(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetEst = ss.getSheetByName('ESTUDIANTES');
    
    if (!sheetEst) {
      sheetEst = ss.insertSheet('ESTUDIANTES');
      sheetEst.appendRow(['ID_Unico', 'Nombre_Completo', 'Grado', 'Saldo_Actual', 'Rango', 'Email']);
    }

    const cedula = String(data.cedula).trim();
    const siglas = String(data.siglas).trim().toUpperCase();
    const anio = String(data.anio).trim();
    const codigo = (cedula + siglas + anio).toUpperCase();
    
    const existentes = sheetEst.getDataRange().getValues();
    for (let i = 1; i < existentes.length; i++) {
      if (String(existentes[i][0]).trim().toLowerCase() === codigo.toLowerCase()) {
        return { success: false, error: "El código " + codigo + " ya existe" };
      }
    }
    
    const saldoInicial = Number(data.saldoInicial) || 0;
    
    sheetEst.appendRow([
      codigo, data.nombre, data.grado, saldoInicial, 'Explorador', data.email || ''
    ]);

    return { success: true, codigo: codigo, message: "Estudiante creado exitosamente" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================================
// 11. CREAR DOCENTE
// ============================================================================
function crearDocente(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetDoc = ss.getSheetByName('DOCENTES');
    
    if (!sheetDoc) {
      sheetDoc = ss.insertSheet('DOCENTES');
      sheetDoc.appendRow(['ID_Docente', 'Nombre', 'PIN_Acceso']);
    }

    if (String(data.pin).length !== 4) return { success: false, error: "El PIN debe tener 4 dígitos" };

    const existentes = sheetDoc.getDataRange().getValues();
    for (let i = 1; i < existentes.length; i++) {
      if (String(existentes[i][2]) === String(data.pin)) {
        return { success: false, error: "El PIN " + data.pin + " ya está en uso" };
      }
    }

    sheetDoc.appendRow([data.id || 'DOC-' + new Date().getTime(), data.nombre, data.pin]);
    return { success: true, message: "Docente creado exitosamente" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================================
// 12. VALIDAR ADMIN
// ============================================================================
function validarAdmin(password) {
  if (password === "mapis2026") { 
    return { success: true }; 
  }
  return { success: false, error: "Contraseña incorrecta" };
}

// ============================================================================
// 13. LISTA PARA EDICIÓN
// ============================================================================
function obtenerListaParaEditar(tipo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (tipo === 'estudiante') {
      const sheet = ss.getSheetByName('ESTUDIANTES');
      if (!sheet) return [];
      const data = sheet.getDataRange().getValues();
      return data.slice(1).map(r => ({ 
        id: String(r[0]).trim().toUpperCase(), 
        nombre: String(r[1]).trim(), 
        valor2: String(r[2] || '').trim(), 
        valor3: String(r[5] || '').trim() 
      })).filter(e => e.id);
    } else {
      const sheet = ss.getSheetByName('DOCENTES');
      if (!sheet) return [];
      const data = sheet.getDataRange().getValues();
      return data.slice(1).map(r => ({ 
        id: String(r[0]).trim(), 
        nombre: String(r[1]).trim(), 
        valor2: String(r[2] || '').trim(), 
        valor3: '' 
      })).filter(e => e.id);
    }
  } catch (e) { return []; }
}

// ============================================================================
// 14. EDITAR ESTUDIANTE
// ============================================================================
function editarEstudiante(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ESTUDIANTES');
    const rows = sheet.getDataRange().getValues();
    
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === String(data.id).trim().toLowerCase()) {
        sheet.getRange(i + 1, 2).setValue(data.nombre);
        sheet.getRange(i + 1, 3).setValue(data.grado);
        sheet.getRange(i + 1, 6).setValue(data.email || '');
        return { success: true, message: "Estudiante actualizado" };
      }
    }
    return { success: false, error: "ID no encontrado" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ============================================================================
// 15. EDITAR DOCENTE
// ============================================================================
function editarDocente(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DOCENTES');
    const rows = sheet.getDataRange().getValues();
    
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.id).trim()) {
        sheet.getRange(i + 1, 2).setValue(data.nombre);
        sheet.getRange(i + 1, 3).setValue(data.pin);
        return { success: true, message: "Docente actualizado" };
      }
    }
    return { success: false, error: "ID no encontrado" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ============================================================================
// 16. GRADOS DEL DOCENTE
// ============================================================================
function obtenerGradosDelDocente(idDocente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('CONFIG_GUIA');
    if (!sheet) return { success: false, grados: [] };
    
    const data = sheet.getDataRange().getValues();
    const grados = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idDocente).trim()) {
        grados.push(String(data[i][1]).trim());
      }
    }
    return { success: true, grados: grados };
  } catch (e) {
    return { success: false, grados: [], error: e.toString() };
  }
}

// ============================================================================
// 16.1 ROL DEL DOCENTE
// ============================================================================
function obtenerRolDelDocente(idDocente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('CONFIG_ROLES');
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idDocente).trim()) {
        return String(data[i][1]).trim().toLowerCase();
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

// ============================================================================
// 17. GESTIÓN DE GUÍAS
// ============================================================================
function obtenerListaGuias() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetGuia = ss.getSheetByName('CONFIG_GUIA');
    const sheetDoc = ss.getSheetByName('DOCENTES');
    if (!sheetGuia) return [];
    
    const guias = sheetGuia.getDataRange().getValues().slice(1);
    const docentes = sheetDoc.getDataRange().getValues().slice(1);
    
    return guias.map(g => {
      const idDoc = String(g[0]).trim();
      const grado = String(g[1]).trim();
      const docente = docentes.find(d => String(d[0]).trim() === idDoc);
      return {
        idDocente: idDoc,
        nombreDocente: docente ? String(docente[1]) : 'Desconocido',
        gradoAsignado: grado
      };
    }).filter(g => g.idDocente && g.gradoAsignado);
  } catch (e) { return []; }
}

function asignarGuia(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('CONFIG_GUIA');
    if (!sheet) {
      sheet = ss.insertSheet('CONFIG_GUIA');
      sheet.appendRow(['ID_Docente', 'Grado_Asignado']);
    }
    
    const dataActual = sheet.getDataRange().getValues();
    for (let i = 1; i < dataActual.length; i++) {
      if (String(dataActual[i][0]).trim() === String(data.idDocente).trim() &&
          String(dataActual[i][1]).trim() === String(data.gradoAsignado).trim()) {
        return { success: false, error: 'Esa asignación ya existe' };
      }
    }
    
    sheet.appendRow([data.idDocente, data.gradoAsignado]);
    return { success: true, message: 'Guía asignado' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function eliminarGuia(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('CONFIG_GUIA');
    const rows = sheet.getDataRange().getValues();
    
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(data.idDocente).trim() &&
          String(rows[i][1]).trim() === String(data.grado).trim()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'No encontrado' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================================
// 17.1 GESTIÓN DE ROLES
// ============================================================================
function obtenerListaRoles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRoles = ss.getSheetByName('CONFIG_ROLES');
    const sheetDoc = ss.getSheetByName('DOCENTES');
    if (!sheetRoles) return [];
    
    const roles = sheetRoles.getDataRange().getValues().slice(1);
    const docentes = sheetDoc.getDataRange().getValues().slice(1);
    
    return roles.map(r => {
      const idDoc = String(r[0]).trim();
      const rol = String(r[1]).trim();
      const docente = docentes.find(d => String(d[0]).trim() === idDoc);
      return {
        idDocente: idDoc,
        nombreDocente: docente ? String(docente[1]) : 'Desconocido',
        rol: rol
      };
    }).filter(r => r.idDocente && r.rol);
  } catch (e) { return []; }
}

function asignarRol(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('CONFIG_ROLES');
    if (!sheet) {
      sheet = ss.insertSheet('CONFIG_ROLES');
      sheet.appendRow(['ID_Docente', 'Rol']);
    }
    
    const dataActual = sheet.getDataRange().getValues();
    for (let i = 1; i < dataActual.length; i++) {
      if (String(dataActual[i][0]).trim() === String(data.idDocente).trim()) {
        sheet.getRange(i + 1, 2).setValue(data.rol);
        return { success: true, message: 'Rol actualizado' };
      }
    }
    
    sheet.appendRow([data.idDocente, data.rol]);
    return { success: true, message: 'Rol asignado' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function eliminarRol(idDocente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('CONFIG_ROLES');
    const rows = sheet.getDataRange().getValues();
    
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(idDocente).trim()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'No encontrado' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================================
// 18. SOLICITAR CANJE
// ============================================================================
function solicitarCanje(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetSol = ss.getSheetByName('SOLICITUDES_CANJE');
    
    if (!sheetSol) {
      sheetSol = ss.insertSheet('SOLICITUDES_CANJE');
      sheetSol.appendRow(['Fecha', 'ID_Estudiante', 'Nombre_Estudiante', 'ID_Premio', 'Nombre_Premio', 'Costo', 'Estado', 'ID_Docente_Aprobador', 'Fecha_Aprobacion']);
    }
    
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    const estudiantes = sheetEst.getDataRange().getValues();
    let nombreEstudiante = 'Desconocido';
    let saldoActual = 0;
    let idEstudianteReal = String(data.idEstudiante).trim().toUpperCase();
    
    for (let i = 1; i < estudiantes.length; i++) {
      if (String(estudiantes[i][0]).trim().toLowerCase() === String(data.idEstudiante).trim().toLowerCase()) {
        nombreEstudiante = estudiantes[i][1];
        saldoActual = Number(estudiantes[i][3]);
        idEstudianteReal = String(estudiantes[i][0]).trim().toUpperCase();
        break;
      }
    }
    
    if (saldoActual < data.costo) {
      return { success: false, error: 'Saldo insuficiente. Tienes ' + saldoActual + ' pts.' };
    }
    
    const existentes = sheetSol.getDataRange().getValues();
    for (let i = 1; i < existentes.length; i++) {
      if (String(existentes[i][1]).trim().toLowerCase() === idEstudianteReal.toLowerCase() &&
          String(existentes[i][3]) === String(data.idPremio) &&
          String(existentes[i][6]) === 'PENDIENTE') {
        return { success: false, error: 'Ya tienes una solicitud pendiente para este premio.' };
      }
    }
    
    sheetSol.appendRow([
      new Date(), idEstudianteReal, nombreEstudiante,
      data.idPremio, data.nombrePremio, data.costo,
      'PENDIENTE', '', ''
    ]);
    
    return { success: true, message: 'Solicitud enviada.' };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================================
// 19. OBTENER SOLICITUDES PENDIENTES
// ============================================================================
function obtenerSolicitudesPendientes(idDocente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSol = ss.getSheetByName('SOLICITUDES_CANJE');
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    if (!sheetSol) return [];
    
    const gradosDocente = obtenerGradosDelDocente(idDocente).grados || [];
    const rol = obtenerRolDelDocente(idDocente);
    const esCoordinadorODirector = (rol === 'coordinador' || rol === 'director');
    
    if (gradosDocente.length === 0 && !esCoordinadorODirector) return [];
    
    const estudiantes = sheetEst.getDataRange().getValues();
    const estudiantesPorGrado = {};
    for (let i = 1; i < estudiantes.length; i++) {
      estudiantesPorGrado[String(estudiantes[i][0]).trim().toUpperCase()] = String(estudiantes[i][2] || '').trim();
    }
    
    const data = sheetSol.getDataRange().getValues();
    const pendientes = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][6]) === 'PENDIENTE') {
        const idEst = String(data[i][1]).trim().toUpperCase();
        const gradoEst = estudiantesPorGrado[idEst] || '';
        
        if (esCoordinadorODirector || gradosDocente.includes(gradoEst)) {
          pendientes.push({
            fila: i + 1,
            fecha: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'dd/MM HH:mm'),
            idEstudiante: idEst,
            nombreEstudiante: String(data[i][2]),
            idPremio: String(data[i][3]),
            nombrePremio: String(data[i][4]),
            costo: Number(data[i][5]),
            gradoEstudiante: gradoEst
          });
        }
      }
    }
    return pendientes;
  } catch (e) { return []; }
}

// ============================================================================
// 20. APROBAR SOLICITUD
// ============================================================================
function aprobarSolicitudCanje(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSol = ss.getSheetByName('SOLICITUDES_CANJE');
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    const sheetTrans = ss.getSheetByName('TRANSACCIONES');
    const sheetDoc = ss.getSheetByName('DOCENTES');
    
    const fila = Number(data.fila);
    const sol = sheetSol.getRange(fila, 1, 1, 9).getValues()[0];
    const idEst = String(sol[1]).trim().toUpperCase();
    const nombreEstudiante = String(sol[2]);
    const nombrePremio = String(sol[4]);
    const costo = Number(sol[5]);
    
    const fechaEfectiva = data.fechaEfectiva || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const idDocente = data.idDocente || 'DOC-UNKNOWN';
    
    let nombreDocente = 'Profesor Guía';
    if (sheetDoc) {
      const docentes = sheetDoc.getDataRange().getValues();
      for (let i = 1; i < docentes.length; i++) {
        if (String(docentes[i][0]).trim() === String(idDocente).trim()) {
          nombreDocente = String(docentes[i][1]);
          break;
        }
      }
    }
    
    const estudiantes = sheetEst.getDataRange().getValues();
    let filaEst = -1, saldoActual = 0;
    
    for (let i = 1; i < estudiantes.length; i++) {
      if (String(estudiantes[i][0]).trim().toLowerCase() === idEst.toLowerCase()) {
        filaEst = i + 1;
        saldoActual = Number(estudiantes[i][3]);
        break;
      }
    }
    
    if (filaEst === -1) return { success: false, error: 'Estudiante no encontrado' };
    if (saldoActual < costo) return { success: false, error: 'Saldo insuficiente al aprobar' };
    
    const nuevoSaldo = saldoActual - costo;
    sheetEst.getRange(filaEst, 4).setValue(nuevoSaldo);
    
    sheetTrans.appendRow([
      new Date(), idEst, idDocente,
      'CANJE: ' + nombrePremio, -costo,
      'Aprobado. Efectivo desde: ' + fechaEfectiva, 'Marketplace'
    ]);
    
    sheetSol.getRange(fila, 7).setValue('APROBADA');
    sheetSol.getRange(fila, 8).setValue(idDocente);
    sheetSol.getRange(fila, 9).setValue(new Date());
    
    try {
      const mensaje = `Tu profesor guía ${nombreDocente} aprobó tu solicitud de "${nombrePremio}". Se descontaron ${costo} pts (nuevo saldo: ${nuevoSaldo}). Efectivo desde: ${fechaEfectiva}.`;
      crearNotificacionEstudiante(idEst, '✅ Canje aprobado', mensaje);
    } catch (notifErr) {
      console.error('Error creando notificación:', notifErr);
    }
    
    return { success: true, nuevoSaldo: nuevoSaldo, notificacionCreada: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 21. RECHAZAR SOLICITUD
// ============================================================================
function rechazarSolicitudCanje(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSol = ss.getSheetByName('SOLICITUDES_CANJE');
    const sheetDoc = ss.getSheetByName('DOCENTES');
    
    const fila = Number(data.fila);
    const motivo = data.motivo || 'No especificado';
    const idDocente = data.idDocente || 'DOC-UNKNOWN';
    
    const sol = sheetSol.getRange(fila, 1, 1, 9).getValues()[0];
    const idEst = String(sol[1]).trim().toUpperCase();
    const nombreEstudiante = String(sol[2]);
    const nombrePremio = String(sol[4]);
    const costo = Number(sol[5]);
    
    let nombreDocente = 'Profesor Guía';
    if (sheetDoc) {
      const docentes = sheetDoc.getDataRange().getValues();
      for (let i = 1; i < docentes.length; i++) {
        if (String(docentes[i][0]).trim() === String(idDocente).trim()) {
          nombreDocente = String(docentes[i][1]);
          break;
        }
      }
    }
    
    sheetSol.getRange(fila, 7).setValue('RECHAZADA');
    sheetSol.getRange(fila, 8).setValue(idDocente);
    sheetSol.getRange(fila, 9).setValue(new Date());
    
    try {
      const mensaje = `Tu profesor guía ${nombreDocente} revisó tu solicitud de "${nombrePremio}". NO fue aprobada. Motivo: ${motivo}. Tus puntos no fueron descontados.`;
      crearNotificacionEstudiante(idEst, '❌ Canje rechazado', mensaje);
    } catch (notifErr) {
      console.error('Error creando notificación:', notifErr);
    }
    
    return { success: true, notificacionCreada: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================================
// 22. DASHBOARD FILTRABLE
// ============================================================================
function obtenerDashboardFiltrable(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTrans = ss.getSheetByName('TRANSACCIONES');
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    const sheetDoc = ss.getSheetByName('DOCENTES');
    if (!sheetTrans || !sheetEst) return { success: false, error: 'BD no configurada' };
    
    const estudiantes = sheetEst.getDataRange().getValues();
    const estudiantesMap = {};
    
    for (let i = 1; i < estudiantes.length; i++) {
      const id = String(estudiantes[i][0]).trim().toUpperCase();
      if (!id) continue;
      const grado = String(estudiantes[i][2] || '');
      const saldo = Number(estudiantes[i][3] || 0);
      estudiantesMap[id.toLowerCase()] = {
        id: id,
        nombre: String(estudiantes[i][1] || ''),
        grado: grado,
        saldo: saldo
      };
    }
    
    const docentes = sheetDoc ? sheetDoc.getDataRange().getValues() : [];
    const docentesMap = {};
    for (let i = 1; i < docentes.length; i++) {
      docentesMap[String(docentes[i][0]).trim()] = String(docentes[i][1] || '');
    }
    
    const filtroDocente = filtros.filtroDocente || null;
    const filtroEstudiante = filtros.filtroEstudiante || null;
    const filtroGrado = filtros.filtroGrado || null;
    
    let fechaDesde = null;
    let fechaHasta = null;
    
    if (filtros.fechaDesde) {
      const [y, m, d] = filtros.fechaDesde.split('-').map(Number);
      fechaDesde = new Date(y, m - 1, d, 0, 0, 0, 0);
    }
    
    if (filtros.fechaHasta) {
      const [y, m, d] = filtros.fechaHasta.split('-').map(Number);
      fechaHasta = new Date(y, m - 1, d, 23, 59, 59, 999);
    }
    
    let totalEstudiantesFiltrados = 0;
    let saldoTotalFiltrado = 0;
    for (const key in estudiantesMap) {
      const est = estudiantesMap[key];
      if (filtroGrado && est.grado !== filtroGrado) continue;
      if (filtroEstudiante && est.id !== filtroEstudiante) continue;
      
      totalEstudiantesFiltrados++;
      saldoTotalFiltrado += est.saldo;
    }
    
    const transacciones = sheetTrans.getDataRange().getValues();
    
    const incidenciasPorTipo = {};
    const incidenciasPorGrado = {};
    const incidenciasPorEstudiante = {};
    const incidenciasPorDocente = {};
    let totalIncidencias = 0;
    let totalCanjes = 0;
    
    for (let i = 1; i < transacciones.length; i++) {
      const row = transacciones[i];
      const fecha = new Date(row[0]);
      const idEst = String(row[1]).trim().toUpperCase();
      const idDoc = String(row[2] || '').trim();
      const tipo = String(row[3] || '');
      const puntos = Number(row[4]);
      
      if (fechaDesde && fecha < fechaDesde) continue;
      if (fechaHasta && fecha > fechaHasta) continue;
      if (filtroDocente && idDoc !== filtroDocente) continue;
      if (filtroEstudiante && idEst !== filtroEstudiante) continue;
      
      const estInfo = estudiantesMap[idEst.toLowerCase()] || { id: idEst, nombre: 'Desconocido', grado: '' };
      if (filtroGrado && estInfo.grado !== filtroGrado) continue;
      
      const esCanje = tipo.includes('CANJE');
      const esBono = tipo.includes('BONO');
      const esIncidencia = !esCanje && !esBono && puntos < 0;
      
      if (esIncidencia) {
        totalIncidencias++;
        incidenciasPorTipo[tipo] = (incidenciasPorTipo[tipo] || 0) + 1;
        incidenciasPorGrado[estInfo.grado] = (incidenciasPorGrado[estInfo.grado] || 0) + 1;
        incidenciasPorEstudiante[idEst] = (incidenciasPorEstudiante[idEst] || 0) + 1;
        incidenciasPorDocente[idDoc] = (incidenciasPorDocente[idDoc] || 0) + 1;
      } else if (esCanje) {
        totalCanjes++;
      }
    }
    
    const topIncidencias = Object.entries(incidenciasPorTipo)
      .map(([tipo, count]) => ({ tipo, count }))
      .sort((a, b) => b.count - a.count).slice(0, 5);
    
    const topGrados = Object.entries(incidenciasPorGrado)
      .map(([grado, count]) => ({ grado, count }))
      .sort((a, b) => b.count - a.count).slice(0, 5);
    
    const topEstudiantes = Object.entries(incidenciasPorEstudiante)
      .map(([id, count]) => {
        const est = estudiantesMap[id.toLowerCase()];
        return {
          id,
          nombre: est ? est.nombre : 'Desconocido',
          grado: est ? est.grado : '',
          count
        };
      })
      .sort((a, b) => b.count - a.count).slice(0, 10);
    
    const topDocentes = Object.entries(incidenciasPorDocente)
      .map(([id, count]) => ({
        id,
        nombre: docentesMap[id] || 'Sistema',
        count
      }))
      .sort((a, b) => b.count - a.count).slice(0, 5);
    
    const saldoPromedio = totalEstudiantesFiltrados > 0 
      ? Math.round(saldoTotalFiltrado / totalEstudiantesFiltrados) 
      : 0;
    
    return {
      success: true,
      metricas: {
        totalIncidencias,
        totalCanjes,
        totalEstudiantes: totalEstudiantesFiltrados,
        saldoPromedio: saldoPromedio
      },
      topIncidencias,
      topGrados,
      topEstudiantes,
      topDocentes
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================================
// 23. FILTROS DISPONIBLES PARA DASHBOARD
// ============================================================================
function obtenerFiltrosDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    const sheetDoc = ss.getSheetByName('DOCENTES');
    
    const estudiantes = sheetEst ? sheetEst.getDataRange().getValues().slice(1).map(r => ({
      id: String(r[0]).trim().toUpperCase(),
      nombre: String(r[1]).trim(),
      grado: String(r[2] || '').trim()
    })).filter(e => e.id) : [];
    
    const docentes = sheetDoc ? sheetDoc.getDataRange().getValues().slice(1).map(r => ({
      id: String(r[0]).trim(),
      nombre: String(r[1]).trim()
    })).filter(d => d.id) : [];
    
    const gradosSet = new Set();
    estudiantes.forEach(e => { if (e.grado) gradosSet.add(e.grado); });
    const grados = Array.from(gradosSet).sort();
    
    return { success: true, estudiantes, docentes, grados };
  } catch (e) {
    return { success: false, error: e.toString(), estudiantes: [], docentes: [], grados: [] };
  }
}

// ============================================================================
// 24. DASHBOARD POR MES (Admin)
// ============================================================================
function obtenerDashboard(filtros) {
  try {
    const mesActual = filtros.mes || (new Date().getMonth() + 1);
    const anioActual = filtros.anio || new Date().getFullYear();
    
    const fechaDesde = new Date(anioActual, mesActual - 1, 1, 0, 0, 0, 0);
    const fechaHasta = new Date(anioActual, mesActual, 0, 23, 59, 59, 999);
    
    const resultado = obtenerDashboardFiltrable({
      filtroDocente: null,
      filtroEstudiante: null,
      filtroGrado: null,
      fechaDesde: Utilities.formatDate(fechaDesde, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      fechaHasta: Utilities.formatDate(fechaHasta, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    });
    
    if (!resultado.success) return resultado;
    
    return {
      success: true,
      periodo: mesActual + '/' + anioActual,
      metricas: {
        totalIncidencias: resultado.metricas.totalIncidencias,
        totalCanjes: resultado.metricas.totalCanjes,
        totalEstudiantes: resultado.metricas.totalEstudiantes,
        saldoPromedio: resultado.metricas.saldoPromedio
      },
      topIncidencias: resultado.topIncidencias,
      topGrados: resultado.topGrados,
      topEstudiantes: resultado.topEstudiantes,
      topDocentes: resultado.topDocentes
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================================
// 25. BUZÓN DE SUGERENCIAS
// ============================================================================
function enviarSugerencia(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('SUGERENCIAS');
    if (!sheet) {
      sheet = ss.insertSheet('SUGERENCIAS');
      sheet.appendRow(['Fecha', 'Autor', 'Rol_Autor', 'Tipo', 'Mensaje', 'Estado']);
    }
    
    sheet.appendRow([
      new Date(),
      data.autor || 'Anónimo',
      data.rolAutor || 'estudiante',
      data.tipo || 'Sugerencia',
      data.mensaje || '',
      'NUEVO'
    ]);
    
    return { success: true, message: 'Sugerencia enviada. ¡Gracias!' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function obtenerSugerencias() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('SUGERENCIAS');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    return data.slice(1).map(r => ({
      fecha: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      autor: String(r[1] || 'Anónimo'),
      rolAutor: String(r[2] || ''),
      tipo: String(r[3] || ''),
      mensaje: String(r[4] || ''),
      estado: String(r[5] || 'NUEVO')
    })).reverse().slice(0, 100);
  } catch (e) {
    return [];
  }
}

// ============================================================================
// 26. NOTIFICACIONES GLOBALES
// ============================================================================
function obtenerNotificacionesActivas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('NOTIFICACIONES');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const hoy = new Date();
    const activas = [];
    
    for (let i = 1; i < data.length; i++) {
      const fechaExp = data[i][3] ? new Date(data[i][3]) : null;
      if (!fechaExp || fechaExp >= hoy) {
        activas.push({
          titulo: String(data[i][1] || ''),
          mensaje: String(data[i][2] || '')
        });
      }
    }
    return activas;
  } catch (e) { return []; }
}

// ============================================================================
// 27. TIENDA
// ============================================================================
function obtenerProductosTienda() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('TIENDA_PRODUCTOS');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    return data.slice(1).map(r => ({
      id: String(r[0] || ''),
      nombre: String(r[1] || ''),
      costo: Number(r[2] || 0),
      descripcion: String(r[3] || ''),
      foto: String(r[4] || ''),
      stock: Number(r[5] || 0)
    })).filter(p => p.id && p.nombre);
  } catch (e) { return []; }
}

// ============================================================================
// 28. FALTAS DE 3ER NIVEL
// ============================================================================
function obtenerFaltasTercerNivel(idDocente) {
  try {
    const rol = obtenerRolDelDocente(idDocente);
    if (rol !== 'coordinador' && rol !== 'director') {
      return { success: false, error: 'Solo coordinador o director' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTrans = ss.getSheetByName('TRANSACCIONES');
    const sheetEst = ss.getSheetByName('ESTUDIANTES');
    
    const estudiantes = sheetEst.getDataRange().getValues();
    const estudiantesMap = {};
    for (let i = 1; i < estudiantes.length; i++) {
      const id = String(estudiantes[i][0]).trim().toUpperCase();
      estudiantesMap[id.toLowerCase()] = {
        nombre: String(estudiantes[i][1] || ''),
        grado: String(estudiantes[i][2] || '')
      };
    }
    
    const transacciones = sheetTrans.getDataRange().getValues();
    const faltas = [];
    const hace3Meses = new Date(new Date().getTime() - (90 * 24 * 60 * 60 * 1000));
    
    for (let i = transacciones.length - 1; i >= 1; i--) {
      const row = transacciones[i];
      const fecha = new Date(row[0]);
      const categoria = String(row[6] || '');
      
      if (categoria === 'Faltas Graves' && fecha >= hace3Meses) {
        const idEst = String(row[1]).trim().toUpperCase();
        const estInfo = estudiantesMap[idEst.toLowerCase()] || { nombre: 'Desconocido', grado: '' };
        faltas.push({
          fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
          idEstudiante: idEst,
          nombreEstudiante: estInfo.nombre,
          grado: estInfo.grado,
          puntos: Number(row[4]),
          nota: String(row[5] || '')
        });
      }
    }
    
    return { success: true, faltas: faltas };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================================
// 29. SISTEMA DE NOTIFICACIONES PARA ESTUDIANTES
// ============================================================================
function crearNotificacionEstudiante(idEstudiante, titulo, mensaje) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('NOTIFICACIONES_ESTUDIANTE');
    if (!sheet) {
      sheet = ss.insertSheet('NOTIFICACIONES_ESTUDIANTE');
      sheet.appendRow(['Timestamp', 'ID_Estudiante', 'Titulo', 'Mensaje', 'Leida']);
    }
    
    sheet.appendRow([
      new Date(),
      String(idEstudiante).trim().toUpperCase(),
      titulo,
      mensaje,
      'NO'
    ]);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function obtenerNotificacionesEstudiante(idEstudiante) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('NOTIFICACIONES_ESTUDIANTE');
    if (!sheet) return { success: true, notificaciones: [], noLeidas: 0 };
    
    const busqueda = String(idEstudiante).trim().toLowerCase();
    const data = sheet.getDataRange().getValues();
    const notificaciones = [];
    let noLeidas = 0;
    
    for (let i = data.length - 1; i >= 1; i--) {
      const idEst = String(data[i][1]).trim().toLowerCase();
      if (idEst !== busqueda) continue;
      
      const leida = String(data[i][4] || 'NO').toUpperCase() === 'SI';
      if (!leida) noLeidas++;
      
      notificaciones.push({
        fila: i + 1,
        fecha: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'dd/MM HH:mm'),
        titulo: String(data[i][2] || ''),
        mensaje: String(data[i][3] || ''),
        leida: leida
      });
      
      if (notificaciones.length >= 30) break;
    }
    
    return { success: true, notificaciones: notificaciones, noLeidas: noLeidas };
  } catch (e) {
    return { success: false, error: e.toString(), notificaciones: [], noLeidas: 0 };
  }
}

function marcarNotificacionLeida(fila) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('NOTIFICACIONES_ESTUDIANTE');
    if (!sheet) return { success: false, error: 'Hoja no encontrada' };
    
    sheet.getRange(Number(fila), 5).setValue('SI');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================================
// 30. NORMALIZAR IDs EN TODAS LAS HOJAS
// ============================================================================
function normalizarIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalCambios = 0;
  
  const hojasConID = [
    { hoja: 'ESTUDIANTES', columna: 1 },
    { hoja: 'TRANSACCIONES', columna: 2 },
    { hoja: 'SOLICITUDES_CANJE', columna: 2 },
    { hoja: 'NOTIFICACIONES_ESTUDIANTE', columna: 2 }
  ];
  
  hojasConID.forEach(({ hoja, columna }) => {
    const sheet = ss.getSheetByName(hoja);
    if (!sheet) {
      Logger.log('⚠️ Hoja no encontrada: ' + hoja);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    let cambios = 0;
    
    for (let i = 1; i < data.length; i++) {
      const id = String(data[i][columna - 1] || '').trim();
      if (id && id !== id.toUpperCase()) {
        sheet.getRange(i + 1, columna).setValue(id.toUpperCase());
        cambios++;
      }
    }
    
    Logger.log('✅ ' + hoja + ': ' + cambios + ' IDs normalizados');
    totalCambios += cambios;
  });
  
  Logger.log('🎉 TOTAL: ' + totalCambios + ' IDs normalizados');
}

// ============================================================================
// 31. RESTAURAR FALTA DE 3ER NIVEL (si se borró por accidente)
// ============================================================================
function restaurarFaltaTercerNivel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCond = ss.getSheetByName('CONFIG_CONDUCTAS');
  
  if (!sheetCond) {
    Logger.log('❌ No existe la hoja CONFIG_CONDUCTAS');
    return;
  }
  
  const data = sheetCond.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const categoria = String(data[i][3] || '').trim();
    if (categoria === 'Faltas Graves') {
      Logger.log('✅ Ya existe la fila de Faltas Graves en la fila ' + (i + 1));
      Logger.log('   Nombre: ' + data[i][1] + ', Puntos: ' + data[i][2]);
      return;
    }
  }
  
  sheetCond.appendRow(['COND06', 'Falta de 3er Nivel', -9999, 'Faltas Graves']);
  Logger.log('✅ Fila de Falta de 3er Nivel agregada exitosamente');
}

// ============================================================================
// 32. DIAGNÓSTICO
// ============================================================================
function diagnosticarEstudiantes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEst = ss.getSheetByName('ESTUDIANTES');
  if (!sheetEst) return;
  const datos = sheetEst.getDataRange().getValues();
  Logger.log('Total filas: ' + datos.length);
  for (let i = 0; i < datos.length; i++) {
    Logger.log('Fila ' + i + ': A=[' + datos[i][0] + '] B=[' + datos[i][1] + ']');
  }
}
