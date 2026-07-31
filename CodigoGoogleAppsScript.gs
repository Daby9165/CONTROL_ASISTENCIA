/**
 * ============================================================
 *  CONTROL DE INGRESO / SALIDA Y HORAS EXTRAS
 *  Backend en Google Apps Script — usa Google Sheets como BD
 * ============================================================
 *
 * REGLAS DE NEGOCIO:
 *  - Solo se marca ENTRADA y SALIDA (no se marca almuerzo). Al calcular las
 *    horas de un día Lunes-Viernes que no sea feriado, el sistema descuenta
 *    automáticamente 1 hora de almuerzo (franja 15:00-16:00) del total
 *    trabajado.
 *  - Lunes a Viernes (no feriado): horas trabajadas = (Salida - Entrada) - 1h almuerzo.
 *      -> hasta 8 horas normales + el resto como horas extra normales.
 *  - Sábado, o cualquier día del calendario de feriados de Puerto Real (Cádiz,
 *    hoja "Feriados"): NO se descuenta almuerzo. TODAS las horas trabajadas
 *    ese día son horas extraordinarias.
 *  - Cada trabajador accede SOLO a su propio registro ingresando su número
 *    de fotocheck (no ve datos de otros trabajadores).
 *  - Existe una CLAVE MAESTRA (ver constante CLAVE_MAESTRA más abajo,
 *    CÁMBIALA) que da acceso a un modo administrador para ver el reporte
 *    diario, semanal o mensual de TODO el personal, y para editar los
 *    datos de cualquier trabajador.
 *
 * INSTALACIÓN:
 *  1. Crea una Hoja de Cálculo de Google nueva (Google Sheets).
 *  2. Menú Extensiones > Apps Script.
 *  3. Borra el contenido por defecto y pega TODO este archivo.
 *  4. Cambia el valor de CLAVE_MAESTRA más abajo por una clave propia.
 *  5. En el desplegable de funciones (arriba), elige "inicializarHojas"
 *     y presiona ▶ Ejecutar. Autoriza los permisos que pida Google.
 *     Esto crea las hojas "Empleados", "Registros" y "Feriados".
 *
 *     SI YA TENÍAS UNA VERSIÓN ANTERIOR con columnas de almuerzo: renombra
 *     tu hoja "Registros" actual a "Registros_OLD" (para no perder el
 *     historial) antes de ejecutar inicializarHojas(), así se crea una
 *     hoja "Registros" nueva con la estructura correcta.
 *  6. Agrega tu personal desde la app con el formulario "Registrar
 *     trabajador", o directamente en la hoja "Empleados".
 *  7. Cada año, actualiza la hoja "Feriados" con el calendario del año
 *     siguiente (Fecha en columna A, Descripción en columna B).
 *  8. Menú Implementar > Nueva implementación:
 *       - Tipo: "Aplicación web"
 *       - Ejecutar como: Yo (tu cuenta)
 *       - Quién tiene acceso: "Cualquier usuario"
 *  9. Copia la URL de la aplicación web y pégala en index.html (frontend).
 * 10. Cada vez que modifiques este código, crea una NUEVA versión de la
 *     implementación (Implementar > Administrar implementaciones > Editar
 *     > Nueva versión).
 * ============================================================
 */

const SHEET_EMPLEADOS = 'Empleados';
const SHEET_REGISTROS = 'Registros';
const SHEET_FERIADOS = 'Feriados';

const JORNADA_NORMAL_HORAS = 8;  // Horas normales L-V
const HORAS_ALMUERZO = 1;        // Se descuenta automáticamente en L-V no feriado
const PROYECTOS_VALIDOS = ['Dolwin 4', 'Borwin 4'];

// ⚠️ CAMBIA ESTA CLAVE antes de publicar la app. Da acceso al modo administrador.
const CLAVE_MAESTRA = 'CAMBIAR-ESTA-CLAVE-2026';

// Calendario laboral 2026 de Puerto Real (Cádiz): nacionales + autonómico
// Andalucía + locales. Fuente: BOE y Junta de Andalucía. Actualizar cada año.
const FERIADOS_PUERTO_REAL_2026 = [
  ['2026-01-01', 'Año Nuevo'],
  ['2026-01-06', 'Epifanía del Señor'],
  ['2026-02-11', 'Fiesta local (Puerto Real)'],
  ['2026-02-28', 'Día de Andalucía'],
  ['2026-04-02', 'Jueves Santo'],
  ['2026-04-03', 'Viernes Santo'],
  ['2026-05-01', 'Fiesta del Trabajo'],
  ['2026-06-08', 'Fiesta local (Puerto Real)'],
  ['2026-08-15', 'Asunción de la Virgen'],
  ['2026-10-12', 'Fiesta Nacional de España'],
  ['2026-11-02', 'Día siguiente a Todos los Santos'],
  ['2026-12-07', 'Lunes siguiente a la Constitución'],
  ['2026-12-08', 'Inmaculada Concepción'],
  ['2026-12-25', 'Natividad del Señor']
];

// ---------- INSTALACIÓN ----------

function inicializarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let empleados = ss.getSheetByName(SHEET_EMPLEADOS);
  if (!empleados) {
    empleados = ss.insertSheet(SHEET_EMPLEADOS);
    empleados.appendRow(['ID', 'Nombre Completo', 'Proyecto', 'Fotocheck', 'Activo']);
    empleados.appendRow(['1001', 'Empleado de ejemplo', 'Dolwin 4', 'FC-0001', 'SI']);
    empleados.setFrozenRows(1);
  }

  let registros = ss.getSheetByName(SHEET_REGISTROS);
  if (!registros) {
    registros = ss.insertSheet(SHEET_REGISTROS);
    registros.appendRow([
      'Fecha', 'Dia Semana', 'ID Empleado', 'Nombre',
      'Entrada', 'Salida',
      'Horas Normales', 'Horas Extra Normales (L-V)', 'Horas Extraordinarias (Sab/Feriado)',
      'Total Horas', 'Observaciones'
    ]);
    registros.setFrozenRows(1);
  }

  let feriados = ss.getSheetByName(SHEET_FERIADOS);
  if (!feriados) {
    feriados = ss.insertSheet(SHEET_FERIADOS);
    feriados.appendRow(['Fecha', 'Descripcion']);
    FERIADOS_PUERTO_REAL_2026.forEach(f => feriados.appendRow(f));
    feriados.setFrozenRows(1);
    feriados.autoResizeColumns(1, 2);
  }

  SpreadsheetApp.flush();
  Browser.msgBox('Listo. Hojas "Empleados", "Registros" y "Feriados" creadas/verificadas.');
}

// ---------- ENDPOINTS HTTP ----------

function doGet(e) {
  const accion = e.parameter.accion;
  try {
    if (accion === 'login') return respuesta(loginPorFotocheck(e.parameter.fotocheck));
    if (accion === 'proyectos') return respuesta(PROYECTOS_VALIDOS);
    if (accion === 'estadoHoy') return respuesta(estadoHoy(e.parameter.idEmpleado));
    if (accion === 'historial') return respuesta(historial(e.parameter.idEmpleado, e.parameter.desde, e.parameter.hasta));
    if (accion === 'empleadosAdmin') return respuesta(empleadosAdmin(e.parameter.clave));
    if (accion === 'reporte') return respuesta(reporte(e.parameter.clave, e.parameter.periodo, e.parameter.fecha));
    return respuesta({ error: 'Accion no reconocida' });
  } catch (err) {
    return respuesta({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.accion === 'marcar') {
      return respuesta(marcar(body.idEmpleado, body.tipo));
    }
    if (body.accion === 'registrar') {
      return respuesta(registrarEmpleado(body.nombreCompleto, body.proyecto, body.fotocheck));
    }
    if (body.accion === 'editarEmpleado') {
      return respuesta(editarEmpleado(body.clave, body.id, body.nombreCompleto, body.proyecto, body.fotocheck, body.activo));
    }
    return respuesta({ error: 'Accion no reconocida' });
  } catch (err) {
    return respuesta({ error: err.message });
  }
}

function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function verificarClave(clave) {
  return clave && String(clave) === CLAVE_MAESTRA;
}

// ---------- EMPLEADOS ----------

function getSheetEmpleados() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMPLEADOS);
}

function filaEmpleadoPorId(idEmpleado) {
  const sh = getSheetEmpleados();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idEmpleado)) return i + 1;
  }
  return -1;
}

function loginPorFotocheck(fotocheck) {
  fotocheck = (fotocheck || '').trim();
  if (!fotocheck) return { error: 'Ingrese su numero de fotocheck.' };

  const sh = getSheetEmpleados();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim().toUpperCase() === fotocheck.toUpperCase() &&
        String(data[i][4]).toUpperCase() === 'SI') {
      return {
        ok: true,
        id: String(data[i][0]),
        nombre: data[i][1],
        proyecto: data[i][2],
        fotocheck: data[i][3]
      };
    }
  }
  return { error: 'No se encontro un trabajador activo con ese fotocheck.' };
}

function nombreEmpleado(idEmpleado) {
  const sh = getSheetEmpleados();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idEmpleado)) return data[i][1];
  }
  return 'Desconocido';
}

function registrarEmpleado(nombreCompleto, proyecto, fotocheck) {
  nombreCompleto = (nombreCompleto || '').trim();
  fotocheck = (fotocheck || '').trim();

  if (!nombreCompleto) return { error: 'El nombre completo es obligatorio.' };
  if (PROYECTOS_VALIDOS.indexOf(proyecto) === -1) return { error: 'Debe elegir Dolwin 4 o Borwin 4.' };
  if (!fotocheck) return { error: 'El numero de fotocheck es obligatorio.' };

  const sh = getSheetEmpleados();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim().toUpperCase() === fotocheck.toUpperCase()) {
      return { error: 'Ya existe un trabajador con ese numero de fotocheck.' };
    }
  }

  let maxId = 1000;
  for (let i = 1; i < data.length; i++) {
    const n = parseInt(data[i][0], 10);
    if (!isNaN(n) && n > maxId) maxId = n;
  }
  const nuevoId = String(maxId + 1);

  sh.appendRow([nuevoId, nombreCompleto, proyecto, fotocheck, 'SI']);

  return { ok: true, mensaje: 'Trabajador registrado correctamente.', id: nuevoId };
}

// Requiere clave maestra: lista TODOS los trabajadores (activos e inactivos)
function empleadosAdmin(clave) {
  if (!verificarClave(clave)) return { error: 'Clave maestra incorrecta.' };
  const sh = getSheetEmpleados();
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    out.push({
      id: String(data[i][0]),
      nombre: data[i][1],
      proyecto: data[i][2],
      fotocheck: data[i][3],
      activo: String(data[i][4]).toUpperCase() === 'SI'
    });
  }
  return out;
}

// Requiere clave maestra: edita nombre, proyecto, fotocheck y estado activo
function editarEmpleado(clave, id, nombreCompleto, proyecto, fotocheck, activo) {
  if (!verificarClave(clave)) return { error: 'Clave maestra incorrecta.' };

  nombreCompleto = (nombreCompleto || '').trim();
  fotocheck = (fotocheck || '').trim();
  if (!nombreCompleto) return { error: 'El nombre completo es obligatorio.' };
  if (PROYECTOS_VALIDOS.indexOf(proyecto) === -1) return { error: 'Debe elegir Dolwin 4 o Borwin 4.' };
  if (!fotocheck) return { error: 'El numero de fotocheck es obligatorio.' };

  const sh = getSheetEmpleados();
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(id) && String(data[i][3]).trim().toUpperCase() === fotocheck.toUpperCase()) {
      return { error: 'Ese numero de fotocheck ya lo tiene otro trabajador.' };
    }
  }

  const fila = filaEmpleadoPorId(id);
  if (fila === -1) return { error: 'Trabajador no encontrado.' };

  sh.getRange(fila, 2, 1, 4).setValues([[nombreCompleto, proyecto, fotocheck, activo ? 'SI' : 'NO']]);
  return { ok: true, mensaje: 'Datos actualizados correctamente.' };
}

// ---------- REGISTROS (marcaje) ----------

function getSheetRegistros() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REGISTROS);
}

function formatearFecha(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function diaSemanaTexto(d) {
  const dias = ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'];
  return dias[d.getDay()];
}

function buscarFilaHoy(idEmpleado) {
  const sh = getSheetRegistros();
  const data = sh.getDataRange().getValues();
  const hoy = formatearFecha(new Date());
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && formatearFecha(new Date(data[i][0])) === hoy && String(data[i][2]) === String(idEmpleado)) {
      return i + 1;
    }
  }
  return -1;
}

function estadoHoy(idEmpleado) {
  const sh = getSheetRegistros();
  const fila = buscarFilaHoy(idEmpleado);
  if (fila === -1) return { marcado: [] };

  const row = sh.getRange(fila, 1, 1, 10).getValues()[0];
  const marcado = [];
  if (row[4]) marcado.push('entrada');
  if (row[5]) marcado.push('salida');

  return {
    marcado: marcado,
    entrada: fmtHora(row[4]),
    salida: fmtHora(row[5]),
    horasNormales: row[6] || 0,
    horasExtraNormales: row[7] || 0,
    horasExtraordinarias: row[8] || 0,
    total: row[9] || 0
  };
}

function marcar(idEmpleado, tipo) {
  if (!idEmpleado) return { error: 'Sesion invalida, vuelva a ingresar su fotocheck.' };
  const sh = getSheetRegistros();
  const ahora = new Date();
  const esDomingo = ahora.getDay() === 0;

  if (esDomingo) return { error: 'No se registran marcas los domingos.' };

  let fila = buscarFilaHoy(idEmpleado);

  if (tipo === 'entrada') {
    if (fila !== -1) return { error: 'Ya registro su entrada hoy.' };
    sh.appendRow([
      new Date(formatearFecha(ahora)), diaSemanaTexto(ahora), String(idEmpleado), nombreEmpleado(idEmpleado),
      ahora, '', '', '', '', '', ''
    ]);
    return { ok: true, mensaje: 'Entrada registrada a las ' + fmtHora(ahora) };
  }

  if (tipo === 'salida') {
    if (fila === -1) return { error: 'Debe registrar la entrada primero.' };
    const valorActual = sh.getRange(fila, 6).getValue();
    if (valorActual) return { error: 'La salida ya fue registrada hoy.' };
    sh.getRange(fila, 6).setValue(ahora);
    calcularHorasFila(fila);
    return { ok: true, mensaje: 'Salida registrada a las ' + fmtHora(ahora) };
  }

  return { error: 'Tipo de marca no valido.' };
}

function calcularHorasFila(fila) {
  const sh = getSheetRegistros();
  const row = sh.getRange(fila, 1, 1, 6).getValues()[0];
  const fecha = new Date(row[0]);
  const diaSemana = fecha.getDay(); // 0 domingo ... 6 sabado

  const entrada = row[4] ? new Date(row[4]) : null;
  const salida = row[5] ? new Date(row[5]) : null;
  if (!entrada || !salida) return;

  const feriado = esFeriado(fecha);
  let horasNormales = 0, horasExtraNormales = 0, horasExtraordinarias = 0;
  const horasBrutas = (salida - entrada) / 3600000;

  if (diaSemana === 6 || feriado) {
    // Sabado o feriado: sin descuento de almuerzo, todo es extraordinario
    horasExtraordinarias = round2(Math.max(horasBrutas, 0));
  } else {
    // Lunes a viernes: se descuenta 1 hora de almuerzo automaticamente
    const horasTrabajadas = Math.max(horasBrutas - HORAS_ALMUERZO, 0);
    horasNormales = round2(Math.min(horasTrabajadas, JORNADA_NORMAL_HORAS));
    horasExtraNormales = round2(Math.max(horasTrabajadas - JORNADA_NORMAL_HORAS, 0));
  }

  sh.getRange(fila, 7).setValue(horasNormales);
  sh.getRange(fila, 8).setValue(horasExtraNormales);
  sh.getRange(fila, 9).setValue(horasExtraordinarias);
  sh.getRange(fila, 10).setValue(round2(horasNormales + horasExtraNormales + horasExtraordinarias));
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function fmtHora(v) {
  if (!v) return '';
  return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'HH:mm');
}

function esFeriado(fecha) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FERIADOS);
  if (!sh) return false;
  const data = sh.getDataRange().getValues();
  const fechaTexto = formatearFecha(fecha);
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const f = data[i][0] instanceof Date ? formatearFecha(data[i][0]) : String(data[i][0]).slice(0, 10);
    if (f === fechaTexto) return true;
  }
  return false;
}

// ---------- CONSULTAS ----------

// Historial de UN trabajador (usado por el propio trabajador tras su login)
function historial(idEmpleado, desde, hasta) {
  const sh = getSheetRegistros();
  const data = sh.getDataRange().getValues();
  const out = [];
  const dDesde = desde ? new Date(desde) : null;
  const dHasta = hasta ? new Date(hasta) : null;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[2]) !== String(idEmpleado)) continue;
    const fecha = new Date(row[0]);
    if (dDesde && fecha < dDesde) continue;
    if (dHasta && fecha > dHasta) continue;
    out.push({
      fecha: formatearFecha(fecha),
      diaSemana: row[1],
      entrada: fmtHora(row[4]),
      salida: fmtHora(row[5]),
      horasNormales: row[6],
      horasExtraNormales: row[7],
      horasExtraordinarias: row[8],
      total: row[9]
    });
  }
  out.sort((a, b) => a.fecha < b.fecha ? 1 : -1);
  return out;
}

// Requiere clave maestra: reporte de TODO el personal por dia/semana/mes
function reporte(clave, periodo, fechaRef) {
  if (!verificarClave(clave)) return { error: 'Clave maestra incorrecta.' };

  const ref = fechaRef ? new Date(fechaRef) : new Date();
  let desde, hasta;

  if (periodo === 'semanal') {
    const diaSemana = ref.getDay(); // 0=domingo
    const offsetLunes = diaSemana === 0 ? 6 : diaSemana - 1;
    desde = new Date(ref); desde.setDate(ref.getDate() - offsetLunes);
    hasta = new Date(desde); hasta.setDate(desde.getDate() + 6);
  } else if (periodo === 'mensual') {
    desde = new Date(ref.getFullYear(), ref.getMonth(), 1);
    hasta = new Date(ref.getFullYear(), ref.getMonth() + 1, 0);
  } else {
    // diario
    desde = new Date(ref);
    hasta = new Date(ref);
  }
  desde.setHours(0, 0, 0, 0);
  hasta.setHours(23, 59, 59, 999);

  const sh = getSheetRegistros();
  const data = sh.getDataRange().getValues();
  const resumen = {};
  const detalle = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const fecha = new Date(row[0]);
    if (fecha < desde || fecha > hasta) continue;

    const id = String(row[2]);
    if (!resumen[id]) resumen[id] = { nombre: row[3], horasNormales: 0, horasExtraNormales: 0, horasExtraordinarias: 0, total: 0 };
    resumen[id].horasNormales += Number(row[6]) || 0;
    resumen[id].horasExtraNormales += Number(row[7]) || 0;
    resumen[id].horasExtraordinarias += Number(row[8]) || 0;
    resumen[id].total += Number(row[9]) || 0;

    detalle.push({
      fecha: formatearFecha(fecha),
      diaSemana: row[1],
      idEmpleado: id,
      nombre: row[3],
      entrada: fmtHora(row[4]),
      salida: fmtHora(row[5]),
      horasNormales: row[6],
      horasExtraNormales: row[7],
      horasExtraordinarias: row[8],
      total: row[9]
    });
  }

  Object.keys(resumen).forEach(id => {
    resumen[id].horasNormales = round2(resumen[id].horasNormales);
    resumen[id].horasExtraNormales = round2(resumen[id].horasExtraNormales);
    resumen[id].horasExtraordinarias = round2(resumen[id].horasExtraordinarias);
    resumen[id].total = round2(resumen[id].total);
  });

  detalle.sort((a, b) => a.fecha < b.fecha ? 1 : -1);

  return {
    desde: formatearFecha(desde),
    hasta: formatearFecha(hasta),
    resumen: resumen,
    detalle: detalle
  };
}
