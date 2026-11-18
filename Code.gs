/**
 * EL PATRÓN — MÓDULO DE FICHAJE (Control de Horarios)
 * Backend optimizado para Google Apps Script + Google Sheets.
 * Zona horaria fija: America/Argentina/Cordoba
 *
 * Incluye:
 * - Cálculo de horas por empleado y por día (Total trabajado, Horas normales y Horas extra en HH:MM).
 * - Jornada diaria configurable por empleado (8 h = 480 min por defecto).
 * - Las horas extra de cada día se suman sin compensarse con jornadas más cortas.
 * - Los turnos abiertos quedan como PENDIENTE sin afectar la suma de cerrados.
 */

const TZ = 'America/Argentina/Cordoba';
const SALT = 'EL_PATRON_FICHAJE_SALT_2026_CBA_AR_#94653f';
const SESSION_TTL_SECONDS = 21600; // 6 horas
const DEFAULT_JORNADA_HORAS = 8; // 8 horas diarias por defecto (480 min)

const SHEET_USUARIOS = 'Usuarios';
const SHEET_FICHAJES = 'Fichajes';

// Caché en memoria durante la ejecución de una petición
let _ssCache = null;
let _shUsuariosCache = null;
let _shFichajesCache = null;

function getSS_() {
  if (!_ssCache) {
    _ssCache = SpreadsheetApp.getActiveSpreadsheet();
  }
  return _ssCache;
}

/**
 * Inicializa la planilla creando las hojas "Usuarios" y "Fichajes"
 * y el usuario inicial: dueno / PIN 1234 (rol: dueño).
 * Conserva todos los registros existentes.
 */
function setup() {
  const ss = getSS_();
  if (!ss) {
    throw new Error('Este script debe estar vinculado a una Google Sheet (Extensiones > Apps Script).');
  }

  try {
    ss.setSpreadsheetTimeZone(TZ);
  } catch (err) {
    Logger.log('Aviso al configurar zona horaria: ' + err);
  }

  // 1. Hoja Usuarios: Usuario, Nombre, Rol, PIN hash, Activo, Creado, Jornada hs
  let shUsuarios = ss.getSheetByName(SHEET_USUARIOS);
  if (!shUsuarios) {
    shUsuarios = ss.insertSheet(SHEET_USUARIOS);
  }
  if (shUsuarios.getLastRow() === 0) {
    const headersU = ['Usuario', 'Nombre', 'Rol', 'PIN hash', 'Activo', 'Creado', 'Jornada hs'];
    shUsuarios.getRange(1, 1, 1, headersU.length).setValues([headersU]);
    shUsuarios.getRange(1, 1, 1, headersU.length)
      .setFontWeight('bold')
      .setBackground('#1a100a')
      .setFontColor('#d9a05f');
    shUsuarios.setFrozenRows(1);
  } else {
    // Asegurar encabezado de columna 7 ("Jornada hs") sin alterar datos existentes
    const col7Header = String(shUsuarios.getRange(1, 7).getValue() || '').trim();
    if (!col7Header) {
      shUsuarios.getRange(1, 7)
        .setValue('Jornada hs')
        .setFontWeight('bold')
        .setBackground('#1a100a')
        .setFontColor('#d9a05f');
    }
  }

  const usuariosData = shUsuarios.getDataRange().getValues();
  let existeDueno = false;
  for (let i = 1; i < usuariosData.length; i++) {
    const uRow = String(usuariosData[i][0]).trim().toLowerCase();
    if (uRow === 'dueno') {
      existeDueno = true;
      break;
    }
  }

  const ahoraStr = formatNowDateTime_();

  if (!existeDueno) {
    const pinHash = hashPin_('1234', 'dueno');
    const nextRow = shUsuarios.getLastRow() + 1;
    const rng = shUsuarios.getRange(nextRow, 1, 1, 7);
    rng.setNumberFormat('@');
    rng.setValues([[
      'dueno',
      'Dueño El Patrón',
      'dueño',
      pinHash,
      'SI',
      ahoraStr,
      String(DEFAULT_JORNADA_HORAS)
    ]]);
  }

  // 2. Hoja Fichajes: ID, Usuario, Nombre, Fecha, Ingreso, Egreso, Horas decimal, Duración hh:mm, Estado, IngresoTS, EgresoTS
  let shFichajes = ss.getSheetByName(SHEET_FICHAJES);
  if (!shFichajes) {
    shFichajes = ss.insertSheet(SHEET_FICHAJES);
  }
  if (shFichajes.getLastRow() === 0) {
    const headersF = [
      'ID',
      'Usuario',
      'Nombre',
      'Fecha',
      'Ingreso',
      'Egreso',
      'Horas decimal',
      'Duración hh:mm',
      'Estado',
      'IngresoTS',
      'EgresoTS'
    ];
    shFichajes.getRange(1, 1, 1, headersF.length).setValues([headersF]);
    shFichajes.getRange(1, 1, 1, headersF.length)
      .setFontWeight('bold')
      .setBackground('#1a100a')
      .setFontColor('#d9a05f');
    shFichajes.setFrozenRows(1);
  }

  _shUsuariosCache = shUsuarios;
  _shFichajesCache = shFichajes;
  invalidateUsersCache_();

  Logger.log('Setup completado con éxito. Usuario inicial: dueno / PIN: 1234');
  return 'Setup completado con éxito. Usuario inicial: dueno / PIN: 1234';
}

function doGet() {
  try {
    setup();
  } catch (e) {
    Logger.log('Auto-setup en doGet: ' + e);
  }
  return jsonResponse_({
    ok: true,
    app: 'El Patrón · Módulo de Fichaje',
    timezone: TZ,
    mensaje: 'El servidor de fichaje está activo y las hojas Usuarios y Fichajes están listas.'
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return errorResponse_('No se recibieron datos en la solicitud.', 'EMPTY_BODY');
    }

    let body;
    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return errorResponse_('El formato JSON enviado no es válido.', 'INVALID_JSON');
    }

    const action = String(body.action || '').trim();
    if (!action) {
      return errorResponse_('Falta especificar la acción (action).', 'MISSING_ACTION');
    }

    if (action === 'login') {
      return handleLogin_(body);
    }

    const token = String(body.token || '').trim();
    const session = getSession_(token);
    if (!session) {
      return errorResponse_('Tu sesión expiró o no es válida. Volvé a iniciar sesión.', 'UNAUTHORIZED');
    }

    const currentUser = findUserByUsernameFast_(session.usuario);
    if (!currentUser || !currentUser.activo) {
      if (token) CacheService.getScriptCache().remove('SESSION_' + token);
      return errorResponse_('Tu usuario fue deshabilitado o eliminado. Consultá con el dueño.', 'USER_DISABLED');
    }

    switch (action) {
      case 'estado':
        return handleEstado_(currentUser);

      case 'ingreso':
        return handleIngreso_(currentUser);

      case 'egreso':
        return handleEgreso_(currentUser);

      case 'misHoras':
        return handleMisHoras_(currentUser, body);

      case 'cambiarPin':
        return handleCambiarPin_(currentUser, body);

      // Acciones exclusivas del rol "dueño"
      case 'empleados':
        if (!isDueno_(currentUser)) return forbiddenResponse_();
        return handleEmpleados_();

      case 'crearEmpleado':
        if (!isDueno_(currentUser)) return forbiddenResponse_();
        return handleCrearEmpleado_(body);

      case 'cambiarJornada':
        if (!isDueno_(currentUser)) return forbiddenResponse_();
        return handleCambiarJornada_(body);

      case 'toggleEmpleado':
        if (!isDueno_(currentUser)) return forbiddenResponse_();
        return handleToggleEmpleado_(currentUser, body);

      case 'eliminarEmpleado':
        if (!isDueno_(currentUser)) return forbiddenResponse_();
        return handleEliminarEmpleado_(currentUser, body);

      case 'resetPin':
        if (!isDueno_(currentUser)) return forbiddenResponse_();
        return handleResetPin_(body);

      case 'egresoForzado':
        if (!isDueno_(currentUser)) return forbiddenResponse_();
        return handleEgresoForzado_(body);

      case 'reporte':
        if (!isDueno_(currentUser)) return forbiddenResponse_();
        return handleReporte_(body);

      default:
        return errorResponse_('Acción desconocida: ' + action, 'UNKNOWN_ACTION');
    }
  } catch (err) {
    return errorResponse_('Error interno en el servidor: ' + (err && err.message ? err.message : String(err)), 'SERVER_ERROR');
  }
}

/* ==========================================================================
 * MANEJADORES DE ACCIONES
 * ========================================================================== */

function handleLogin_(body) {
  const usuarioInput = normalizeUsername_(body.usuario);
  const pinInput = String(body.pin || '').trim();

  if (!usuarioInput) {
    return errorResponse_('Ingresá tu nombre de usuario.', 'MISSING_USER');
  }
  if (!/^\d{4,8}$/.test(pinInput)) {
    return errorResponse_('El PIN debe tener entre 4 y 8 dígitos numéricos.', 'INVALID_PIN_FORMAT');
  }

  const user = findUserByUsernameFast_(usuarioInput);
  const expectedHashUser = hashPin_(pinInput, usuarioInput);
  const expectedHashPlain = hashPinLegacy_(pinInput);
  const pinValido = user && (user.pinHash === expectedHashUser || user.pinHash === expectedHashPlain);

  if (!user || !pinValido) {
    return jsonResponse_({
      ok: false,
      error: 'Usuario o PIN incorrecto.',
      code: 'INVALID_CREDENTIALS'
    });
  }

  if (!user.activo) {
    return errorResponse_('Tu cuenta está deshabilitada. Hablá con el dueño.', 'USER_DISABLED');
  }

  const cache = CacheService.getScriptCache();
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const sessionData = {
    usuario: user.usuario,
    nombre: user.nombre,
    rol: user.rol,
    loginAt: Date.now()
  };
  cache.put('SESSION_' + token, JSON.stringify(sessionData), SESSION_TTL_SECONDS);

  const shifts = getAllShifts_();
  const estadoData = buildEstadoUsuarioFromShifts_(user, shifts);

  let empleadosInit = null;
  if (isDueno_(user)) {
    empleadosInit = buildEmpleadosPayload_(getAllUsers_(), shifts);
  }

  return jsonResponse_({
    ok: true,
    token: token,
    usuario: {
      usuario: user.usuario,
      nombre: user.nombre,
      rol: user.rol,
      jornadaHoras: user.jornadaHoras
    },
    estado: estadoData,
    empleadosInit: empleadosInit
  });
}

function handleEstado_(user) {
  const shifts = getAllShifts_();
  const estado = buildEstadoUsuarioFromShifts_(user, shifts);
  return jsonResponse_({
    ok: true,
    usuario: {
      usuario: user.usuario,
      nombre: user.nombre,
      rol: user.rol,
      jornadaHoras: user.jornadaHoras
    },
    ...estado
  });
}

function handleIngreso_(user) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return errorResponse_('El servidor está procesando otro fichaje. Intentá de nuevo en unos segundos.', 'LOCK_TIMEOUT');
  }

  try {
    const sh = getSheetFichajes_();
    const shifts = getAllShifts_();
    let openShift = null;
    for (let i = shifts.length - 1; i >= 0; i--) {
      if (shifts[i].usuario === user.usuario && shifts[i].estado === 'ABIERTO') {
        openShift = shifts[i];
        break;
      }
    }

    if (openShift) {
      return errorResponse_(
        'Ya tenés un ingreso abierto desde las ' + openShift.ingreso + ' (' + openShift.fecha + '). Tenés que marcar EGRESO antes de volver a ingresar.',
        'ALREADY_OPEN'
      );
    }

    const now = new Date();
    const nowMs = now.getTime();
    const fechaHoy = formatDateYMD_(now);
    const horaIngreso = formatTimeHM_(now);
    const id = 'F-' + Utilities.formatDate(now, TZ, 'yyyyMMdd-HHmmss') + '-' + Math.floor(1000 + Math.random() * 9000);

    const newRow = [
      id,
      user.usuario,
      user.nombre,
      fechaHoy,
      horaIngreso,
      '',
      '0',
      '00:00',
      'ABIERTO',
      String(nowMs),
      ''
    ];

    const rowIdx = sh.getLastRow() + 1;
    const rng = sh.getRange(rowIdx, 1, 1, newRow.length);
    rng.setNumberFormat('@');
    rng.setValues([newRow]);

    shifts.push({
      rowNumber: rowIdx,
      id: id,
      usuario: user.usuario,
      nombre: user.nombre,
      fecha: fechaHoy,
      ingreso: horaIngreso,
      egreso: '',
      horasDecimal: 0,
      duracion: '00:00',
      estado: 'ABIERTO',
      ingresoTS: nowMs,
      egresoTS: null
    });

    const estado = buildEstadoUsuarioFromShifts_(user, shifts);
    return jsonResponse_({
      ok: true,
      mensaje: 'Ingreso registrado a las ' + horaIngreso + ' hs.',
      ...estado
    });
  } finally {
    lock.releaseLock();
  }
}

function handleEgreso_(user) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return errorResponse_('El servidor está procesando otro fichaje. Intentá de nuevo en unos segundos.', 'LOCK_TIMEOUT');
  }

  try {
    const sh = getSheetFichajes_();
    const shifts = getAllShifts_();
    let openShift = null;
    for (let i = shifts.length - 1; i >= 0; i--) {
      if (shifts[i].usuario === user.usuario && shifts[i].estado === 'ABIERTO') {
        openShift = shifts[i];
        break;
      }
    }

    if (!openShift) {
      return errorResponse_(
        'No tenés ningún turno abierto para cerrar. Primero tenés que marcar INGRESO.',
        'NO_OPEN_SHIFT'
      );
    }

    const now = new Date();
    const egresoMs = Math.max(now.getTime(), openShift.ingresoTS + 1000);
    const horaEgreso = formatTimeHM_(new Date(egresoMs));
    const calc = calculateDurationExact_(openShift.ingreso, horaEgreso, openShift.ingresoTS, egresoMs);

    const rng = sh.getRange(openShift.rowNumber, 6, 1, 6);
    rng.setNumberFormat('@');
    rng.setValues([[
      horaEgreso,
      String(calc.horasDecimal),
      calc.duracionHHMM,
      'CERRADO',
      String(openShift.ingresoTS),
      String(egresoMs)
    ]]);

    openShift.egreso = horaEgreso;
    openShift.horasDecimal = calc.horasDecimal;
    openShift.duracion = calc.duracionHHMM;
    openShift.estado = 'CERRADO';
    openShift.egresoTS = egresoMs;

    const estado = buildEstadoUsuarioFromShifts_(user, shifts);
    return jsonResponse_({
      ok: true,
      mensaje: 'Egreso registrado a las ' + horaEgreso + ' hs. Duración del turno: ' + calc.duracionHHMM + ' hs.',
      ...estado
    });
  } finally {
    lock.releaseLock();
  }
}

function handleMisHoras_(user, body) {
  const defaults = getCurrentMonthRange_();
  const desde = normalizeDateFilter_(body.desde, defaults.desde);
  const hasta = normalizeDateFilter_(body.hasta, defaults.hasta);

  const allShifts = getAllShifts_();
  const filtered = allShifts.filter(function(item) {
    return item.usuario === user.usuario && item.fecha >= desde && item.fecha <= hasta;
  });

  const jornadaMap = {};
  jornadaMap[user.usuario] = user.jornadaHoras || DEFAULT_JORNADA_HORAS;

  const processed = computeDailyAndMonthlyHours_(filtered, jornadaMap, [user]);

  return jsonResponse_({
    ok: true,
    desde: desde,
    hasta: hasta,
    jornadaHoras: user.jornadaHoras || DEFAULT_JORNADA_HORAS,
    fichajes: processed.fichajesEnriquecidos,
    resumenDiario: processed.resumenDiario,
    totalDuracionHHMM: processed.totalesGenerales.totalTrabajadoHHMM,
    totalNormalesHHMM: processed.totalesGenerales.horasNormalesHHMM,
    totalExtrasHHMM: processed.totalesGenerales.horasExtrasHHMM,
    totalHorasDecimal: processed.totalesGenerales.horasDecimal,
    totalMinutos: processed.totalesGenerales.minutosTrabajados,
    cantidadTurnos: filtered.length
  });
}

function handleCambiarPin_(user, body) {
  const pinActual = String(body.pinActual || '').trim();
  const pinNuevo = String(body.pinNuevo || '').trim();

  if (!/^\d{4,8}$/.test(pinActual)) {
    return errorResponse_('Ingresá tu PIN actual (4 a 8 dígitos).', 'INVALID_CURRENT_PIN');
  }
  if (!/^\d{4,8}$/.test(pinNuevo)) {
    return errorResponse_('El nuevo PIN debe tener entre 4 y 8 dígitos numéricos.', 'INVALID_NEW_PIN');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const freshUser = findUserByUsername_(user.usuario);
    const hashUser = hashPin_(pinActual, user.usuario);
    const hashPlain = hashPinLegacy_(pinActual);

    if (!freshUser || (freshUser.pinHash !== hashUser && freshUser.pinHash !== hashPlain)) {
      return errorResponse_('El PIN actual que ingresaste es incorrecto.', 'WRONG_CURRENT_PIN');
    }

    const nuevoHash = hashPin_(pinNuevo, user.usuario);
    const sh = getSheetUsuarios_();
    sh.getRange(freshUser.rowNumber, 4).setNumberFormat('@').setValue(nuevoHash);
    invalidateUsersCache_();

    return jsonResponse_({
      ok: true,
      mensaje: 'Tu PIN se actualizó correctamente.'
    });
  } finally {
    lock.releaseLock();
  }
}

function handleEmpleados_() {
  const users = getAllUsers_();
  const shifts = getAllShifts_();
  const payload = buildEmpleadosPayload_(users, shifts);
  return jsonResponse_({
    ok: true,
    ...payload
  });
}

function handleCrearEmpleado_(body) {
  const nombre = String(body.nombre || '').trim();
  const usuario = normalizeUsername_(body.usuario);
  const pin = String(body.pin || '').trim();
  const rolInput = String(body.rol || 'empleado').trim().toLowerCase();
  const rol = (rolInput === 'dueño' || rolInput === 'dueno') ? 'dueño' : 'empleado';
  const jornadaHoras = parseJornadaHoras_(body.jornadaHoras);

  if (!nombre || nombre.length < 2) {
    return errorResponse_('Ingresá el nombre completo del empleado.', 'INVALID_NAME');
  }
  if (!usuario || !/^[a-z0-9._-]{3,30}$/.test(usuario)) {
    return errorResponse_('El usuario debe tener entre 3 y 30 caracteres (letras, números, punto o guión, sin espacios).', 'INVALID_USERNAME');
  }
  if (!/^\d{4,8}$/.test(pin)) {
    return errorResponse_('El PIN debe tener entre 4 y 8 dígitos numéricos.', 'INVALID_PIN');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const users = getAllUsers_();
    for (let i = 0; i < users.length; i++) {
      if (users[i].usuario === usuario) {
        return errorResponse_('Ya existe un empleado con el usuario "' + usuario + '". Elegí otro.', 'USER_EXISTS');
      }
    }

    const sh = getSheetUsuarios_();
    const pinHash = hashPin_(pin, usuario);
    const creado = formatNowDateTime_();
    const rowIdx = sh.getLastRow() + 1;

    const rng = sh.getRange(rowIdx, 1, 1, 7);
    rng.setNumberFormat('@');
    rng.setValues([[
      usuario,
      nombre,
      rol,
      pinHash,
      'SI',
      creado,
      String(jornadaHoras)
    ]]);
    invalidateUsersCache_();

    users.push({
      rowNumber: rowIdx,
      usuario: usuario,
      nombre: nombre,
      rol: rol,
      pinHash: pinHash,
      activo: true,
      creado: creado,
      jornadaHoras: jornadaHoras
    });

    const payload = buildEmpleadosPayload_(users, getAllShifts_());
    return jsonResponse_({
      ok: true,
      mensaje: 'Empleado "' + nombre + '" (' + usuario + ') creado con jornada diaria de ' + jornadaHoras + ' h.',
      ...payload
    });
  } finally {
    lock.releaseLock();
  }
}

function handleCambiarJornada_(body) {
  const targetUsername = normalizeUsername_(body.usuario);
  const jornadaHoras = parseJornadaHoras_(body.jornadaHoras);

  if (!targetUsername) {
    return errorResponse_('Falta indicar el usuario del empleado.', 'MISSING_USER');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const users = getAllUsers_();
    let target = null;
    for (let i = 0; i < users.length; i++) {
      if (users[i].usuario === targetUsername) {
        target = users[i];
        break;
      }
    }
    if (!target) {
      return errorResponse_('No se encontró el empleado "' + targetUsername + '".', 'USER_NOT_FOUND');
    }

    const sh = getSheetUsuarios_();
    if (sh.getRange(1, 7).getValue() === '') {
      sh.getRange(1, 7).setValue('Jornada hs').setFontWeight('bold').setBackground('#1a100a').setFontColor('#d9a05f');
    }
    sh.getRange(target.rowNumber, 7).setNumberFormat('@').setValue(String(jornadaHoras));
    target.jornadaHoras = jornadaHoras;
    invalidateUsersCache_();

    const payload = buildEmpleadosPayload_(users, getAllShifts_());
    return jsonResponse_({
      ok: true,
      jornadaHoras: jornadaHoras,
      mensaje: 'Jornada diaria de "' + target.nombre + '" actualizada a ' + jornadaHoras + ' h (' + minutesToHHMM_(Math.round(jornadaHoras * 60)) + ').',
      ...payload
    });
  } finally {
    lock.releaseLock();
  }
}

function handleToggleEmpleado_(currentUser, body) {
  const targetUsername = normalizeUsername_(body.usuario);
  if (!targetUsername) {
    return errorResponse_('Falta indicar el usuario del empleado.', 'MISSING_USER');
  }
  if (targetUsername === currentUser.usuario) {
    return errorResponse_('No podés deshabilitar tu propio usuario.', 'CANNOT_DISABLE_SELF');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const users = getAllUsers_();
    let target = null;
    for (let i = 0; i < users.length; i++) {
      if (users[i].usuario === targetUsername) {
        target = users[i];
        break;
      }
    }
    if (!target) {
      return errorResponse_('No se encontró el empleado "' + targetUsername + '".', 'USER_NOT_FOUND');
    }

    const nuevoEstado = !target.activo;
    const sh = getSheetUsuarios_();
    sh.getRange(target.rowNumber, 5).setNumberFormat('@').setValue(nuevoEstado ? 'SI' : 'NO');
    target.activo = nuevoEstado;
    invalidateUsersCache_();

    const payload = buildEmpleadosPayload_(users, getAllShifts_());
    return jsonResponse_({
      ok: true,
      activo: nuevoEstado,
      mensaje: 'El empleado "' + target.nombre + '" ahora está ' + (nuevoEstado ? 'HABILITADO' : 'DESHABILITADO') + '.',
      ...payload
    });
  } finally {
    lock.releaseLock();
  }
}

function handleEliminarEmpleado_(currentUser, body) {
  const targetUsername = normalizeUsername_(body.usuario);
  if (!targetUsername) {
    return errorResponse_('Falta indicar el usuario del empleado a eliminar.', 'MISSING_USER');
  }
  if (targetUsername === currentUser.usuario) {
    return errorResponse_('No podés eliminar tu propio usuario.', 'CANNOT_DELETE_SELF');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const users = getAllUsers_();
    let targetIdx = -1;
    for (let i = 0; i < users.length; i++) {
      if (users[i].usuario === targetUsername) {
        targetIdx = i;
        break;
      }
    }
    if (targetIdx === -1) {
      return errorResponse_('No se encontró el empleado "' + targetUsername + '".', 'USER_NOT_FOUND');
    }

    const target = users[targetIdx];
    const sh = getSheetUsuarios_();
    sh.deleteRow(target.rowNumber);
    users.splice(targetIdx, 1);
    invalidateUsersCache_();

    const payload = buildEmpleadosPayload_(users, getAllShifts_());
    return jsonResponse_({
      ok: true,
      mensaje: 'El usuario "' + target.nombre + '" (@' + target.usuario + ') fue eliminado.',
      ...payload
    });
  } finally {
    lock.releaseLock();
  }
}

function handleResetPin_(body) {
  const targetUsername = normalizeUsername_(body.usuario);
  const nuevoPin = String(body.pin || '').trim();

  if (!targetUsername) {
    return errorResponse_('Falta indicar el usuario del empleado.', 'MISSING_USER');
  }
  if (!/^\d{4,8}$/.test(nuevoPin)) {
    return errorResponse_('El nuevo PIN debe tener entre 4 y 8 dígitos numéricos.', 'INVALID_PIN');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const target = findUserByUsername_(targetUsername);
    if (!target) {
      return errorResponse_('No se encontró el empleado "' + targetUsername + '".', 'USER_NOT_FOUND');
    }

    const nuevoHash = hashPin_(nuevoPin, target.usuario);
    const sh = getSheetUsuarios_();
    sh.getRange(target.rowNumber, 4).setNumberFormat('@').setValue(nuevoHash);
    invalidateUsersCache_();

    return jsonResponse_({
      ok: true,
      mensaje: 'Se actualizó el PIN de "' + target.nombre + '" (' + target.usuario + ').'
    });
  } finally {
    lock.releaseLock();
  }
}

function handleEgresoForzado_(body) {
  const targetUsername = normalizeUsername_(body.usuario);
  const horaOpcional = String(body.hora || '').trim();

  if (!targetUsername) {
    return errorResponse_('Falta indicar el usuario del empleado.', 'MISSING_USER');
  }
  if (horaOpcional && !/^([01]\d|2[0-3]):([0-5]\d)$/.test(horaOpcional)) {
    return errorResponse_('El formato de hora debe ser HH:MM (por ejemplo 18:30).', 'INVALID_TIME_FORMAT');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheetFichajes_();
    const shifts = getAllShifts_();
    let openShift = null;
    for (let i = shifts.length - 1; i >= 0; i--) {
      if (shifts[i].usuario === targetUsername && shifts[i].estado === 'ABIERTO') {
        openShift = shifts[i];
        break;
      }
    }

    if (!openShift) {
      return errorResponse_('El empleado "' + targetUsername + '" no tiene ningún turno abierto en este momento.', 'NO_OPEN_SHIFT');
    }

    let egresoMs;
    let horaEgreso;

    if (horaOpcional) {
      horaEgreso = horaOpcional;
      const candidateMs = parseCordobaDateTime_(openShift.fecha, horaOpcional);
      if (candidateMs < openShift.ingresoTS) {
        egresoMs = candidateMs + 24 * 60 * 60 * 1000;
      } else {
        egresoMs = candidateMs;
      }
    } else {
      const now = new Date();
      egresoMs = Math.max(now.getTime(), openShift.ingresoTS + 1000);
      horaEgreso = formatTimeHM_(new Date(egresoMs));
    }

    const calc = calculateDurationExact_(openShift.ingreso, horaEgreso, openShift.ingresoTS, egresoMs);

    const rng = sh.getRange(openShift.rowNumber, 6, 1, 6);
    rng.setNumberFormat('@');
    rng.setValues([[
      horaEgreso,
      String(calc.horasDecimal),
      calc.duracionHHMM,
      'CERRADO',
      String(openShift.ingresoTS),
      String(egresoMs)
    ]]);

    openShift.egreso = horaEgreso;
    openShift.horasDecimal = calc.horasDecimal;
    openShift.duracion = calc.duracionHHMM;
    openShift.estado = 'CERRADO';
    openShift.egresoTS = egresoMs;

    const payload = buildEmpleadosPayload_(getAllUsers_(), shifts);
    return jsonResponse_({
      ok: true,
      mensaje: 'Egreso forzado aplicado a "' + openShift.nombre + '" a las ' + horaEgreso + ' hs (' + calc.duracionHHMM + ' hs).',
      ...payload
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * REPORTE MENSUAL Y POR PERÍODO (Solo dueño):
 * Calcula por empleado y por día:
 * - Total trabajado (HH:MM)
 * - Horas normales (hasta su jornada diaria configurada, 8 h por defecto)
 * - Horas extra (lo que supere la jornada diaria ese día, sin compensar con días más cortos)
 * - Turnos abiertos quedan como PENDIENTE
 */
function handleReporte_(body) {
  const defaults = getCurrentMonthRange_();
  const desde = normalizeDateFilter_(body.desde, defaults.desde);
  const hasta = normalizeDateFilter_(body.hasta, defaults.hasta);
  const empleadoFiltro = normalizeUsername_(body.usuario || '');

  const users = getAllUsers_();
  const allShifts = getAllShifts_();

  const jornadaMap = {};
  users.forEach(function(u) {
    jornadaMap[u.usuario] = u.jornadaHoras || DEFAULT_JORNADA_HORAS;
  });

  const filtered = allShifts.filter(function(s) {
    const matchDate = s.fecha >= desde && s.fecha <= hasta;
    const matchUser = !empleadoFiltro || empleadoFiltro === 'todos' || s.usuario === empleadoFiltro;
    return matchDate && matchUser;
  });

  const targetUsers = (!empleadoFiltro || empleadoFiltro === 'todos')
    ? users
    : users.filter(function(u) { return u.usuario === empleadoFiltro; });

  const processed = computeDailyAndMonthlyHours_(filtered, jornadaMap, targetUsers);

  return jsonResponse_({
    ok: true,
    desde: desde,
    hasta: hasta,
    usuarioFiltro: empleadoFiltro || 'todos',
    empleadosLista: users.map(function(u) {
      return {
        usuario: u.usuario,
        nombre: u.nombre,
        activo: u.activo,
        jornadaHoras: u.jornadaHoras || DEFAULT_JORNADA_HORAS
      };
    }),
    resumenPorEmpleado: processed.resumenPorEmpleado,
    resumenDiario: processed.resumenDiario,
    fichajes: processed.fichajesEnriquecidos,
    totalGeneralHHMM: processed.totalesGenerales.totalTrabajadoHHMM,
    totalGeneralNormalesHHMM: processed.totalesGenerales.horasNormalesHHMM,
    totalGeneralExtrasHHMM: processed.totalesGenerales.horasExtrasHHMM,
    totalGeneralDecimal: processed.totalesGenerales.horasDecimal,
    cantidadTurnos: filtered.length
  });
}

/* ==========================================================================
 * MOTOR DE CÁLCULO DIARIO Y MENSUAL (NORMALES Y EXTRAS EN MINUTOS EXACTOS)
 * ========================================================================== */

/**
 * Calcula las horas trabajadas, normales y extras agrupando por empleado y por día.
 * - Suma los turnos cerrados del día y registra como extra lo que supere la jornada diaria (8 h por defecto).
 * - Suma las extras de cada día sin compensarlas con jornadas más cortas.
 * - Deja los turnos abiertos como PENDIENTE.
 */
function computeDailyAndMonthlyHours_(shifts, jornadaMap, usersList) {
  // 1. Clonar y ordenar cronológicamente (ascendente) para acumular correctamente dentro de cada día
  const sortedAsc = shifts.map(function(s) {
    return Object.assign({}, s);
  }).sort(function(a, b) {
    if (a.fecha !== b.fecha) return a.fecha < b.fecha ? -1 : 1;
    return (a.ingresoTS || 0) - (b.ingresoTS || 0);
  });

  const dayAccumByUserDate = {};
  const dayGroupMap = {};

  sortedAsc.forEach(function(s) {
    const u = s.usuario;
    const f = s.fecha;
    const key = u + '|' + f;
    const jornadaHoras = (jornadaMap && jornadaMap[u]) ? Number(jornadaMap[u]) : DEFAULT_JORNADA_HORAS;
    const jornadaMin = Math.round(jornadaHoras * 60);

    if (!dayAccumByUserDate[key]) {
      dayAccumByUserDate[key] = 0;
    }
    if (!dayGroupMap[key]) {
      dayGroupMap[key] = {
        usuario: u,
        nombre: s.nombre,
        fecha: f,
        jornadaHoras: jornadaHoras,
        jornadaHHMM: minutesToHHMM_(jornadaMin),
        ingresos: [],
        egresos: [],
        turnosCerrados: 0,
        turnosPendientes: 0,
        minutosTrabajados: 0,
        minutosNormales: 0,
        minutosExtras: 0
      };
    }

    const grp = dayGroupMap[key];
    if (s.ingreso) grp.ingresos.push(s.ingreso);

    if (s.estado !== 'CERRADO') {
      // Turno abierto -> queda como PENDIENTE
      s.esPendiente = true;
      s.estadoDisplay = 'PENDIENTE';
      s.minutosTrabajados = 0;
      s.minutosNormales = 0;
      s.minutosExtras = 0;
      s.totalHHMM = 'Pendiente';
      s.normalesHHMM = 'Pendiente';
      s.extrasHHMM = 'Pendiente';
      grp.turnosPendientes += 1;
      grp.egresos.push('Pendiente');
    } else {
      const durMin = getExactShiftMinutes_(s);
      const prevAcum = dayAccumByUserDate[key];
      const nuevoAcum = prevAcum + durMin;
      dayAccumByUserDate[key] = nuevoAcum;

      const normMin = Math.max(0, Math.min(nuevoAcum, jornadaMin) - Math.min(prevAcum, jornadaMin));
      const extMin = Math.max(0, durMin - normMin);

      s.esPendiente = false;
      s.estadoDisplay = 'CERRADO';
      s.minutosTrabajados = durMin;
      s.minutosNormales = normMin;
      s.minutosExtras = extMin;
      s.duracion = minutesToHHMM_(durMin);
      s.totalHHMM = minutesToHHMM_(durMin);
      s.normalesHHMM = minutesToHHMM_(normMin);
      s.extrasHHMM = minutesToHHMM_(extMin);
      s.horasDecimal = Math.round((durMin / 60) * 100) / 100;

      grp.turnosCerrados += 1;
      if (s.egreso) grp.egresos.push(s.egreso);
      grp.minutosTrabajados = nuevoAcum;
      grp.minutosNormales = Math.min(nuevoAcum, jornadaMin);
      grp.minutosExtras = Math.max(0, nuevoAcum - jornadaMin);
    }
  });

  // 2. Construir lista de resumen diario ordenada de más reciente a más antigua
  const resumenDiario = Object.keys(dayGroupMap).map(function(k) {
    const d = dayGroupMap[k];
    return {
      usuario: d.usuario,
      nombre: d.nombre,
      fecha: d.fecha,
      jornadaHoras: d.jornadaHoras,
      jornadaHHMM: d.jornadaHHMM,
      ingreso: d.ingresos.length > 0 ? d.ingresos.join(' / ') : '—',
      egreso: d.egresos.length > 0 ? d.egresos.join(' / ') : '—',
      turnosCerrados: d.turnosCerrados,
      turnosPendientes: d.turnosPendientes,
      minutosTrabajados: d.minutosTrabajados,
      minutosNormales: d.minutosNormales,
      minutosExtras: d.minutosExtras,
      totalHHMM: minutesToHHMM_(d.minutosTrabajados),
      normalesHHMM: minutesToHHMM_(d.minutosNormales),
      extrasHHMM: minutesToHHMM_(d.minutosExtras),
      estado: d.turnosPendientes > 0 ? 'PENDIENTE' : 'CERRADO'
    };
  }).sort(function(a, b) {
    if (a.fecha !== b.fecha) return a.fecha > b.fecha ? -1 : 1;
    return a.nombre.localeCompare(b.nombre);
  });

  // 3. Resumen mensual por empleado (sumando cada día sin compensar extras con días más cortos)
  const empMap = {};
  if (Array.isArray(usersList)) {
    usersList.forEach(function(u) {
      const jH = (jornadaMap && jornadaMap[u.usuario]) ? Number(jornadaMap[u.usuario]) : (u.jornadaHoras || DEFAULT_JORNADA_HORAS);
      empMap[u.usuario] = {
        usuario: u.usuario,
        nombre: u.nombre,
        jornadaHoras: jH,
        jornadaHHMM: minutesToHHMM_(Math.round(jH * 60)),
        diasTrabajados: 0,
        turnosCerrados: 0,
        turnosAbiertos: 0,
        minutosTrabajados: 0,
        minutosNormales: 0,
        minutosExtras: 0
      };
    });
  }

  Object.keys(dayGroupMap).forEach(function(k) {
    const d = dayGroupMap[k];
    if (!empMap[d.usuario]) {
      empMap[d.usuario] = {
        usuario: d.usuario,
        nombre: d.nombre,
        jornadaHoras: d.jornadaHoras,
        jornadaHHMM: d.jornadaHHMM,
        diasTrabajados: 0,
        turnosCerrados: 0,
        turnosAbiertos: 0,
        minutosTrabajados: 0,
        minutosNormales: 0,
        minutosExtras: 0
      };
    }
    const emp = empMap[d.usuario];
    if (d.turnosCerrados > 0) emp.diasTrabajados += 1;
    emp.turnosCerrados += d.turnosCerrados;
    emp.turnosAbiertos += d.turnosPendientes;
    emp.minutosTrabajados += d.minutosTrabajados;
    emp.minutosNormales += d.minutosNormales;
    emp.minutosExtras += d.minutosExtras;
  });

  let globalTrabajadosMin = 0;
  let globalNormalesMin = 0;
  let globalExtrasMin = 0;
  let globalCerrados = 0;
  let globalPendientes = 0;

  const resumenPorEmpleado = Object.keys(empMap).map(function(uKey) {
    const e = empMap[uKey];
    globalTrabajadosMin += e.minutosTrabajados;
    globalNormalesMin += e.minutosNormales;
    globalExtrasMin += e.minutosExtras;
    globalCerrados += e.turnosCerrados;
    globalPendientes += e.turnosAbiertos;

    const dec = Math.round((e.minutosTrabajados / 60) * 100) / 100;
    return {
      usuario: e.usuario,
      nombre: e.nombre,
      jornadaHoras: e.jornadaHoras,
      jornadaHHMM: e.jornadaHHMM,
      diasTrabajados: e.diasTrabajados,
      turnosCerrados: e.turnosCerrados,
      turnosAbiertos: e.turnosAbiertos,
      totalTurnos: e.turnosCerrados + e.turnosAbiertos,
      minutosTrabajados: e.minutosTrabajados,
      minutosNormales: e.minutosNormales,
      minutosExtras: e.minutosExtras,
      duracionHHMM: minutesToHHMM_(e.minutosTrabajados),
      totalTrabajadoHHMM: minutesToHHMM_(e.minutosTrabajados),
      horasNormalesHHMM: minutesToHHMM_(e.minutosNormales),
      horasExtrasHHMM: minutesToHHMM_(e.minutosExtras),
      horasDecimal: dec
    };
  }).sort(function(a, b) {
    return b.minutosTrabajados - a.minutosTrabajados;
  });

  // Ordenar fichajes enriquecidos del más reciente al más antiguo para mostrar en tabla
  const fichajesDesc = sortedAsc.slice().sort(function(a, b) {
    if (a.fecha !== b.fecha) return a.fecha > b.fecha ? -1 : 1;
    return (b.ingresoTS || 0) - (a.ingresoTS || 0);
  });

  return {
    fichajesEnriquecidos: fichajesDesc,
    resumenDiario: resumenDiario,
    resumenPorEmpleado: resumenPorEmpleado,
    totalesGenerales: {
      minutosTrabajados: globalTrabajadosMin,
      minutosNormales: globalNormalesMin,
      minutosExtras: globalExtrasMin,
      totalTrabajadoHHMM: minutesToHHMM_(globalTrabajadosMin),
      horasNormalesHHMM: minutesToHHMM_(globalNormalesMin),
      horasExtrasHHMM: minutesToHHMM_(globalExtrasMin),
      horasDecimal: Math.round((globalTrabajadosMin / 60) * 100) / 100,
      turnosCerrados: globalCerrados,
      turnosPendientes: globalPendientes
    }
  };
}

/* ==========================================================================
 * FUNCIONES AUXILIARES DE DATOS Y HOJAS
 * ========================================================================== */

function getSheetUsuarios_() {
  if (_shUsuariosCache) return _shUsuariosCache;
  const ss = getSS_();
  let sh = ss.getSheetByName(SHEET_USUARIOS);
  if (!sh) {
    setup();
    sh = ss.getSheetByName(SHEET_USUARIOS);
  }
  _shUsuariosCache = sh;
  return sh;
}

function getSheetFichajes_() {
  if (_shFichajesCache) return _shFichajesCache;
  const ss = getSS_();
  let sh = ss.getSheetByName(SHEET_FICHAJES);
  if (!sh) {
    setup();
    sh = ss.getSheetByName(SHEET_FICHAJES);
  }
  _shFichajesCache = sh;
  return sh;
}

function invalidateUsersCache_() {
  try {
    CacheService.getScriptCache().remove('USERS_LIST_CACHE_V2');
  } catch (e) {}
}

function getAllUsersFast_() {
  try {
    const cached = CacheService.getScriptCache().get('USERS_LIST_CACHE_V2');
    if (cached) return JSON.parse(cached);
  } catch (e) {}
  const list = getAllUsers_();
  try {
    CacheService.getScriptCache().put('USERS_LIST_CACHE_V2', JSON.stringify(list), 120);
  } catch (e) {}
  return list;
}

function findUserByUsernameFast_(username) {
  const target = normalizeUsername_(username);
  if (!target) return null;
  const users = getAllUsersFast_();
  for (let i = 0; i < users.length; i++) {
    if (users[i].usuario === target) return users[i];
  }
  return findUserByUsername_(target);
}

function parseJornadaHoras_(val) {
  if (val === undefined || val === null || String(val).trim() === '') {
    return DEFAULT_JORNADA_HORAS;
  }
  const s = String(val).trim().replace(',', '.');
  if (/^\d{1,2}:\d{2}$/.test(s)) {
    const parts = s.split(':');
    const h = parseInt(parts[0], 10) + parseInt(parts[1], 10) / 60;
    return h > 0 && h <= 24 ? Math.round(h * 100) / 100 : DEFAULT_JORNADA_HORAS;
  }
  const num = parseFloat(s);
  if (isNaN(num) || num <= 0 || num > 24) return DEFAULT_JORNADA_HORAS;
  return Math.round(num * 100) / 100;
}

function getAllUsers_() {
  const sh = getSheetUsuarios_();
  const values = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const usuario = normalizeUsername_(row[0]);
    if (!usuario) continue;
    const activoRaw = String(row[4] || '').trim().toUpperCase();
    const activo = activoRaw === 'SI' || activoRaw === 'TRUE' || activoRaw === '1' || activoRaw === 'ACTIVO';
    const jornadaHoras = parseJornadaHoras_(row[6]);
    list.push({
      rowNumber: i + 1,
      usuario: usuario,
      nombre: String(row[1] || '').trim(),
      rol: String(row[2] || 'empleado').trim().toLowerCase(),
      pinHash: String(row[3] || '').trim(),
      activo: activo,
      creado: formatCellString_(row[5]),
      jornadaHoras: jornadaHoras
    });
  }
  return list;
}

function findUserByUsername_(username) {
  const target = normalizeUsername_(username);
  if (!target) return null;
  const users = getAllUsers_();
  for (let i = 0; i < users.length; i++) {
    if (users[i].usuario === target) {
      return users[i];
    }
  }
  return null;
}

function getAllShifts_() {
  const sh = getSheetFichajes_();
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];
  const values = sh.getRange(2, 1, lastRow - 1, 11).getValues();
  const list = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const id = String(row[0] || '').trim();
    const usuario = normalizeUsername_(row[1]);
    if (!id && !usuario) continue;

    const fecha = normalizeCellDate_(row[3]);
    const ingreso = normalizeCellTime_(row[4]);
    const egreso = normalizeCellTime_(row[5]);
    const horasDecRaw = String(row[6] !== undefined && row[6] !== null ? row[6] : '0').replace(',', '.');
    const horasDecimal = Math.round((parseFloat(horasDecRaw) || 0) * 100) / 100;
    const duracion = normalizeCellDuration_(row[7]);
    const estado = String(row[8] || 'CERRADO').trim().toUpperCase();
    const ingresoTS = Number(row[9]) || parseCordobaDateTime_(fecha, ingreso || '00:00');
    const egresoTS = row[10] ? Number(row[10]) : null;

    list.push({
      rowNumber: i + 2,
      id: id,
      usuario: usuario,
      nombre: String(row[2] || '').trim(),
      fecha: fecha,
      ingreso: ingreso,
      egreso: egreso,
      horasDecimal: horasDecimal,
      duracion: duracion,
      estado: estado,
      ingresoTS: ingresoTS,
      egresoTS: egresoTS
    });
  }
  return list;
}

function buildEmpleadosPayload_(users, shifts) {
  const openByUser = {};
  for (let i = 0; i < shifts.length; i++) {
    const s = shifts[i];
    if (s.estado === 'ABIERTO') {
      openByUser[s.usuario] = s;
    }
  }

  const empleados = users.map(function(u) {
    const turnoAbierto = openByUser[u.usuario] || null;
    const jH = u.jornadaHoras || DEFAULT_JORNADA_HORAS;
    return {
      usuario: u.usuario,
      nombre: u.nombre,
      rol: u.rol,
      activo: u.activo,
      creado: u.creado,
      jornadaHoras: jH,
      jornadaHHMM: minutesToHHMM_(Math.round(jH * 60)),
      enTurno: Boolean(turnoAbierto),
      turnoAbierto: turnoAbierto
    };
  });

  const adentroAhora = empleados
    .filter(function(e) { return e.enTurno; })
    .map(function(e) {
      return {
        usuario: e.usuario,
        nombre: e.nombre,
        fecha: e.turnoAbierto.fecha,
        ingreso: e.turnoAbierto.ingreso,
        ingresoTS: e.turnoAbierto.ingresoTS,
        fichajeId: e.turnoAbierto.id
      };
    });

  return {
    empleados: empleados,
    adentroAhora: adentroAhora
  };
}

function buildEstadoUsuarioFromShifts_(user, shifts) {
  const hoy = formatDateYMD_(new Date());
  let turnoAbierto = null;
  const fichajesHoyRaw = [];

  for (let i = 0; i < shifts.length; i++) {
    const s = shifts[i];
    if (s.usuario === user.usuario) {
      if (s.estado === 'ABIERTO') {
        turnoAbierto = s;
      }
      if (s.fecha === hoy || s.estado === 'ABIERTO') {
        fichajesHoyRaw.push(s);
      }
    }
  }

  const jornadaMap = {};
  jornadaMap[user.usuario] = user.jornadaHoras || DEFAULT_JORNADA_HORAS;
  const processedHoy = computeDailyAndMonthlyHours_(fichajesHoyRaw, jornadaMap, [user]);

  return {
    enTurno: Boolean(turnoAbierto),
    turnoAbierto: turnoAbierto,
    fechaHoy: hoy,
    jornadaHoras: user.jornadaHoras || DEFAULT_JORNADA_HORAS,
    serverTimeMs: Date.now(),
    fichajesHoy: processedHoy.fichajesEnriquecidos,
    totalHoyHHMM: processedHoy.totalesGenerales.totalTrabajadoHHMM,
    totalHoyNormalesHHMM: processedHoy.totalesGenerales.horasNormalesHHMM,
    totalHoyExtrasHHMM: processedHoy.totalesGenerales.horasExtrasHHMM,
    totalHoyDecimal: processedHoy.totalesGenerales.horasDecimal
  };
}

/* ==========================================================================
 * CÁLCULOS DE TIEMPO EXACTOS, HASH Y UTILIDADES
 * ========================================================================== */

function hashPin_(pin, username) {
  const raw = SALT + ':' + normalizeUsername_(username) + ':' + String(pin).trim();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytesToHex_(digest);
}

function hashPinLegacy_(pin) {
  const raw = SALT + ':' + String(pin).trim();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return bytesToHex_(digest);
}

function bytesToHex_(bytes) {
  return bytes.map(function(b) {
    const v = (b < 0) ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function getSession_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('SESSION_' + token);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function isDueno_(user) {
  if (!user || !user.rol) return false;
  const r = String(user.rol).trim().toLowerCase();
  return r === 'dueño' || r === 'dueno';
}

function normalizeUsername_(u) {
  return String(u || '').trim().toLowerCase();
}

/**
 * Calcula la duración exacta en minutos entre la hora de ingreso (HH:MM) y egreso (HH:MM).
 */
function calculateDurationExact_(ingresoHM, egresoHM, ingresoTS, egresoTS) {
  const inMin = parseHMToMinutes_(ingresoHM);
  const outMin = parseHMToMinutes_(egresoHM);
  let totalMinutes = 0;

  if (inMin !== null && outMin !== null) {
    if (outMin >= inMin) {
      totalMinutes = outMin - inMin;
    } else {
      totalMinutes = (outMin + 1440) - inMin;
    }
  } else {
    const diffMs = Math.max(0, (egresoTS || 0) - (ingresoTS || 0));
    totalMinutes = Math.round(diffMs / 60000);
  }

  const duracionHHMM = minutesToHHMM_(totalMinutes);
  const horasDecimal = Math.round((totalMinutes / 60) * 100) / 100;
  return {
    totalMinutes: totalMinutes,
    duracionHHMM: duracionHHMM,
    horasDecimal: horasDecimal
  };
}

function getExactShiftMinutes_(s) {
  if (!s || s.estado !== 'CERRADO') return 0;
  // 1. Si tiene Duración hh:mm válida mayor a 0, usar esos minutos exactos
  const durStr = String(s.duracion || '').trim();
  if (/^\d{1,4}:\d{2}$/.test(durStr)) {
    const parts = durStr.split(':');
    const mins = parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
    if (mins > 0) return mins;
  }
  // 2. Si tiene Ingreso y Egreso en formato HH:MM, calcular diferencia exacta
  const inMin = parseHMToMinutes_(s.ingreso);
  const outMin = parseHMToMinutes_(s.egreso);
  if (inMin !== null && outMin !== null) {
    return outMin >= inMin ? (outMin - inMin) : ((outMin + 1440) - inMin);
  }
  // 3. Fallback a horasDecimal
  return Math.round((Number(s.horasDecimal) || 0) * 60);
}

function parseHMToMinutes_(hmStr) {
  const s = String(hmStr || '').trim();
  const m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
}

function minutesToHHMM_(totalMinutes) {
  const m = Math.max(0, Math.round(Number(totalMinutes) || 0));
  const h = Math.floor(m / 60);
  const mins = m % 60;
  return pad2_(h) + ':' + pad2_(mins);
}

function pad2_(n) {
  return n < 10 ? '0' + n : String(n);
}

function formatDateYMD_(dateObj) {
  return Utilities.formatDate(dateObj, TZ, 'yyyy-MM-dd');
}

function formatTimeHM_(dateObj) {
  return Utilities.formatDate(dateObj, TZ, 'HH:mm');
}

function formatNowDateTime_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function getCurrentMonthRange_() {
  const now = new Date();
  const ym = Utilities.formatDate(now, TZ, 'yyyy-MM');
  const hoy = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  return {
    desde: ym + '-01',
    hasta: hoy
  };
}

function normalizeDateFilter_(val, fallback) {
  const s = String(val || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return fallback;
}

function parseCordobaDateTime_(ymd, hm) {
  const safeYMD = /^\d{4}-\d{2}-\d{2}$/.test(ymd) ? ymd : formatDateYMD_(new Date());
  const safeHM = /^([01]\d|2[0-3]):([0-5]\d)$/.test(hm) ? hm : '00:00';
  const iso = safeYMD + 'T' + safeHM + ':00-03:00';
  const parsed = new Date(iso).getTime();
  return isNaN(parsed) ? Date.now() : parsed;
}

function normalizeCellDate_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, TZ, 'yyyy-MM-dd');
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

function normalizeCellTime_(val) {
  if (!val && val !== 0) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, TZ, 'HH:mm');
  }
  const s = String(val).trim();
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(s)) {
    const parts = s.split(':');
    return pad2_(parseInt(parts[0], 10)) + ':' + parts[1];
  }
  return s;
}

function normalizeCellDuration_(val) {
  if (!val) return '00:00';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, TZ, 'HH:mm');
  }
  return String(val).trim() || '00:00';
}

function formatCellString_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, TZ, 'yyyy-MM-dd HH:mm:ss');
  }
  return String(val).trim();
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse_(msg, code) {
  return jsonResponse_({
    ok: false,
    error: msg,
    code: code || 'ERROR'
  });
}

function forbiddenResponse_() {
  return errorResponse_('No tenés permisos de dueño para realizar esta acción.', 'FORBIDDEN');
}
