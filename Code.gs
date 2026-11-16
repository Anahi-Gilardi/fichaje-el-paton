/**
 * EL PATRÓN — MÓDULO DE FICHAJE (Control de Horarios)
 * Backend para Google Apps Script vinculado a Google Sheets.
 * Zona horaria fija: America/Argentina/Cordoba
 *
 * INSTRUCCIONES RÁPIDAS:
 * 1. En el editor de Apps Script, seleccioná la función "setup" arriba y hacé clic en "Ejecutar".
 * 2. Autorizá los permisos cuando Google te lo pida.
 * 3. Hacé clic en "Implementar" > "Nueva implementación" > Tipo: "Aplicación web".
 * 4. Ejecutar como: "Yo" | Quién tiene acceso: "Cualquier persona".
 * 5. Copiá la URL que termina en /exec y pegala en el frontend (index.html).
 */

const TZ = 'America/Argentina/Cordoba';
const SALT = 'EL_PATRON_FICHAJE_SALT_2026_CBA_AR_#94653f'; // Podés cambiar este SALT antes de correr setup()
const SESSION_TTL_SECONDS = 21600; // 6 horas (6 * 60 * 60)
const MAX_LOGIN_ATTEMPTS = 5;
const LOCKOUT_SECONDS = 300; // 5 minutos (5 * 60)

const SHEET_USUARIOS = 'Usuarios';
const SHEET_FICHAJES = 'Fichajes';

/**
 * Inicializa la planilla creando las hojas "Usuarios" y "Fichajes"
 * y el usuario inicial: dueno / PIN 1234 (rol: dueño).
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Este script debe estar vinculado a una Google Sheet (Extensiones > Apps Script).');
  }

  try {
    ss.setSpreadsheetTimeZone(TZ);
  } catch (err) {
    Logger.log('Aviso al configurar zona horaria: ' + err);
  }

  // 1. Hoja Usuarios: Usuario, Nombre, Rol, PIN hash, Activo, Creado
  let shUsuarios = ss.getSheetByName(SHEET_USUARIOS);
  if (!shUsuarios) {
    shUsuarios = ss.insertSheet(SHEET_USUARIOS);
  }
  if (shUsuarios.getLastRow() === 0) {
    const headersU = ['Usuario', 'Nombre', 'Rol', 'PIN hash', 'Activo', 'Creado'];
    shUsuarios.getRange(1, 1, 1, headersU.length).setValues([headersU]);
    shUsuarios.getRange(1, 1, 1, headersU.length)
      .setFontWeight('bold')
      .setBackground('#1a100a')
      .setFontColor('#d9a05f');
    shUsuarios.setFrozenRows(1);
  }

  // Verificar si existen los usuarios iniciales "dueno" y "mozo"
  const usuariosData = shUsuarios.getDataRange().getValues();
  let existeDueno = false;
  let existeMozo = false;
  for (let i = 1; i < usuariosData.length; i++) {
    const uRow = String(usuariosData[i][0]).trim().toLowerCase();
    if (uRow === 'dueno') existeDueno = true;
    if (uRow === 'mozo') existeMozo = true;
  }

  const ahoraStr = formatNowDateTime_();

  if (!existeDueno) {
    const pinHash = hashPin_('1234', 'dueno');
    const nextRow = shUsuarios.getLastRow() + 1;
    shUsuarios.getRange(nextRow, 1, 1, 6).setNumberFormat('@');
    shUsuarios.getRange(nextRow, 1, 1, 6).setValues([[
      'dueno',
      'Dueño El Patrón',
      'dueño',
      pinHash,
      'SI',
      ahoraStr
    ]]);
  }

  if (!existeMozo) {
    const pinHashMozo = hashPin_('1234', 'mozo');
    const nextRow = shUsuarios.getLastRow() + 1;
    shUsuarios.getRange(nextRow, 1, 1, 6).setNumberFormat('@');
    shUsuarios.getRange(nextRow, 1, 1, 6).setValues([[
      'mozo',
      'Martín Gómez (Mozo)',
      'empleado',
      pinHashMozo,
      'SI',
      ahoraStr
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

  Logger.log('Setup completado con éxito. Usuario inicial: dueno / PIN: 1234');
  return 'Setup completado con éxito. Usuario inicial: dueno / PIN: 1234';
}

/**
 * Endpoint informativo para verificar en el navegador que el Web App está activo.
 */
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
    mensaje: 'El servidor de fichaje está activo y las hojas Usuarios y Fichajes fueron inicializadas correctamente.'
  });
}

/**
 * Endpoint principal. Recibe JSON enviado como text/plain para evitar el preflight CORS.
 * Cuerpo esperado: { action, token, ...datos }
 */
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

    // Acción pública: login
    if (action === 'login') {
      return handleLogin_(body);
    }

    // Todas las demás acciones requieren token de sesión válido
    const token = String(body.token || '').trim();
    const session = getSession_(token);
    if (!session) {
      return errorResponse_('Tu sesión expiró o no es válida. Volvé a iniciar sesión.', 'UNAUTHORIZED');
    }

    // Revalidar que el usuario siga activo en la hoja
    const currentUser = findUserByUsername_(session.usuario);
    if (!currentUser || !currentUser.activo) {
      if (token) CacheService.getScriptCache().remove('SESSION_' + token);
      return errorResponse_('Tu usuario fue deshabilitado. Consultá con el dueño.', 'USER_DISABLED');
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

      case 'toggleEmpleado':
        if (!isDueno_(currentUser)) return forbiddenResponse_();
        return handleToggleEmpleado_(currentUser, body);

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

/**
 * 1. LOGIN: Valida usuario y PIN y genera token en CacheService por 6 horas.
 */
function handleLogin_(body) {
  const usuarioInput = normalizeUsername_(body.usuario);
  const pinInput = String(body.pin || '').trim();

  if (!usuarioInput) {
    return errorResponse_('Ingresá tu nombre de usuario.', 'MISSING_USER');
  }
  if (!/^\d{4,8}$/.test(pinInput)) {
    return errorResponse_('El PIN debe tener entre 4 y 8 dígitos numéricos.', 'INVALID_PIN_FORMAT');
  }

  ensureSheetsExist_();
  const user = findUserByUsername_(usuarioInput);

  // Verificar existencia y PIN (soportando hash con usuario y hash simple con SALT)
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

  const estadoData = buildEstadoUsuario_(user);

  return jsonResponse_({
    ok: true,
    token: token,
    usuario: {
      usuario: user.usuario,
      nombre: user.nombre,
      rol: user.rol
    },
    estado: estadoData
  });
}

/**
 * 2. ESTADO: Devuelve si el empleado está en turno, el turno abierto (si tiene)
 * y la lista de fichajes de hoy con su total.
 */
function handleEstado_(user) {
  const estado = buildEstadoUsuario_(user);
  return jsonResponse_({
    ok: true,
    usuario: {
      usuario: user.usuario,
      nombre: user.nombre,
      rol: user.rol
    },
    ...estado
  });
}

/**
 * 3. INGRESO: Registra entrada del empleado.
 * Usa LockService para evitar doble clic o doble pestaña.
 */
function handleIngreso_(user) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return errorResponse_('El servidor está procesando otro fichaje. Intentá de nuevo en unos segundos.', 'LOCK_TIMEOUT');
  }

  try {
    const sh = getSheetFichajes_();
    const openShift = findOpenShiftForUser_(sh, user.usuario);
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
      0,
      '00:00',
      'ABIERTO',
      String(nowMs),
      ''
    ];

    const rowIdx = sh.getLastRow() + 1;
    sh.getRange(rowIdx, 1, 1, newRow.length).setNumberFormat('@');
    sh.getRange(rowIdx, 1, 1, newRow.length).setValues([newRow]);
    SpreadsheetApp.flush();

    const estado = buildEstadoUsuario_(user);
    return jsonResponse_({
      ok: true,
      mensaje: 'Ingreso registrado a las ' + horaIngreso + ' hs.',
      ...estado
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 4. EGRESO: Cierra el turno abierto del empleado.
 * Usa LockService para evitar doble clic o doble pestaña.
 */
function handleEgreso_(user) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return errorResponse_('El servidor está procesando otro fichaje. Intentá de nuevo en unos segundos.', 'LOCK_TIMEOUT');
  }

  try {
    const sh = getSheetFichajes_();
    const openShift = findOpenShiftForUser_(sh, user.usuario);
    if (!openShift) {
      return errorResponse_(
        'No tenés ningún turno abierto para cerrar. Primero tenés que marcar INGRESO.',
        'NO_OPEN_SHIFT'
      );
    }

    const now = new Date();
    const egresoMs = Math.max(now.getTime(), openShift.ingresoTS + 1000);
    const horaEgreso = formatTimeHM_(new Date(egresoMs));
    const calc = calculateDuration_(openShift.ingresoTS, egresoMs);

    // Columnas: 6=Egreso, 7=Horas decimal, 8=Duración hh:mm, 9=Estado, 11=EgresoTS
    sh.getRange(openShift.rowNumber, 6, 1, 6).setNumberFormat('@');
    sh.getRange(openShift.rowNumber, 6).setValue(horaEgreso);
    sh.getRange(openShift.rowNumber, 7).setValue(String(calc.horasDecimal));
    sh.getRange(openShift.rowNumber, 8).setValue(calc.duracionHHMM);
    sh.getRange(openShift.rowNumber, 9).setValue('CERRADO');
    sh.getRange(openShift.rowNumber, 11).setValue(String(egresoMs));
    SpreadsheetApp.flush();

    const estado = buildEstadoUsuario_(user);
    return jsonResponse_({
      ok: true,
      mensaje: 'Egreso registrado a las ' + horaEgreso + ' hs. Duración del turno: ' + calc.duracionHHMM + ' hs.',
      ...estado
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 5. MIS HORAS: Filtra los fichajes del usuario logueado por rango de fechas (desde, hasta).
 */
function handleMisHoras_(user, body) {
  const defaults = getCurrentMonthRange_();
  const desde = normalizeDateFilter_(body.desde, defaults.desde);
  const hasta = normalizeDateFilter_(body.hasta, defaults.hasta);

  const allShifts = getAllShifts_();
  const filtered = allShifts.filter(function(item) {
    return item.usuario === user.usuario && item.fecha >= desde && item.fecha <= hasta;
  });

  // Ordenar del más reciente al más antiguo
  filtered.sort(function(a, b) {
    return (b.ingresoTS || 0) - (a.ingresoTS || 0);
  });

  const totals = summarizeShifts_(filtered);

  return jsonResponse_({
    ok: true,
    desde: desde,
    hasta: hasta,
    fichajes: filtered,
    totalHorasDecimal: totals.totalHorasDecimal,
    totalDuracionHHMM: totals.totalDuracionHHMM,
    totalMinutos: totals.totalMinutos,
    cantidadTurnos: filtered.length
  });
}

/**
 * 6. CAMBIAR PIN (Mi cuenta): Permite al usuario cambiar su propio PIN validando el actual.
 */
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
    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      mensaje: 'Tu PIN se actualizó correctamente.'
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 7. EMPLEADOS (Solo dueño): Lista todos los empleados y quién está adentro ahora.
 */
function handleEmpleados_() {
  const users = getAllUsers_();
  const shifts = getAllShifts_();

  const openByUser = {};
  shifts.forEach(function(s) {
    if (s.estado === 'ABIERTO') {
      openByUser[s.usuario] = s;
    }
  });

  const empleados = users.map(function(u) {
    const turnoAbierto = openByUser[u.usuario] || null;
    return {
      usuario: u.usuario,
      nombre: u.nombre,
      rol: u.rol,
      activo: u.activo,
      creado: u.creado,
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

  return jsonResponse_({
    ok: true,
    empleados: empleados,
    adentroAhora: adentroAhora
  });
}

/**
 * 8. CREAR EMPLEADO (Solo dueño): Crea un nuevo usuario con nombre, usuario y PIN.
 */
function handleCrearEmpleado_(body) {
  const nombre = String(body.nombre || '').trim();
  const usuario = normalizeUsername_(body.usuario);
  const pin = String(body.pin || '').trim();
  const rolInput = String(body.rol || 'empleado').trim().toLowerCase();
  const rol = (rolInput === 'dueño' || rolInput === 'dueno') ? 'dueño' : 'empleado';

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
    const existing = findUserByUsername_(usuario);
    if (existing) {
      return errorResponse_('Ya existe un empleado con el usuario "' + usuario + '". Elegí otro.', 'USER_EXISTS');
    }

    const sh = getSheetUsuarios_();
    const pinHash = hashPin_(pin, usuario);
    const creado = formatNowDateTime_();
    const rowIdx = sh.getLastRow() + 1;

    sh.getRange(rowIdx, 1, 1, 6).setNumberFormat('@');
    sh.getRange(rowIdx, 1, 1, 6).setValues([[
      usuario,
      nombre,
      rol,
      pinHash,
      'SI',
      creado
    ]]);
    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      mensaje: 'Empleado "' + nombre + '" (' + usuario + ') creado con éxito.'
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 9. TOGGLE EMPLEADO (Solo dueño): Habilita o deshabilita un empleado.
 */
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
    const target = findUserByUsername_(targetUsername);
    if (!target) {
      return errorResponse_('No se encontró el empleado "' + targetUsername + '".', 'USER_NOT_FOUND');
    }

    const nuevoEstado = !target.activo;
    const sh = getSheetUsuarios_();
    sh.getRange(target.rowNumber, 5).setNumberFormat('@').setValue(nuevoEstado ? 'SI' : 'NO');
    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      activo: nuevoEstado,
      mensaje: 'El empleado "' + target.nombre + '" ahora está ' + (nuevoEstado ? 'HABILITADO' : 'DESHABILITADO') + '.'
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 10. RESET PIN (Solo dueño): Cambia el PIN de cualquier empleado.
 */
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

    // También limpiamos cualquier bloqueo por intentos fallidos que tuviera ese usuario
    const cache = CacheService.getScriptCache();
    cache.remove('LOGIN_LOCK_' + target.usuario);
    cache.remove('LOGIN_FAIL_' + target.usuario);
    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      mensaje: 'Se actualizó el PIN de "' + target.nombre + '" (' + target.usuario + ').'
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 11. EGRESO FORZADO (Solo dueño): Cierra el turno abierto de un empleado,
 * aceptando una hora opcional HH:MM (si se omite, usa la hora actual de Córdoba).
 */
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
    const openShift = findOpenShiftForUser_(sh, targetUsername);
    if (!openShift) {
      return errorResponse_('El empleado "' + targetUsername + '" no tiene ningún turno abierto en este momento.', 'NO_OPEN_SHIFT');
    }

    let egresoMs;
    let horaEgreso;

    if (horaOpcional) {
      horaEgreso = horaOpcional;
      // Construir timestamp en zona horaria Córdoba (-03:00) sobre la fecha del ingreso
      const candidateMs = parseCordobaDateTime_(openShift.fecha, horaOpcional);
      if (candidateMs < openShift.ingresoTS) {
        // Si la hora indicada es menor a la de ingreso (ej. turno nocturno que cruza medianoche), sumar 1 día
        egresoMs = candidateMs + 24 * 60 * 60 * 1000;
      } else {
        egresoMs = candidateMs;
      }
    } else {
      const now = new Date();
      egresoMs = Math.max(now.getTime(), openShift.ingresoTS + 1000);
      horaEgreso = formatTimeHM_(new Date(egresoMs));
    }

    const calc = calculateDuration_(openShift.ingresoTS, egresoMs);

    sh.getRange(openShift.rowNumber, 6, 1, 6).setNumberFormat('@');
    sh.getRange(openShift.rowNumber, 6).setValue(horaEgreso);
    sh.getRange(openShift.rowNumber, 7).setValue(String(calc.horasDecimal));
    sh.getRange(openShift.rowNumber, 8).setValue(calc.duracionHHMM);
    sh.getRange(openShift.rowNumber, 9).setValue('CERRADO');
    sh.getRange(openShift.rowNumber, 11).setValue(String(egresoMs));
    SpreadsheetApp.flush();

    return jsonResponse_({
      ok: true,
      mensaje: 'Egreso forzado aplicado a "' + openShift.nombre + '" a las ' + horaEgreso + ' hs (' + calc.duracionHHMM + ' hs).'
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 12. REPORTE (Solo dueño): Filtra por fechas y empleado (o todos),
 * devuelve resumen por empleado y tabla de fichajes.
 */
function handleReporte_(body) {
  const defaults = getCurrentMonthRange_();
  const desde = normalizeDateFilter_(body.desde, defaults.desde);
  const hasta = normalizeDateFilter_(body.hasta, defaults.hasta);
  const empleadoFiltro = normalizeUsername_(body.usuario || '');

  const users = getAllUsers_();
  const allShifts = getAllShifts_();

  const filtered = allShifts.filter(function(s) {
    const matchDate = s.fecha >= desde && s.fecha <= hasta;
    const matchUser = !empleadoFiltro || empleadoFiltro === 'todos' || s.usuario === empleadoFiltro;
    return matchDate && matchUser;
  });

  // Ordenar por fecha/ingreso descendente
  filtered.sort(function(a, b) {
    return (b.ingresoTS || 0) - (a.ingresoTS || 0);
  });

  // Agrupar resumen por empleado
  const mapResumen = {};
  filtered.forEach(function(s) {
    if (!mapResumen[s.usuario]) {
      mapResumen[s.usuario] = {
        usuario: s.usuario,
        nombre: s.nombre,
        turnosCerrados: 0,
        turnosAbiertos: 0,
        minutosTotales: 0
      };
    }
    const item = mapResumen[s.usuario];
    if (s.estado === 'ABIERTO') {
      item.turnosAbiertos += 1;
    } else {
      item.turnosCerrados += 1;
      item.minutosTotales += durationToMinutes_(s.duracion, s.horasDecimal);
    }
  });

  const resumenPorEmpleado = Object.keys(mapResumen).map(function(key) {
    const r = mapResumen[key];
    const dec = Math.round((r.minutosTotales / 60) * 100) / 100;
    return {
      usuario: r.usuario,
      nombre: r.nombre,
      turnosCerrados: r.turnosCerrados,
      turnosAbiertos: r.turnosAbiertos,
      totalTurnos: r.turnosCerrados + r.turnosAbiertos,
      horasDecimal: dec,
      duracionHHMM: minutesToHHMM_(r.minutosTotales)
    };
  }).sort(function(a, b) {
    return b.horasDecimal - a.horasDecimal;
  });

  const totals = summarizeShifts_(filtered);

  return jsonResponse_({
    ok: true,
    desde: desde,
    hasta: hasta,
    usuarioFiltro: empleadoFiltro || 'todos',
    empleadosLista: users.map(function(u) {
      return { usuario: u.usuario, nombre: u.nombre, activo: u.activo };
    }),
    resumenPorEmpleado: resumenPorEmpleado,
    fichajes: filtered,
    totalGeneralDecimal: totals.totalHorasDecimal,
    totalGeneralHHMM: totals.totalDuracionHHMM,
    cantidadTurnos: filtered.length
  });
}

/* ==========================================================================
 * FUNCIONES AUXILIARES DE DATOS Y HOJAS
 * ========================================================================== */

function ensureSheetsExist_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET_USUARIOS) || !ss.getSheetByName(SHEET_FICHAJES)) {
    setup();
  }
}

function getSheetUsuarios_() {
  ensureSheetsExist_();
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USUARIOS);
}

function getSheetFichajes_() {
  ensureSheetsExist_();
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FICHAJES);
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
    list.push({
      rowNumber: i + 1,
      usuario: usuario,
      nombre: String(row[1] || '').trim(),
      rol: String(row[2] || 'empleado').trim().toLowerCase(),
      pinHash: String(row[3] || '').trim(),
      activo: activo,
      creado: formatCellString_(row[5])
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
  const values = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const id = String(row[0] || '').trim();
    const usuario = normalizeUsername_(row[1]);
    if (!id && !usuario) continue;

    const fecha = normalizeCellDate_(row[3]);
    const ingreso = normalizeCellTime_(row[4]);
    const egreso = normalizeCellTime_(row[5]);
    const horasDecRaw = String(row[6] !== undefined && row[6] !== null ? row[6] : '0').replace(',', '.');
    const horasDecimal = Math.round(( parseFloat(horasDecRaw) || 0 ) * 100) / 100;
    const duracion = normalizeCellDuration_(row[7]);
    const estado = String(row[8] || 'CERRADO').trim().toUpperCase();
    const ingresoTS = Number(row[9]) || parseCordobaDateTime_(fecha, ingreso || '00:00');
    const egresoTS = row[10] ? Number(row[10]) : null;

    list.push({
      rowNumber: i + 1,
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

function findOpenShiftForUser_(sh, username) {
  const target = normalizeUsername_(username);
  const shifts = getAllShifts_();
  for (let i = shifts.length - 1; i >= 0; i--) {
    if (shifts[i].usuario === target && shifts[i].estado === 'ABIERTO') {
      return shifts[i];
    }
  }
  return null;
}

function buildEstadoUsuario_(user) {
  const hoy = formatDateYMD_(new Date());
  const shifts = getAllShifts_();

  let turnoAbierto = null;
  const fichajesHoy = [];

  for (let i = 0; i < shifts.length; i++) {
    const s = shifts[i];
    if (s.usuario === user.usuario) {
      if (s.estado === 'ABIERTO') {
        turnoAbierto = s;
      }
      if (s.fecha === hoy || (s.estado === 'ABIERTO' && s.usuario === user.usuario)) {
        fichajesHoy.push(s);
      }
    }
  }

  fichajesHoy.sort(function(a, b) {
    return (b.ingresoTS || 0) - (a.ingresoTS || 0);
  });

  const totalsHoy = summarizeShifts_(fichajesHoy);

  return {
    enTurno: Boolean(turnoAbierto),
    turnoAbierto: turnoAbierto,
    fechaHoy: hoy,
    serverTimeMs: Date.now(),
    fichajesHoy: fichajesHoy,
    totalHoyDecimal: totalsHoy.totalHorasDecimal,
    totalHoyHHMM: totalsHoy.totalDuracionHHMM
  };
}

/* ==========================================================================
 * CÁLCULOS DE TIEMPO, HASH Y UTILIDADES
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

function calculateDuration_(ingresoTS, egresoTS) {
  const diffMs = Math.max(0, egresoTS - ingresoTS);
  const totalMinutes = Math.max(1, Math.round(diffMs / 60000));
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  const duracionHHMM = pad2_(hours) + ':' + pad2_(minutes);
  const horasDecimal = Math.round((diffMs / 3600000) * 100) / 100;
  return {
    totalMinutes: totalMinutes,
    duracionHHMM: duracionHHMM,
    horasDecimal: horasDecimal
  };
}

function summarizeShifts_(shifts) {
  let totalMinutos = 0;
  shifts.forEach(function(s) {
    if (s.estado === 'CERRADO') {
      totalMinutos += durationToMinutes_(s.duracion, s.horasDecimal);
    }
  });
  const totalHorasDecimal = Math.round((totalMinutos / 60) * 100) / 100;
  return {
    totalMinutos: totalMinutos,
    totalHorasDecimal: totalHorasDecimal,
    totalDuracionHHMM: minutesToHHMM_(totalMinutos)
  };
}

function durationToMinutes_(duracionStr, horasDecFallback) {
  if (duracionStr && /^\d{1,4}:\d{2}$/.test(String(duracionStr).trim())) {
    const parts = String(duracionStr).trim().split(':');
    return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
  }
  return Math.round((Number(horasDecFallback) || 0) * 60);
}

function minutesToHHMM_(totalMinutes) {
  const m = Math.max(0, Math.round(totalMinutes));
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
  // Córdoba (Argentina) usa UTC-3 fijo sin horario de verano
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
