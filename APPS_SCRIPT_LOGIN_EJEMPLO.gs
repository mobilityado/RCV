
const ID_SHEET = '1-ubGZl24R0QMMcF9acK2_6wX4lOZcVPT0OOKTOLT0Mo';

function doGet(e) {
  // Permite ejecutar doGet manualmente sin que e sea undefined.
  e = e || { parameter: {} };

  const accion = String(e.parameter.accion || '').toLowerCase();

  if (!accion) {
    return json({
      ok: true,
      mensaje: 'API REPORT.IA RCV activa',
      acciones: ['usuarios', 'login']
    });
  }

  if (accion === 'usuarios') {
    return obtenerUsuarios();
  }

  if (accion === 'login') {
    return login(
      e.parameter.usuario,
      e.parameter.contrasena
    );
  }

  return json({
    ok: false,
    mensaje: 'Acción no válida.'
  });
}

function obtenerUsuarios() {
  const ss = SpreadsheetApp.openById(ID_SHEET);
  const sh = ss.getSheets()[0];
  const values = sh.getDataRange().getDisplayValues();

  const usuarios = [];

  for (let i = 1; i < values.length; i++) {
    const usuario = String(values[i][0] || '').trim();
    const tipo = String(values[i][2] || 'USUARIO').trim().toUpperCase();

    if (usuario) {
      usuarios.push({
        usuario: usuario,
        tipo: tipo
      });
    }
  }

  return json({
    ok: true,
    usuarios: usuarios
  });
}

function login(usuario, contrasena) {
  const ss = SpreadsheetApp.openById(ID_SHEET);
  const sh = ss.getSheets()[0];
  const values = sh.getDataRange().getDisplayValues();

  const u = String(usuario || '').trim().toUpperCase();
  const p = String(contrasena || '').trim();

  for (let i = 1; i < values.length; i++) {
    const user = String(values[i][0] || '').trim().toUpperCase();
    const pass = String(values[i][1] || '').trim();
    const tipo = String(values[i][2] || 'USUARIO').trim().toUpperCase();

    if (user === u && pass === p) {
      return json({
        ok: true,
        usuario: values[i][0],
        tipo: tipo
      });
    }
  }

  return json({
    ok: false,
    mensaje: 'Usuario o contraseña incorrectos.'
  });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
