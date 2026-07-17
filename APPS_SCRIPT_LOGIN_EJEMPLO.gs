
const ID_SHEET = '1-ubGZl24R0QMMcF9acK2_6wX4lOZcVPT0OOKTOLT0Mo';

function doGet(e) {
  e = e || { parameter: {} };

  const accion = String(e.parameter.accion || '').trim().toLowerCase();
  const callback = String(e.parameter.callback || '').trim();

  let resultado;

  if (!accion) {
    resultado = {
      ok: true,
      mensaje: 'API REPORT.IA RCV activa',
      acciones: ['usuarios', 'login']
    };
    return responder(resultado, callback);
  }

  if (accion === 'usuarios') {
    resultado = obtenerUsuariosDatos();
    return responder(resultado, callback);
  }

  if (accion === 'login') {
    resultado = validarLoginDatos(
      e.parameter.usuario,
      e.parameter.contrasena
    );
    return responder(resultado, callback);
  }

  return responder({
    ok: false,
    mensaje: 'Acción no válida.'
  }, callback);
}

function obtenerUsuariosDatos() {
  try {
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

    return {
      ok: true,
      usuarios: usuarios
    };

  } catch (error) {
    return {
      ok: false,
      mensaje: String(error)
    };
  }
}

function validarLoginDatos(usuario, contrasena) {
  try {
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
        return {
          ok: true,
          usuario: values[i][0],
          tipo: tipo
        };
      }
    }

    return {
      ok: false,
      mensaje: 'Usuario o contraseña incorrectos.'
    };

  } catch (error) {
    return {
      ok: false,
      mensaje: String(error)
    };
  }
}

function responder(objeto, callback) {
  const json = JSON.stringify(objeto);

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
