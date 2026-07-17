
const ID_SHEET = '1-ubGZl24R0QMMcF9acK2_6wX4lOZcVPT0OOKTOLT0Mo';

function doGet(e) {
  e = e || { parameter: {} };

  const accion = String(e.parameter.accion || '').trim().toLowerCase();
  const callback = String(e.parameter.callback || '').trim();

  let resultado;

  if (!accion) {
    resultado = {
      ok: true,
      mensaje: 'API REPORT.IA RCV activa'
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
    mensaje: 'Acción inválida'
  }, callback);
}

function obtenerUsuariosDatos() {
  const ss = SpreadsheetApp.openById(ID_SHEET);
  const sh = ss.getSheets()[0];
  const datos = sh.getDataRange().getDisplayValues();

  const usuarios = [];

  for (let i = 1; i < datos.length; i++) {
    const usuario = String(datos[i][0] || '').trim();
    const tipo = String(datos[i][2] || 'USUARIO').trim().toUpperCase();

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
}

function validarLoginDatos(usuario, contrasena) {
  const ss = SpreadsheetApp.openById(ID_SHEET);
  const sh = ss.getSheets()[0];
  const datos = sh.getDataRange().getDisplayValues();

  const usuarioBuscado = String(usuario || '').trim().toUpperCase();
  const passwordBuscado = String(contrasena || '').trim();

  for (let i = 1; i < datos.length; i++) {
    const usuarioHoja = String(datos[i][0] || '').trim().toUpperCase();
    const passwordHoja = String(datos[i][1] || '').trim();
    const tipo = String(datos[i][2] || 'USUARIO').trim().toUpperCase();

    if (usuarioHoja === usuarioBuscado && passwordHoja === passwordBuscado) {
      return {
        ok: true,
        usuario: datos[i][0],
        tipo: tipo
      };
    }
  }

  return {
    ok: false,
    mensaje: 'Usuario o contraseña incorrectos'
  };
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
