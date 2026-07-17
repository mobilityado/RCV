
const ID_SHEET = '1-ubGZl24R0QMMcF9acK2_6wX4lOZcVPT0OOKTOLT0Mo';

function doGet(e) {
  e = e || { parameter: {} };

  const accion = String(e.parameter.accion || '').trim().toLowerCase();

  if (!accion) {
    return respuesta({
      ok: true,
      mensaje: 'API REPORT.IA RCV activa',
      uso: '?accion=usuarios | ?accion=login&usuario=...&contrasena=...'
    });
  }

  if (accion === 'usuarios') {
    return obtenerUsuarios();
  }

  if (accion === 'login') {
    return validarLogin(
      e.parameter.usuario,
      e.parameter.contrasena
    );
  }

  return respuesta({
    ok: false,
    mensaje: 'Acción inválida'
  });
}

function obtenerUsuarios() {
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

  return respuesta({
    ok: true,
    usuarios: usuarios
  });
}

function validarLogin(usuario, contrasena) {
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
      return respuesta({
        ok: true,
        usuario: datos[i][0],
        tipo: tipo
      });
    }
  }

  return respuesta({
    ok: false,
    mensaje: 'Usuario o contraseña incorrectos'
  });
}

function respuesta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
