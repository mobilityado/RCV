
const ID_SHEET = '1-ubGZl24R0QMMcF9acK2_6wX4lOZcVPT0OOKTOLT0Mo';

function doGet(e) {
  const accion = String(e.parameter.accion || '').toLowerCase();
  if (accion === 'login') {
    return login(e.parameter.usuario, e.parameter.contrasena);
  }
  return json({ ok: true, mensaje: 'REPORT.IA RCV API activa' });
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
  return json({ ok: false, mensaje: 'Usuario o contraseña incorrectos.' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
