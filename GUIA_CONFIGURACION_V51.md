# REPORT.IA RCV v51 · Sesión persistente + Push por Gerencia

## Qué ya queda programado

- La sesión se guarda en `localStorage` y el backend conserva un token persistente en Google Sheets. Cerrar la app, navegador, celular o PC NO cierra la sesión.
- La sesión se invalida al pulsar **Cerrar sesión** o al cambiar la contraseña. También se invalida automáticamente si alguien modifica la contraseña directamente en la hoja de usuarios.
- **Mi perfil** muestra usuario, región, gerencia, estado de notificaciones y cambio de contraseña.
- En las cuentas contables, el administrador puede elegir **GERENCIA DESTINO** antes de enviar una observación.
- Los usuarios reciben en el Centro de Notificaciones únicamente los mensajes compatibles con su región y gerencia.
- El administrador recibe confirmación de la gerencia destino y cantidad de dispositivos alcanzados.
- La PWA incluye `manifest.webmanifest`, iconos 192/512 y `firebase-messaging-sw.js`.
- Cuando Firebase esté configurado, los avisos pueden aparecer con la app cerrada.

## 1. PREPARAR LA HOJA DE USUARIOS

La hoja que ya usa REPORT.IA conserva sus primeras columnas. Agrega la columna E:

A = USUARIO
B = CONTRASEÑA
C = TIPO
D = REGIÓN
E = GERENCIA

Ejemplo:
`JOSE | contraseña | USUARIO | VILLAHERMOSA | GERENCIA ADMINISTRATIVA`

Para ADMINISTRADOR puedes dejar E vacía; REPORT.IA usará `TODAS LAS GERENCIAS`.

IMPORTANTE: escribe la gerencia exactamente como quieras identificarla. Los mensajes se agrupan por ese valor.

## 2. ACTUALIZAR GOOGLE APPS SCRIPT

1. Abre el proyecto de Apps Script que actualmente usa REPORT.IA.
2. Haz una copia del código actual por seguridad.
3. Reemplaza el contenido por `codigo-google-RCV-NUBE-v51.txt`.
4. Guarda.
5. Implementar > Administrar implementaciones > editar la implementación Web existente > Nueva versión > Implementar.
6. Si conservas la misma implementación, normalmente la URL `/exec` no cambia. Si cambia, actualiza `API_URL` en `config.js`.

Al primer uso el backend crea automáticamente:
- `V51_SESIONES_PERSISTENTES`
- `V51_PUSH_DEVICES`
- `V51_PUSH_LOG`

También amplía `V38_NOTAS_INCIDENCIAS` con la columna `GERENCIA_DESTINO`.

## 3. PRUEBA PRIMERO LA SESIÓN PERSISTENTE (SIN FIREBASE)

Antes de configurar push:
1. Sube esta v51 a GitHub Pages.
2. Inicia sesión como un usuario.
3. Cierra completamente la PWA/navegador y vuelve a abrir.
4. Debe entrar directamente con el mismo usuario.
5. Pulsa Cerrar sesión: ahora sí debe pedir acceso de nuevo.
6. En Mi perfil cambia la contraseña. Debe cerrar la sesión y pedir la nueva contraseña.

## 4. CREAR FIREBASE PARA PUSH

1. Entra a Firebase Console y crea un proyecto, por ejemplo `reportia-rcv`.
2. Dentro del proyecto: **Project settings / Configuración del proyecto > General > Tus apps > Web (</>)**.
3. Registra una app Web, por ejemplo `REPORT.IA RCV`. No necesitas mover el hosting de GitHub Pages.
4. Firebase mostrará un objeto `firebaseConfig`. Copia sus valores a `firebase-config.js`.
5. Ve a **Project settings > Cloud Messaging > Web Push certificates** y genera un par de claves.
6. Copia la clave pública al valor `REPORTIA_VAPID_KEY` en `firebase-config.js`.

## 5. CREDENCIALES DEL SERVIDOR FIREBASE

Para que Apps Script pueda ENVIAR las notificaciones necesita una cuenta de servicio:
1. Firebase > Project settings > Service accounts.
2. Genera una nueva clave privada JSON. NO subas ese JSON a GitHub.
3. Del JSON necesitas `project_id`, `client_email` y `private_key`.
4. En Apps Script abre **Project Settings / Configuración del proyecto > Script properties / Propiedades del script** y crea:
   - `FCM_PROJECT_ID` = valor de `project_id`
   - `FCM_CLIENT_EMAIL` = valor de `client_email`
   - `FCM_PRIVATE_KEY` = valor COMPLETO de `private_key`, incluyendo BEGIN/END PRIVATE KEY.

La clave privada queda solo en Apps Script; jamás va en los archivos públicos de GitHub.

## 6. SUBIR LOS ARCHIVOS A GITHUB

Sube la carpeta v51 completa, en especial:
- `index.html`
- `auth-v51.js`
- `push-v51.js`
- `v51.js`
- `v51.css`
- `firebase-config.js` ya rellenado
- `firebase-messaging-sw.js`
- `manifest.webmanifest`
- `icon-192.png`
- `icon-512.png`

NO subas el JSON de la cuenta de servicio.

## 7. ACTIVAR NOTIFICACIONES EN CADA DISPOSITIVO

El usuario inicia sesión > **Mi perfil > Activar notificaciones**. El navegador preguntará permiso. Debe elegir Permitir. El dispositivo queda asociado a USUARIO + GERENCIA.

Si José tiene celular y PC, puede activarlas en ambos y ambos recibirán los avisos. Si pulsa Cerrar sesión se desactiva el dispositivo de esa sesión. Si cambia contraseña se desactivan todos sus dispositivos hasta volver a entrar y habilitarlos.

## 8. PRUEBA FINAL

1. Inicia como administrador.
2. Abre Gastos o Costos > una cuenta contable.
3. En Seguimiento selecciona **GERENCIA DESTINO**.
4. Escribe el mensaje y pulsa **Enviar a gerencia**.
5. La pantalla debe confirmar algo como `Mensaje enviado a GERENCIA X · 2 dispositivo(s) notificado(s).`
6. En un celular perteneciente a esa gerencia, cierra REPORT.IA y bloquea la pantalla.
7. Debe llegar la notificación del sistema. Al tocarla abrirá REPORT.IA.

## Observación importante para iPhone/iPad

En iOS/iPadOS el comportamiento de Web Push requiere usar la web como app instalada en la pantalla de inicio y conceder permiso de notificaciones desde esa PWA. En Android/Chrome y navegadores de escritorio compatibles, el flujo suele ser más directo.
