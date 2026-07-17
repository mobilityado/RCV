# REPORT.IA RCV · Login Garantizado

## Qué cambia
La lista de usuarios ya no depende de Internet ni de Apps Script.

Los nombres se cargan directamente desde el proyecto usando la lista de `RCV USUARIOS.xlsx`.

La contraseña sigue validándose contra:

https://script.google.com/macros/s/AKfycbxUhENeMAGaVJx2Gs4yR_qncJJxHyq8NlFFSfa9qu7XBDgcDu4L9HasfrzZSQBOKgwp/exec

## Ventaja
Aunque GitHub Pages no pueda leer la lista desde Google, el desplegable SIEMPRE aparecerá.

## Seguridad
El proyecto solo contiene:
- nombre del usuario;
- tipo/rol.

No contiene contraseñas.

## Apps Script
Para la validación más compatible usa:
`APPS_SCRIPT_LOGIN_GARANTIZADO.gs`

Después publica una nueva versión de la implementación.
