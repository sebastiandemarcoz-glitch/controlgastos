# Control de Gastos

App web personal de control de gastos (Google Apps Script + Google Sheets).

- `app/Cargador.gs`: el único archivo que va en Apps Script. Baja la app de esta carpeta y la guarda en la pestaña oculta `_app` de la planilla.
- `app/Code.gs` e `app/Index.html`: la lógica y la pantalla de la app, que se actualizan solas.
- `app/version.json`: la versión publicada. La app la revisa todas las noches.
- `android/`: la app "Gastos · Avisos" para Android (SMS y notificaciones). La compila GitHub Actions y se publica en Releases.

Este repositorio solo tiene código: los datos y las claves quedan en el Google Sheet de cada usuario.

## Cómo se publica una versión
Las fuentes están en `src/`. `python3 src/build.py app` (desde `src/`, con salida en `../app`) arma `app/Index.html` (página chica) y `app/Code.gs` (lógica + código de la pantalla en `PANTALLA_JS`). Después se actualiza `app/version.json` (tiene que coincidir con `VERSION_APP`).
