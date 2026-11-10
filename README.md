# Control de Gastos

App web personal de control de gastos (Google Apps Script + Google Sheets).

La carpeta `app/` tiene la versión publicada. La app instalada revisa `app/version.json` todas las noches y, si hay una versión nueva, baja `Code.gs`, `Index.html` y `appsscript.json` y se actualiza sola.

Este repositorio solo tiene código: los datos (movimientos, cuentas, saldos) y las claves quedan en el Google Sheet de cada usuario.
