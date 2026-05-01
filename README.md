# Documentación de AppScript.js

El script `AppScript.js` es una automatización diseñada para gestionar los niveles de tóner de impresoras leyendo reportes enviados por correo electrónico y actualizando una hoja de cálculo (Google Sheets) que funciona como bitácora.

## Análisis de Funciones

### `CONFIG` (Objeto de Configuración)
Almacena todas las variables importantes que el script necesita para funcionar, como los IDs de las carpetas de Google Drive, los IDs de las hojas de cálculo, el correo para enviar alertas, los umbrales de nivel de tóner para cambiar colores o enviar alertas, y los nombres de las hojas (regiones).

### `procesoCompletoDescargaYActualizacion()`
Es la **función principal o maestra** (programada con un activador de tiempo). Funciona como un director de orquesta:
* Primero llama a la función que descarga los correos.
* Si detecta que se descargaron archivos nuevos, entonces llama a la función que actualiza la bitácora.
* Si no hay nada nuevo, simplemente termina y no hace trabajo innecesario.

### `guardarAdjuntosCSV()`
Se encarga de la extracción de datos:
* Busca en la bandeja de entrada de Gmail correos que vengan de `hp-sds-latam@insightportal.net` y que tengan archivos adjuntos.
* Ignora los correos que ya están marcados con estrella (para no procesar el mismo archivo dos veces).
* Si encuentra archivos `.csv`, los guarda en una carpeta específica de Google Drive y les cambia el nombre agregándoles la fecha y hora.
* Al finalizar, marca el correo con una estrella para saber que ya fue procesado. Devuelve `true` si guardó algo nuevo, o `false` si no hubo correos nuevos.

### `ejecutarProcesoToner()`
Es el motor de actualización de datos. Hace lo siguiente:
* Abre la Bitácora Principal de Google Sheets y la hoja de Auditoría.
* Lee los datos del último archivo CSV que se guardó en Drive.
* Recorre una por una las hojas de las distintas regiones (`CENTRAL`, `MANAGUA`, etc.).
* Por cada impresora en la hoja, busca si hay información nueva en el CSV (usando el número de serie y SKU como llave).
* **Lógica de negocio:**
  * Envía una alerta crítica por correo si el nivel de tóner cayó por debajo del 10%.
  * Actualiza el porcentaje de tóner en la celda correspondiente si este cambió.
  * Actualiza el Número de Serie (S/N) del tóner si detecta que es uno nuevo.
  * Colorea la fila de **Rojo** si el nivel está por debajo del umbral de peligro, o le quita el color si ya se recuperó el nivel (se cambió el tóner).
* Finalmente, guarda un historial de todos estos cambios en la hoja "Cambios_Consolidado".

---

## Error Común: `Exception: Illegal spreadsheet id or key`

Si recibes el error:
`Exception: Illegal spreadsheet id or key: ID_DE_TU_BITACORA_REGIONES`

Este error ocurre **exclusivamente en la línea 6** del script dentro del bloque de `CONFIG`:

```javascript
    ID_BITACORA_PRINCIPAL: "ID_DE_TU_BITACORA_REGIONES",
```

**Causa del problema:**
La función `ejecutarProcesoToner` intenta abrir el archivo de Excel (Google Sheets) usando `SpreadsheetApp.openById(CONFIG.ID_BITACORA_PRINCIPAL)`. El script está intentando buscar un archivo en Google Drive cuyo ID literal sea `"ID_DE_TU_BITACORA_REGIONES"`. Ese texto es solo un **texto de relleno o ejemplo (placeholder)**. Los verdaderos IDs de Google Sheets son una cadena larga de letras y números.

**Cómo solucionarlo:**
1. Abre tu hoja de cálculo "Bitácora de Regiones" en el navegador.
2. Mira la barra de direcciones (URL). Será algo como: `https://docs.google.com/spreadsheets/d/1abc123DEF456ghi789jkl/edit#gid=0`
3. Copia **solo** la parte que está entre `/d/` y `/edit` (en el ejemplo: `1abc123DEF456ghi789jkl`).
4. Ve a tu script y reemplaza el texto de relleno con tu ID real:
   ```javascript
   ID_BITACORA_PRINCIPAL: "1abc123DEF456ghi789jkl", // ¡Pon el tuyo aquí!
   ```
Una vez que hagas ese cambio y guardes, el activador dejará de fallar.
