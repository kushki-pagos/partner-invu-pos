/**
 * Leads - Form InvuPOS
 * Recibe los envíos de la landing y los agrega como una fila nueva del Sheet.
 *
 * Cómo publicarlo:
 *   1. En el Sheet: Extensiones → Apps Script
 *   2. Pega este archivo completo (reemplaza el "function myFunction()" que viene)
 *   3. Guardar
 *   4. Implementar → Nueva implementación → Tipo: Aplicación web
 *        · Ejecutar como:        Yo (tu cuenta)
 *        · Quién tiene acceso:   Cualquier persona     ← importante
 *   5. Copia la URL que termina en /exec y pégala en index.html,
 *      en la constante APPS_SCRIPT_URL
 *
 * Ojo: cada vez que edites este script hay que volver a implementar
 * (Implementar → Administrar implementaciones → editar → Versión: Nueva).
 */

// Marcador para saber desde fuera qué versión está implementada.
// Súbelo cada vez que edites el script: si abres la URL /exec y la
// "version" que responde no es esta, la reimplementación no tomó.
const VERSION = 'v3';

// Pestaña donde se escriben los leads. Si no existe, usa la primera.
const HOJA = 'Hoja 1';

// Encabezado del Sheet  →  campo que manda el formulario.
// null = la llena el script (fecha del servidor).
// El orden de las columnas lo decide el Sheet, no este objeto:
// si mueves o agregas columnas, el script se acomoda solo.
const COLUMNAS = {
  'FECHA':             null,
  'NOMBRE':            'nombre',
  'CORREO':            'email',
  'TELEFONO':          'telefono',
  'EMPRESA':           'empresa',
  'CANAL DE CONTACTO': 'canal',
  'ORIGEN':            'origen'   // opcional, por si luego agregas la columna
};


function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);   // evita que dos envíos simultáneos se pisen

  try {
    const libro = SpreadsheetApp.getActiveSpreadsheet();
    const hoja  = libro.getSheetByName(HOJA) || libro.getSheets()[0];

    const encabezados = hoja
      .getRange(1, 1, 1, hoja.getLastColumn())
      .getValues()[0]
      .map(normalizar);

    const datos = (e && e.parameter) ? e.parameter : {};

    const fila = encabezados.map(function (encabezado) {
      const campo = COLUMNAS[encabezado];
      if (campo === null)      return new Date();   // FECHA
      if (campo === undefined) return '';           // columna que no conocemos
      return textoSeguro(datos[campo] || '');
    });

    // appendRow (y no setValues) es lo que respeta el apóstrofo inicial
    // que textoSeguro() agrega para forzar texto plano.
    hoja.appendRow(fila);
    return responder({ ok: true });

  } catch (err) {
    return responder({ ok: false, error: String(err) });

  } finally {
    lock.releaseLock();
  }
}


// Abre la URL /exec en el navegador para comprobar dos cosas:
// que está publicada, y QUÉ VERSIÓN del código está corriendo.
// Si "version" no coincide con la constante de arriba, la
// reimplementación no tomó (guardar no basta).
function doGet() {
  return responder({ ok: true, version: VERSION, mensaje: 'Endpoint activo' });
}


/**
 * Sheets interpreta como fórmula cualquier valor que empiece con
 * = + - @ , y el formato de celda no lo evita: setValues escribe
 * como si lo tecleara una persona. Por eso "+52 55 1234 5678"
 * llegaba con error.
 *
 * El apóstrofo inicial fuerza texto. No se guarda como parte del
 * valor ni se muestra en la celda: al leerla vuelve "+52 55 1234 5678".
 */
function textoSeguro(valor) {
  const texto = String(valor);
  return /^[=+\-@]/.test(texto) ? "'" + texto : texto;
}


// Compara encabezados sin importar mayúsculas, espacios ni acentos,
// así "TELEFONO" y "Teléfono" apuntan a la misma columna.
function normalizar(texto) {
  return String(texto)
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}


function responder(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}
