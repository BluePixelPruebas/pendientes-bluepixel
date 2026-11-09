// Palomas compartidas de la lista de pendientes (https://bluepixelpruebas.github.io/pendientes-bluepixel/).
// Va pegado en Extensiones > Apps Script de una hoja de cálculo del Drive de BluePixel.
// El código de equipo NO se escribe aquí: se guarda en Configuración del proyecto > Propiedades del script,
// con el nombre CODIGO_EQUIPO. Sin ese código, la página puede leer las palomas pero no cambiarlas.

const HOJA_ESTADO = "Estado"
const HOJA_HISTORIAL = "Historial"

function hoja_(nombre, encabezados) {
  const libro = SpreadsheetApp.getActiveSpreadsheet()
  let h = libro.getSheetByName(nombre)
  if (!h) {
    h = libro.insertSheet(nombre)
    h.appendRow(encabezados)
    h.setFrozenRows(1)
  }
  return h
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)
}

// El apóstrofo inicial obliga a Sheets a guardar texto: así «1.10» no se vuelve 1.1 y un nombre no se toma como fórmula.
const texto_ = (v) => "'" + v

function estado_() {
  const h = hoja_(HOJA_ESTADO, ["pendiente", "hecho", "por", "fecha"])
  const filas = h.getDataRange().getValues().slice(1)
  const r = {}
  for (const [n, hecho, por, fecha] of filas) {
    if (hecho === true) r[String(n)] = { por: String(por), fecha: fecha instanceof Date ? fecha.toISOString() : String(fecha) }
  }
  return r
}

function doGet() {
  return json_({ ok: true, estado: estado_() })
}

function doPost(e) {
  let d
  try {
    d = JSON.parse(e.postData.contents)
  } catch (_) {
    return json_({ ok: false, error: "formato" })
  }
  const codigo = PropertiesService.getScriptProperties().getProperty("CODIGO_EQUIPO")
  if (!codigo || String(d.codigo || "") !== codigo) return json_({ ok: false, error: "codigo" })

  const n = String(d.n || "").slice(0, 12)
  const por = String(d.nombre || "").trim().slice(0, 40)
  const hecho = d.hecho === true
  if (!/^[A-Z0-9.]+$/.test(n) || !por) return json_({ ok: false, error: "datos" })

  const candado = LockService.getScriptLock()
  candado.waitLock(10000)
  try {
    const ahora = new Date()
    const h = hoja_(HOJA_ESTADO, ["pendiente", "hecho", "por", "fecha"])
    const ids = h.getRange(1, 1, h.getLastRow(), 1).getValues().map((f) => String(f[0]))
    const i = ids.indexOf(n)
    const fila = [texto_(n), hecho, texto_(por), ahora]
    if (i > 0) h.getRange(i + 1, 1, 1, 4).setValues([fila])
    else h.appendRow(fila)
    hoja_(HOJA_HISTORIAL, ["fecha", "pendiente", "acción", "por"]).appendRow([ahora, texto_(n), hecho ? "marcó" : "desmarcó", texto_(por)])
  } finally {
    candado.releaseLock()
  }
  return json_({ ok: true, estado: estado_() })
}
