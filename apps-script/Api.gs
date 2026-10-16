// ==== API JSON de la app (fase 1: inicio y deuda) ====
// La contraseña vive en Propiedades del script (Configuración del proyecto → Propiedades del script),
// con el nombre APP_PASSWORD. No va escrita en el código.
//
// En el proyecto actual hay que BORRAR: doGet, verificarClave y generarDashboardHtml (la web app vieja),
// y la constante PASSWORD. Este archivo las reemplaza.

const MESES_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
const MESES_LARGO = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

function doGet() {
  return ContentService.createTextOutput("API Finanzas");
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const esperada = PropertiesService.getScriptProperties().getProperty("APP_PASSWORD");
    if (!esperada) return respuesta({ ok: false, error: "Falta APP_PASSWORD en las propiedades del script" });
    if (String(body.clave || "") !== esperada) {
      Utilities.sleep(1000);
      return respuesta({ ok: false, error: "Contraseña incorrecta" });
    }
    if (body.accion === "inicio") return respuesta({ ok: true, data: apiInicio() });
    if (body.accion === "deuda") return respuesta({ ok: true, data: apiDeuda() });
    return respuesta({ ok: false, error: "Acción desconocida" });
  } catch (err) {
    return respuesta({ ok: false, error: "Error: " + err.message });
  }
}

function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Convierte una celda (Date o texto "yyyy-MM-dd") en Date local sin corrimiento de zona horaria.
function aFecha(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function apiInicio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const f = ss.getSheetByName("Flujo").getRange("A1:B6").getValues();
  const mesRaw = f[0][1];
  const mes = (mesRaw instanceof Date) ? MESES_LARGO[mesRaw.getMonth()] + " " + mesRaw.getFullYear() : String(mesRaw);

  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const datos = ss.getSheetByName(SHEET_FACTURAS).getDataRange().getValues();
  const vencimientos = [];
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (!fila[0] || fila[7] === "Pagado") continue;
    const vto = aFecha(fila[3]);
    if (!vto || vto < hoy) continue;
    vencimientos.push({
      concepto: fila[0], monto: Number(fila[2]) || 0,
      fecha: Utilities.formatDate(vto, tz, "dd/MM/yyyy"),
      dias: Math.round((vto - hoy) / 86400000), _t: vto.getTime()
    });
  }
  vencimientos.sort((a, b) => a._t - b._t);
  vencimientos.forEach(v => delete v._t);

  return {
    mes: mes, ingreso: f[1][1], fijos: f[2][1], cuotas: f[3][1],
    pagoUnico: f[4][1] || 0, resultado: f[5][1],
    vencimientos: vencimientos.slice(0, 10)
  };
}

function apiDeuda() {
  const datos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CUOTAS).getDataRange().getValues();
  const cuotas = [];
  for (let i = 1; i < datos.length; i++) {
    const nombre = String(datos[i][0]).trim();
    if (nombre.toLowerCase() === "total") break;
    const monto = Number(datos[i][1]) || 0;
    const restantes = Number(datos[i][2]) || 0;
    if (!nombre || monto <= 0 || restantes <= 0) continue;
    cuotas.push({ concepto: nombre, monto: monto, restantes: restantes, origen: datos[i][4] || "" });
  }

  const hoy = new Date();
  const meses = [];
  for (let k = 0; k < 12; k++) {
    const d = new Date(hoy.getFullYear(), hoy.getMonth() + k, 1);
    const total = cuotas.reduce((s, c) => s + (c.restantes > k ? c.monto : 0), 0);
    meses.push({ label: MESES_ES[d.getMonth()] + " " + String(d.getFullYear()).slice(2), total: total });
  }
  const total = cuotas.reduce((s, c) => s + c.monto * c.restantes, 0);
  return { total: total, meses: meses, cuotas: cuotas };
}
