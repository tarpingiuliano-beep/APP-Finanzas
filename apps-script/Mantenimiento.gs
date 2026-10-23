// ==== Limpieza automática de Facturas ====
// Borra de la hoja Facturas las filas que ya cumplieron su ciclo: rescate y pago avisados
// y vencimiento ya pasado. Antes de borrar guarda, por concepto, el último vencimiento cargado
// (propiedad ULTIMA_CARGA), que después sirve para los recordatorios de carga inteligentes.
// Se programa desde Activadores (limpiarFacturas, según tiempo, diario), no desde código,
// para no pedir permisos extra.

function limpiarFacturas() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FACTURAS);
  const datos = hoja.getDataRange().getValues();
  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const props = PropertiesService.getScriptProperties();
  const ultima = JSON.parse(props.getProperty("ULTIMA_CARGA") || "{}");

  let borradas = 0;
  for (let i = datos.length - 1; i >= 1; i--) {      // de abajo hacia arriba para no correr filas
    const f = datos[i];
    if (!f[0]) continue;
    const vto = aFecha(f[3]);
    if (!vto || vto >= hoy) continue;                 // todavía no venció
    if (f[5] === "Pendiente" || f[7] === "Pendiente") continue; // falta algún aviso
    const clave = String(f[0]).trim();
    const guardada = ultima[clave] ? aFecha(ultima[clave]) : null;
    if (!guardada || vto > guardada) ultima[clave] = Utilities.formatDate(vto, Session.getScriptTimeZone(), "yyyy-MM-dd");
    hoja.deleteRow(i + 1);
    borradas++;
  }
  props.setProperty("ULTIMA_CARGA", JSON.stringify(ultima));
  return borradas;
}
