/**
 * @file importarEventosConImporte.gs
 * @description Script de Google Apps Script que sincroniza eventos del calendario de Google
 * con una hoja de cálculo de Google Sheets, organizados por mes.
 *
 * Funcionamiento general:
 * - Lee todos los eventos del año en curso del calendario indicado.
 * - Filtra únicamente los eventos de color 11 (rojo tomate) o 4 (flamingo),
 * que representan pagos o vencimientos.
 * - Para cada mes, crea o actualiza una hoja con las columnas:
 * FECHA | DESCRIPCIÓN | VENCIMIENTO | IMPORTE (€) | ESTADO | ID (oculta)
 * - Evita duplicados usando el ID nativo del evento de Google Calendar (columna F, oculta).
 * - Genera una tabla resumen (columnas G-I) con el total a pagar por día
 * y el total acumulado de eventos marcados como PAGADO.
 * 
 * @version 3.2
 */


// ---------------------------------------------------------------------------
// CONSTANTES DE CONFIGURACIÓN
// ---------------------------------------------------------------------------

/** ID del calendario de Google desde el que se importan los eventos. */
var CALENDAR_ID = 'bimotor8@gmail.com';

/** Colores de evento que se importan (11 = rojo tomate, 4 = flamingo). */
var COLORES_VALIDOS = ["11", "4"];

/** Nombres de los meses en español, en orden. Determinan el nombre de cada hoja. */
var NOMBRES_MESES = [
  "enero", "febrero", "marzo", "abril", "mayo", "junio",
  "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
];

/** Índice (base 1) de cada columna en la hoja de datos. */
var COL = {
  FECHA:       1,
  DESCRIPCION: 2,
  VENCIMIENTO: 3,
  IMPORTE:     4,
  ESTADO:      5,
  ID_EVENTO:   6  // Columna auxiliar oculta con el ID nativo del evento de Google Calendar
};

/** Índice (base 1) de cada columna en la tabla resumen (columnas G-I). */
var COL_RESUMEN = {
  FECHA_PAGO:   7,
  TOTAL_DIA:    8,
  TOTAL_PAGADO: 9
};


// ---------------------------------------------------------------------------
// FUNCIÓN PRINCIPAL
// ---------------------------------------------------------------------------

/**
 * Importa y sincroniza los eventos del calendario con la hoja de cálculo.
 *
 * Itera sobre los 12 meses del año en curso. Por cada mes:
 * 1. Obtiene o crea la hoja correspondiente.
 * 2. Lee los IDs de eventos ya existentes (columna F) para evitar duplicados.
 * 3. Filtra los eventos del calendario por mes y color.
 * 4. Inserta filas nuevas o actualiza las existentes según corresponda.
 * 5. Aplica formato condicional al estado (PAGADO / SIN PAGAR).
 * 6. Genera la tabla de totales diarios y el total pagado en columnas G-I.
 */
function importarEventosConImporte() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cal = CalendarApp.getCalendarById(CALENDAR_ID);
  var tz  = Session.getScriptTimeZone(); // Zona horaria del script, evita hardcodear GMT+1

  if (!cal) {
    Logger.log("Error: no se encontró el calendario con ID: " + CALENDAR_ID);
    return;
  }

  // Rango completo del año en curso
  var añoActual = new Date().getFullYear();
  var inicio    = new Date(añoActual, 0, 1);
  var fin       = new Date(añoActual, 11, 31, 23, 59, 59); // 23:59:59 para incluir el 31/12

  var eventos = cal.getEvents(inicio, fin);

  NOMBRES_MESES.forEach(function(mes, indexMes) {
    var hojaMes = obtenerOCrearHoja(ss, mes);
    inicializarCabecera(hojaMes);

    var idsExistentes = leerIdsExistentes(hojaMes);
    var eventosDelMes = filtrarEventosPorMesYColor(eventos, indexMes, COLORES_VALIDOS);

    sincronizarEventos(hojaMes, eventosDelMes, idsExistentes, tz);
    
    // ORDENAR POR FECHA: Evita que los eventos nuevos se queden al final de la lista
    var ultimaFila = hojaMes.getLastRow();
    if (ultimaFila > 1) {
      hojaMes.getRange(2, 1, ultimaFila - 1, COL.ID_EVENTO).sort({column: 1, ascending: true});
    }

    aplicarFormatoCondicional(hojaMes);
    generarTablaResumen(hojaMes, tz);

    // Anchos de columna para mejor legibilidad
    hojaMes.setColumnWidth(COL.DESCRIPCION,          350);
    hojaMes.setColumnWidth(COL.ESTADO,               120);
    hojaMes.setColumnWidth(COL_RESUMEN.FECHA_PAGO,   100);
    hojaMes.setColumnWidth(COL_RESUMEN.TOTAL_PAGADO, 120);

    // Ocultamos la columna F (ID de evento), es auxiliar y no debe editarse manualmente
    hojaMes.hideColumns(COL.ID_EVENTO);
  });
}


// ---------------------------------------------------------------------------
// FUNCIONES AUXILIARES
// ---------------------------------------------------------------------------

/**
 * Devuelve la hoja con el nombre indicado, o la crea si no existe.
 *
 * @param {Spreadsheet} ss     - Hoja de cálculo activa.
 * @param {string}      nombre - Nombre de la hoja a obtener o crear.
 * @returns {Sheet} La hoja encontrada o recién creada.
 */
function obtenerOCrearHoja(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}


/**
 * Escribe la fila de cabecera en la hoja si aún no tiene ningún contenido.
 * No sobreescribe si la hoja ya tiene datos.
 *
 * @param {Sheet} hoja - Hoja sobre la que operar.
 */
function inicializarCabecera(hoja) {
  if (hoja.getLastRow() === 0) {
    hoja.getRange("A1:F1")
      .setValues([['FECHA', 'DESCRIPCIÓN', 'VENCIMIENTO', 'IMPORTE (€)', 'ESTADO', 'ID_EVENTO']])
      .setFontWeight("bold")
      .setBackground("#f3f3f3");
  }
}


/**
 * Lee la columna F (ID de evento de Google Calendar) de todas las filas existentes
 * y devuelve un mapa { idEvento -> númeroDeFila } para detección rápida de duplicados.
 *
 * Usar el ID nativo del evento (en lugar de construir claves artificiales con fecha+título)
 * garantiza unicidad aunque dos eventos compartan fecha y título.
 *
 * @param {Sheet} hoja - Hoja de la que leer los IDs.
 * @returns {Object} Mapa de ID de evento a número de fila (base 1).
 */
function leerIdsExistentes(hoja) {
  var idsExistentes = {};
  var ultimaFila = hoja.getLastRow();

  if (ultimaFila > 1) {
    var datos = hoja.getRange(2, COL.ID_EVENTO, ultimaFila - 1, 1).getValues();
    datos.forEach(function(fila, i) {
      if (fila[0]) {
        idsExistentes[fila[0].toString()] = i + 2; // +2 porque la fila 1 es cabecera
      }
    });
  }

  return idsExistentes;
}


/**
 * Filtra una lista de eventos de calendario quedándose únicamente con los que:
 * - Pertenecen al mes indicado (indexMes: 0 = enero, 11 = diciembre).
 * - Tienen un color incluido en la lista de colores válidos.
 *
 * @param {CalendarEvent[]} eventos        - Lista completa de eventos del año.
 * @param {number}          indexMes       - Índice del mes a filtrar (0-11).
 * @param {string[]}        coloresValidos - Array de códigos de color aceptados (ej. ["11","4"]).
 * @returns {CalendarEvent[]} Subconjunto de eventos que cumplen los criterios.
 */
function filtrarEventosPorMesYColor(eventos, indexMes, coloresValidos) {
  return eventos.filter(function(ev) {
    var color = ev.getColor();
    return ev.getStartTime().getMonth() === indexMes && coloresValidos.indexOf(color) !== -1;
  });
}


/**
 * Extrae de un título de evento el importe numérico y la fecha de vencimiento,
 * usando expresiones regulares.
 *
 * - Importe: primer número con decimales encontrado (ej. "123,45" o "123.45").
 * - Vencimiento: primer patrón DD/MM encontrado (ej. "15/03").
 *
 * @param {string} titulo - Título del evento de calendario.
 * @returns {{ importeNum: number, vencimiento: string }}
 * importeNum:  valor numérico del importe (0 si no se encontró).
 * vencimiento: texto con formato "DD/MM" precedido de apóstrofo para evitar
 * que Sheets lo interprete como fecha, o "-" si no se encontró.
 */
function extraerDatosDelTitulo(titulo) {
  var matchImporte     = titulo.match(/(\d+[\.,]\d+)/);
  var importeNum       = matchImporte ? parseFloat(matchImporte[0].replace(',', '.')) : 0;

  var matchVencimiento = titulo.match(/(\d+\/\d+)/);
  var vencimiento      = matchVencimiento ? "'" + matchVencimiento[0] : "-";

  return { importeNum: importeNum, vencimiento: vencimiento };
}


/**
 * Recorre los eventos filtrados del mes y los sincroniza con la hoja:
 * - Si el ID del evento ya existe en la hoja (columna F), actualiza vencimiento e importe.
 * - Si no existe, inserta una fila nueva con todos los datos y el ID en columna F.
 *
 * Usa un contador de fila local (en lugar de llamar a getLastRow() en cada iteración)
 * para minimizar llamadas a la API de Sheets y mejorar el rendimiento.
 *
 * @param {Sheet}           hoja          - Hoja del mes sobre la que operar.
 * @param {CalendarEvent[]} eventos       - Eventos filtrados del mes.
 * @param {Object}          idsExistentes - Mapa { idEvento -> númeroDeFila } de filas ya presentes.
 * @param {string}          tz            - Zona horaria del script (ej. "Europe/Madrid").
 */
function sincronizarEventos(hoja, eventos, idsExistentes, tz) {
  var contadorFila = hoja.getLastRow(); // Contador local para evitar getLastRow() en bucle

  eventos.forEach(function(ev) {
    var titulo     = ev.getTitle();
    var fechaTexto = Utilities.formatDate(ev.getStartTime(), tz, "dd/MM/yyyy");
    var idEvento   = ev.getId(); // ID nativo y único del evento en Google Calendar
    var datos      = extraerDatosDelTitulo(titulo);

    if (idsExistentes[idEvento]) {
      // El evento ya existe: solo actualizamos vencimiento e importe por si han cambiado
      var filaDestino = idsExistentes[idEvento];
      hoja.getRange(filaDestino, COL.VENCIMIENTO).setValue(datos.vencimiento);
      hoja.getRange(filaDestino, COL.IMPORTE)
          .setValue(datos.importeNum)
          .setNumberFormat("#,##0.00\" €\"");
    } else {
      // Evento nuevo: insertamos fila completa
      contadorFila++;
      hoja.getRange(contadorFila, 1, 1, 6).setValues([[
        fechaTexto,
        titulo,
        datos.vencimiento,
        datos.importeNum,
        "SIN PAGAR",
        idEvento
      ]]);
      // Texto plano en fechas para evitar conversión automática de Sheets
      hoja.getRange(contadorFila, COL.FECHA).setNumberFormat("@");
      hoja.getRange(contadorFila, COL.VENCIMIENTO).setNumberFormat("@");
      hoja.getRange(contadorFila, COL.IMPORTE).setNumberFormat("#,##0.00\" €\"");

      // Desplegable de validación en la columna ESTADO
      var regla = SpreadsheetApp.newDataValidation()
        .requireValueInList(['SIN PAGAR', 'PAGADO'], true).build();
      hoja.getRange(contadorFila, COL.ESTADO).setDataValidation(regla);
    }
  });
}


/**
 * Aplica (o reaplica) las reglas de formato condicional a la columna ESTADO:
 * - "SIN PAGAR" → fondo rojo claro (#ffcfc9).
 * - "PAGADO"    → fondo verde claro (#d4edbc).
 *
 * Limpia las reglas existentes antes de aplicar las nuevas para evitar acumulación
 * en cada ejecución del script.
 *
 * @param {Sheet} hoja - Hoja sobre la que aplicar el formato condicional.
 */
function aplicarFormatoCondicional(hoja) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) return;

  var rangoEstado = hoja.getRange(2, COL.ESTADO, ultimaFila - 1, 1);
  hoja.clearConditionalFormatRules();
  hoja.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('SIN PAGAR').setBackground('#ffcfc9').setRanges([rangoEstado]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('PAGADO').setBackground('#d4edbc').setRanges([rangoEstado]).build()
  ]);
}


/**
 * Genera la tabla resumen en las columnas G-I de la hoja:
 * - Columna G: fechas únicas de pago.
 * - Columna H: suma de importes de todos los eventos de esa fecha (total a pagar ese día).
 * - Columna I (solo fila 2): suma acumulada de importes con estado "PAGADO" en el mes.
 * - Columna I (fila 4): suma total prevista del mes (todos los importes, independiente del estado).
 *
 * La tabla se regenera completamente en cada ejecución para reflejar cambios de estado.
 *
 * @param {Sheet}  hoja - Hoja del mes sobre la que generar el resumen.
 * @param {string} tz   - Zona horaria del script para formatear fechas correctamente.
 */
function generarTablaResumen(hoja, tz) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) return;

  // Limpiamos la zona de resumen antes de regenerarla
  hoja.getRange("G:I").clearContent().clearFormat();
  hoja.getRange("G1:I1")
    .setValues([["FECHA PAGO", "TOTAL DÍA", "TOTAL PAGADO"]])
    .setBackground("#d9ead3")
    .setFontWeight("bold");

  // Leemos columnas A-E (ignoramos F que es el ID interno)
  var datos = hoja.getRange(2, 1, ultimaFila - 1, 5).getValues();

  var totalesDiarios = {};
  var totalPagado    = 0;
  var totalPrevisto  = 0; // Inicializado para sumar correctamente

  datos.forEach(function(fila) {
    var celdaFecha = fila[COL.FECHA - 1];
    var importe    = parseFloat(fila[COL.IMPORTE - 1]) || 0;
    var estado     = fila[COL.ESTADO - 1];

    if (!celdaFecha || !importe) return;

    // Sumamos al total previsto general del mes
    totalPrevisto += importe;

    // Normalizamos la clave de fecha siempre a texto formateado
    var fKey = (celdaFecha instanceof Date)
      ? Utilities.formatDate(celdaFecha, tz, "dd/MM/yyyy")
      : celdaFecha.toString();

    // Total del día: suma todos los eventos del día independientemente del estado
    totalesDiarios[fKey] = (totalesDiarios[fKey] || 0) + importe;

    // Total pagado del mes
    if (estado === "PAGADO") {
      totalPagado += importe;
    }
  });

  var fechasArr = Object.keys(totalesDiarios);
  if (fechasArr.length === 0) return;

  // Escribimos la tabla de totales diarios en G-H
  var filasTotales = fechasArr.map(function(f) { return [f, totalesDiarios[f]]; });
  hoja.getRange(2, COL_RESUMEN.FECHA_PAGO, filasTotales.length, 1).setNumberFormat("@");
  hoja.getRange(2, COL_RESUMEN.FECHA_PAGO, filasTotales.length, 2).setValues(filasTotales);
  hoja.getRange(2, COL_RESUMEN.TOTAL_DIA,  filasTotales.length, 1).setNumberFormat("#,##0.00\" €\"");

  // Total pagado del mes en la celda I2, con fondo verde para distinguirlo
  hoja.getRange(2, COL_RESUMEN.TOTAL_PAGADO)
    .setValue(totalPagado)
    .setNumberFormat("#,##0.00\" €\"")
    .setBackground("#d4edbc")
    .setFontWeight("bold");

  // Total Previsto (Suma total del mes)
  hoja.getRange(3, COL_RESUMEN.TOTAL_PAGADO).setValue("TOTAL PREVISTO").setFontWeight("bold").setBackground("#cfe2f3");
  hoja.getRange(4, COL_RESUMEN.TOTAL_PAGADO)
    .setValue(totalPrevisto)
    .setNumberFormat("#,##0.00\" €\"")
    .setFontWeight("bold");
}
