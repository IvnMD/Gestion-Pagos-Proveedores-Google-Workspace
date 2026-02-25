# Changelog

Historial de cambios del proyecto **Google Calendar → Sheets: Gestión de Pagos a Proveedores**.

---

## [3.1] - Corrección de duplicados

### Problema detectado
Al ejecutar el script más de una vez, los eventos ya existentes en la hoja se volvían a insertar como filas nuevas, generando duplicados y descontrolando los totales.

### Causa raíz
El script identificaba los eventos existentes construyendo una clave artificial con `fecha + "_" + titulo`. Si el evento había sido editado en Google Calendar (corrección de una falta de ortografía, cambio de hora, etc.) la clave no coincidía y el evento se insertaba de nuevo.

### Solución
Se sustituyó el sistema de claves artificiales por el uso de `ev.getId()`, el identificador único y permanente que Google Calendar asigna a cada evento. Este ID se almacena en la **columna F** (oculta) de la hoja, y es la referencia que usa el script para decidir si un evento ya existe o es nuevo.

**Ventaja clave:** el ID de un evento no cambia aunque se modifique su título, fecha, hora o color. Esto hace que la sincronización sea robusta frente a cualquier edición del evento en Calendar.

### Otros cambios técnicos incluidos en esta versión
- **Zona horaria dinámica:** se sustituye `"GMT+1"` hardcodeado por `Session.getScriptTimeZone()`, evitando desfases de fecha en entornos con configuración diferente.
- **Rendimiento:** se elimina la llamada a `getLastRow()` dentro del bucle de inserción, sustituyéndola por un contador local que se incrementa manualmente.
- **Consistencia de tipos:** la columna de fecha (A) ahora siempre se almacena y lee como texto formateado (`setNumberFormat("@")`), eliminando comportamientos inconsistentes cuando Sheets convertía el texto a objeto Date.
- **Rango de fechas correcto:** el límite final del año se establece a `new Date(año, 11, 31, 23, 59, 59)` para incluir correctamente los eventos del 31 de diciembre.
- **Refactorización:** la función única se divide en funciones con responsabilidad única (`obtenerOCrearHoja`, `inicializarCabecera`, `leerIdsExistentes`, `filtrarEventosPorMesYColor`, `extraerDatosDelTitulo`, `sincronizarEventos`, `aplicarFormatoCondicional`, `generarTablaResumen`), mejorando la legibilidad y el mantenimiento.

---

## [3.0] - Verificación de pagos y tabla de totales

### Añadido
- **Columna ESTADO:** desplegable con opciones `SIN PAGAR` / `PAGADO` en cada fila de evento.
- **Formato condicional:** las filas con estado `SIN PAGAR` se marcan en rojo claro y las de `PAGADO` en verde claro, de forma automática.
- **Tabla resumen ampliada (columnas G-I):**
  - Columna G: fechas únicas de pago del mes.
  - Columna H: total a pagar ese día (suma de todos los eventos de esa fecha).
  - Columna I: total pagado acumulado del mes + total previsto del mes.

---

## [2.0] - Extracción del número de vencimiento

### Añadido
- **Columna VENCIMIENTO (C):** el script ahora extrae el número de vencimiento bancario del título del evento mediante expresión regular. Se detecta cualquier patrón `DD/MM` presente en el título.
- El valor se almacena como texto plano (precedido de apóstrofo) para evitar que Google Sheets lo interprete como una fecha.

---

## [1.0] - Versión inicial

### Añadido
- Lectura de todos los eventos del año en curso desde un calendario de Google.
- Filtrado por color de evento (rojo tomate y flamingo) para identificar facturas de proveedores.
- Extracción de **fecha** e **importe** del título del evento mediante expresión regular.
- Creación automática de una hoja por cada mes del año.
- Inserción de los eventos como filas con columnas: `FECHA`, `DESCRIPCIÓN`, `IMPORTE (€)`.
- Tabla de **totales diarios** en columnas auxiliares, mostrando cuánto se debe pagar cada día del mes.
