# 📅 Gestión de Pagos a Proveedores: Google Calendar → Sheets
> Automatización de la comprobación de vencimientos de facturas de compras, desarrollada como solución real para una empresa que gestiona sus pagos a proveedores a través de Google Workspace.

---

## 🧩 El problema

La empresa tenía un flujo de trabajo manual para controlar el vencimiento de sus facturas de compras a proveedores. El proceso consistía en revisar uno a uno cada banco y cada aplicación donde se realizaban los cobros domiciliados, lo cual resultaba:

- **Lento**: cada comprobación requería acceder a múltiples aplicaciones bancarias.
- **Propenso a errores**: la información estaba fragmentada y sin una vista unificada.
- **Difícil de delegar**: sin una fuente de verdad centralizada, no había forma sencilla de que otra persona llevara el control.

La empresa ya usaba **Google Calendar** para registrar sus facturas de compra, anotando en cada evento el importe y la fecha de vencimiento dentro del título. El reto era convertir esa información dispersa en un **dashboard centralizado y siempre actualizado** dentro de Google Sheets.

---

## 💡 La solución

Un script de **Google Apps Script** que sincroniza automáticamente los eventos del calendario con una hoja de cálculo, organizados por mes, con totales diarios, estado de pago y resumen mensual.

```
Google Calendar  ──►  Apps Script  ──►  Google Sheets
(eventos/facturas)    (sincronización)   (dashboard de pagos)
```

### ¿Qué hace el script?

- Lee todos los eventos del año en curso del calendario de la empresa.
- Filtra los eventos marcados con colores específicos que identifican facturas (rojo tomate y flamingo).
- Extrae del título de cada evento: **importe**, **número de vencimiento** y **fecha**.
- Crea automáticamente una hoja por cada mes con los datos estructurados.
- Permite marcar cada factura como **PAGADO / SIN PAGAR** mediante un desplegable.
- Genera una tabla resumen con el **total a pagar por día**, el **total pagado** y el **total previsto del mes**.
- Evita duplicados de forma robusta usando el ID nativo de cada evento de Google Calendar.

---

## 🚀 Evolución del proyecto

El proyecto se desarrolló de forma iterativa, respondiendo a necesidades reales del cliente:

### v1.0 — Sincronización básica
Primera versión funcional. Recoge los eventos del calendario, extrae fecha, nombre e importe del título, y los vuelca en una hoja de cálculo organizada por mes. Añade una columna de totales diarios para tener una visión rápida de cuánto se debe pagar cada día.

### v2.0 — Extracción del número de vencimiento
El cliente necesitaba también ver el número de vencimiento bancario (formato `DD/MM` separado con `/` en el título del evento). Se añadió extracción mediante expresión regular y una columna específica en la hoja.

### v3.0 — Verificación de pagos y tabla de totales
Se incorpora la columna **ESTADO** con desplegable `PAGADO / SIN PAGAR` y formato condicional visual (rojo/verde). Se amplía la tabla resumen para mostrar el **total pagado** y el **total previsto** del mes, permitiendo al cliente saber de un vistazo cuánto queda por pagar.

### v3.1 — Corrección de duplicados con ID nativo
El cliente detectó que al re-ejecutar el script se duplicaban filas. La solución fue usar `ev.getId()`, el identificador único nativo de Google Calendar, como clave de sincronización (almacenado en una columna oculta).

> **Por qué esto es importante:** antes el script dependía de que el título y la fecha del evento no cambiaran. Con el ID nativo, se puede renombrar o editar el evento en Calendar sin que el script cree una fila duplicada. Es la forma correcta de gestionar sincronizaciones entre sistemas.

---

## ⚙️ Instalación y configuración

### Requisitos
- Cuenta de Google con acceso a **Google Sheets** y **Google Calendar**.
- Los eventos de facturas deben estar marcados con color **rojo tomate (11)** o **flamingo (4)** en Google Calendar.
- El título de cada evento debe seguir el formato:  
  `Nombre proveedor 15/03 123,45` *(vencimiento y/o importe en cualquier posición del título)*

### Pasos

1. Abre o crea una **Google Spreadsheet** en blanco.
2. Ve a **Extensiones → Apps Script**.
3. Copia el contenido de `importarEventosConImporte.gs` en el editor.
4. Edita la constante `CALENDAR_ID` con el ID de tu calendario:
   ```javascript
   var CALENDAR_ID = 'tu_calendario@gmail.com';
   ```
   *(Encuéntralo en Google Calendar → ⚙️ Configuración del calendario → ID del calendario)*
5. Guarda el proyecto y ejecuta la función `importarEventosConImporte`.
6. Acepta los permisos de acceso a Calendar y Sheets cuando se soliciten.
7. *(Opcional)* Configura un **trigger por tiempo** para que se ejecute automáticamente cada día o semana.

---

## 📊 Estructura de la hoja resultante

| Columna | Contenido |
|---------|-----------|
| A | Fecha del evento (DD/MM/YYYY) |
| B | Descripción / título del evento |
| C | Número de vencimiento (DD/MM) |
| D | Importe (€) |
| E | Estado: `SIN PAGAR` / `PAGADO` |
| F | ID del evento *(oculta, uso interno)* |
| G | Fechas únicas del mes (tabla resumen) |
| H | Total a pagar ese día |
| I | Total pagado / Total previsto del mes |

---

## 🤖 Sobre el uso de IA en este proyecto

Este proyecto se desarrolló combinando conocimiento propio con el apoyo de herramientas de IA (**Claude** de Anthropic y **Gemini** de Google) en un flujo de trabajo de *vibe coding*.

El proceso fue el siguiente:

- El **análisis de requisitos**, la **comprensión del flujo de trabajo del cliente** y las **decisiones de diseño** (qué extraer, cómo estructurar la hoja, cómo gestionar los colores de calendario como señal) fueron tomadas de forma autónoma.
- La IA se utilizó como **par de programación** para agilizar la escritura de código repetitivo, resolver dudas sobre la API de Google Apps Script y recibir consejo en partes que se atragantaban, como la gestión correcta de zonas horarias o la estrategia de detección de duplicados.
- El **debugging**, la **detección de errores lógicos** (como el problema de duplicados en v3.1 o la variable `totalPrevisto` no declarada) y la **toma de decisiones técnicas** fueron realizados de forma crítica y consciente, usando la IA como herramienta, no como sustituto del criterio propio.

La IA no sabe qué necesita el cliente. Eso lo sabe el desarrollador.

---

## 📁 Estructura del repositorio

```
├── importarEventosConImporte.gs   # Script principal de Google Apps Script
├── CHANGELOG.md                   # Historial detallado de versiones
└── README.md                      # Este archivo
```

---

## 📄 Licencia

Este proyecto se comparte con fines educativos y de demostración. Libre de usar y adaptar con atribución. GNU GPL v3.0
