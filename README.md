# Gestión de Equipos en Servicio Técnico

Aplicación **HTML + JavaScript** para registrar, por etapas, el proceso de
gestión de equipos críticos en servicio técnico —desde la solicitud de trabajo
hasta el cierre del ciclo—, con **persistencia local** en el navegador y
**exportación a Excel (.xlsx)**.

No requiere servidor, instalación ni conexión a internet: basta con abrir
`index.html` en cualquier navegador moderno.

## Cómo usar

1. Abra **`index.html`** (doble clic, o arrástrelo a una pestaña del navegador).
2. Use el menú lateral para ir a la etapa que desea registrar.
3. Complete el formulario y presione **Guardar registro**.
4. Para obtener el reporte, presione **⬇️ Exportar a Excel** (arriba a la derecha).

> Los datos se guardan en el `localStorage` de **ese** navegador y equipo. Para
> trasladarlos use **Configuración → Descargar respaldo (JSON)** y luego
> **Restaurar respaldo (JSON)** en el otro equipo.

## Características clave

- **Cada etapa es un registro independiente.** Se crea por sí solo, sin obligar
  a seguir la secuencia: puede registrar una reparación o un cierre aunque no
  exista la solicitud previa.
- **Folio manual.** El folio de la solicitud de trabajo se ingresa a mano. En
  el resto de las etapas puede reutilizarlo desde una lista de folios ya usados;
  **al elegir un equipo, las sugerencias de folio se filtran automáticamente a
  las solicitudes asociadas a ese equipo** (puede escribir cualquier folio igual).
- **Contadores de "trabajo abierto".** Los números (badges) del menú izquierdo
  cuentan solo el trabajo pendiente: **al cerrar un ciclo, los registros de ese
  folio dejan de sumar** en los contadores de cada etapa, pero permanecen en las
  tablas, en «Todos los registros» y en la exportación.
- **Técnicos desde lista desplegable** en todas las etapas (administrables en
  Configuración). Lista inicial tomada del documento de proceso (v2.0).
- **Vínculo con el listado de equipos críticos** (965 equipos cargados como
  referencia): busque por inventario, equipo, serie, marca, servicio o ubicación.
- **Inventario de equipos con estado automático.** Vista del inventario completo
  donde el estado de cada equipo —**Operativo**, **No operativo**, **En servicio
  técnico** o **Desconocido** (sin eventos)— se calcula a partir del **último
  evento** registrado, junto con la **fecha de la última actualización**. Al
  hacer clic en un equipo se abre su ficha con **todos sus registros**. Filtra
  por estado y se puede exportar a `.xlsx`.
- **Exportación a `.xlsx` ordenada y tabulada**, sin librerías externas:
  - Hoja **«Bitácora general»** con todos los registros ordenados por folio,
    fecha y etapa.
  - Una hoja por cada etapa con sus campos, ordenadas por folio y fecha.
  - Encabezados resaltados, fila de título congelada y autofiltro.
- **Persistencia local** (`localStorage`) + respaldo/restauración en JSON.

## Etapas del proceso

| # | Etapa | Vía |
|---|-------|-----|
| 1 | Solicitud de trabajo | Inicio (folio manual) |
| 2 | Envío a servicio técnico | Vía A · 6.1 |
| 3 | Estado en servicio técnico | Vía A · 6.2 |
| 4 | Recepción del equipo | Vía A · 6.3 |
| 5 | Visita técnica / diagnóstico | Vía B |
| 6 | Solicitud de cotización | Subflujo comercial · 8.1 |
| 7 | Gestión de orden de compra | Subflujo comercial · 8.2 |
| 8 | Emisión de orden de compra | Subflujo comercial · 8.3 |
| 9 | Reparación del equipo | Común · 9 |
| 10 | Cierre del ciclo | Cierre · 10 |

Las vías A y B son mutuamente excluyentes y el subflujo comercial es opcional;
la aplicación no fuerza estas reglas, solo organiza el registro.

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `index.html` | Página principal de la aplicación. |
| `styles.css` | Estilos de la interfaz. |
| `app.js` | Lógica: formularios por etapa, persistencia y exportación. |
| `xlsx.js` | Generador de archivos `.xlsx` (ZIP + Open XML) sin dependencias. |
| `equipos.js` | Listado de equipos críticos (datos de referencia). |

## Notas técnicas

- Se generan archivos `.xlsx` reales (OOXML) mediante un escritor propio de ZIP
  (método *store*) y XML; abren en Excel, LibreOffice Calc y Google Sheets.
- El listado de equipos se generó a partir de `Listado_de_equipos_Criticos.xlsx`.
  Para actualizarlo, reemplace el contenido de `window.EQUIPOS` en `equipos.js`.
