# Gestión Equipos Críticos · SIGEM / HHHA

Aplicación web **autocontenida y offline** para la gestión de equipos biomédicos
críticos de un hospital: inventario, programa de mantención preventiva (MP),
cumplimiento, ciclos correctivos, pendientes, conciliación con el archivo maestro
Excel, contactos por servicio y respaldo/sincronización opcional en Google Sheets.

## Cómo usar

Abre `index.html` en un navegador moderno. Todo corre en el cliente; los datos se
guardan en `localStorage` (comprimidos con LZString) y, si se configura, se
sincronizan con un Google Sheet vía Google Apps Script.

## Arquitectura

`index.html` es un *build* de archivo único que ensambla, sin reescribir la lógica:

| Módulo | Rol |
| --- | --- |
| `ui/styles.css` | Estilos (capa visual) |
| `src/seed-data.js` | Datos semilla del inventario |
| `src/hhha-core.js` | **Núcleo lógico**: modelo de dominio, estado, persistencia, reglas de negocio, motor de estados, conciliación con el maestro y operaciones (MP / eventos / pendientes / baja) |
| `ui/app.js` | Capa de vistas (presentación); toda la lógica vive en el núcleo |
| `ui/vendor/lz-string.min.js`, `ui/vendor/xlsx.full.min.js` | Compresión y lectura/escritura de Excel (offline) |

El núcleo (`HHHA`) expone una API pura e inyectable (UI y entorno desacoplados),
de modo que la lógica puede ejercitarse también fuera del navegador (Node).

## Vistas principales

Equipos · Pendientes · MP del mes · Ciclos correctivos · Bitácora de eventos ·
Cumplimiento por servicio · Tiempos de resolución · Recordatorios · Contactos ·
Conciliación · Configuración.

## Registro masivo de MP

La vista **«MP del mes»** permite seleccionar varios equipos (casillas) y, además
de **asignar responsable** en bloque, **registrar la MP del mes en bloque**
(mismo resultado, fecha y ejecutor). El registro masivo reúsa el motor del núcleo
(`registrarMPMasiva`): aplica los efectos de cada MP (estado del equipo, marca del
mes en la carta gantt, pendientes automáticos por causal C1–C8) y puede omitir los
equipos que ya tienen una MP registrada ese mes.
