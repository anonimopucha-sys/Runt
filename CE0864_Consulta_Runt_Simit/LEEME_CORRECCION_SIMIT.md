# CE0864 Consulta RUNT / SIMIT — Corrección del módulo SIMIT

Fecha: junio 2026

## Qué se corrigió

El subproceso **`Aut. Simit x Placa y Cedula`** dejó de generar reportes
(filas en blanco en `Consulta_SIMIT.xlsx`) a partir de abril/2026. Causa raíz,
confirmada con los logs de ejecución:

1. **Selectores rotos.** El portal del SIMIT (fcm.org.co/simit) cambió de
   maquetación. La extracción del resumen dependía de rutas CSS frágiles
   (`div.col-lg-3.col-md-3.col-6 > span`...) que dejaron de coincidir, por lo
   que el bot terminaba "bien" pero sin escribir datos.
   - Corridas limpias de 196–204 registros hasta feb/2026.
   - 17-abr: 140/204 excepciones. 22-abr: 202/204. Luego corridas que
     "terminan" con decenas de "Error al extraer/escribir el resumen".
2. **Excel corrupto / bloqueado.** Varias corridas del 30-abr murieron con
   `zipfile.BadZipFile: File is not a zip file` al abrir `Consulta_SIMIT.xlsx`.
   El bot leía y escribía el mismo archivo y lo guardaba **en cada fila**
   (~204 guardados/corrida); cualquier interrupción o sincronización (OneDrive)
   lo dejaba a medio escribir.

## Qué cambió en el código

Solo se modificó: `Aut. Simit x Placa y Cedula/Funciones/f_navegar_simit.py`
(el original quedó como `f_navegar_simit_ORIGINAL.py.bak`).
Los demás subprocesos (RUNT cédula/placa, licencia, VIM) NO se tocaron.

- **Extracción por ETIQUETA visible** ("Comparendos", "Multas", "Acuerdos de
  pago", "Total") en vez de clases Bootstrap. Sobrevive a los rediseños.
  El detalle sigue por `data-label` (estable). Se conservan los selectores
  viejos como *fallback*.
- **Guardado atómico por lotes** (cada 5 filas y al final; archivo `.tmp` +
  `os.replace`; `.bak` al inicio). Elimina el `BadZipFile`.
- **Fallas visibles**: filas marcadas `SIN RESPUESTA` / `ERROR` en vez de
  quedar en blanco; errores con traceback real en el log.
- **Detección de captcha y portal caído** con reintento; en modo visible deja
  25 s para resolver el captcha a mano.
- **Reintentos por registro** con verificación de que cargó un resultado.

## Cómo ponerlo en producción

1. Copiar la carpeta tal cual a `C:\RPA\CE0864_Consulta_Runt_Simit\`
   (mantener esa ruta; los `.bat` y rutas internas la usan).
2. Asegurar que `C:\CE0864_Consulta_RUNT_SIMIT\Consulta_SIMIT.xlsx`
   esté **cerrado** antes de correr y, de ser posible, que esa carpeta esté
   **fuera de OneDrive** (o con la sincronización en pausa durante la corrida).
3. Ejecutar como siempre: `Aut. Simit x Placa y Cedula\Data\ConsultaSimit.bat`.

## Calibración (solo si hiciera falta ajustar el campo de búsqueda o el modal)

En `f_navegar_simit.py`, arriba, poner `CALIBRAR = True`, correr una vez, y
revisar el HTML/captura que quedan en
`Aut. Simit x Placa y Cedula\Data\Calibracion\`. Con eso se confirma el `id`
exacto del campo de búsqueda y del modal de "múltiples personas" y, si fuera
necesario, se ajustan en el bloque de configuración (todo centralizado al
inicio del archivo). Luego volver a dejar `CALIBRAR = False`.

## Requisitos

Sin cambios respecto al original: Python 3.13, selenium 4.35, openpyxl 3.1.5,
pandas, Microsoft Edge + msedgedriver compatibles. Ver `Data/requirements.txt`.
