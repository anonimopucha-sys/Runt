# 📋 Guía: Captura de HTML y Corrección de Selectores SIMIT

## ¿Qué es esto?

Cuando el portal SIMIT cambió su estructura en abril 2026, los selectores CSS quedaron obsoletos. Este proceso te ayudará a **identificar los selectores exactos** en el portal real y actualizar el código.

## Pasos

### 1. Ejecutar el script de captura

En tu máquina Windows, en la carpeta del proyecto:

```bash
cd "CE0864_Consulta_Runt_Simit\Aut. Simit x Placa y Cedula"
python capturar_html_simit.py
```

### 2. Usar el navegador

- Se abrirá **Edge** en el portal SIMIT
- Ingresa una **placa o cédula** que sepas que existe (ej: una placa de un vehículo que tengas)
- Haz clic en **"Consultar"**
- Espera a que carguen los resultados
- **Presiona ENTER en la terminal**

### 3. Script generará 3 archivos

Se guardarán en: `Aut. Simit x Placa y Cedula\Data\Calibracion\`

```
page_2026-06-16_14-30-45.html          ← HTML REAL del portal
screenshot_2026-06-16_14-30-45.png     ← Captura visual
selectores_2026-06-16_14-30-45.json    ← Reporte de qué selectores existen
```

### 4. Verificar en el HTML

**Opción A: En VS Code**
- Abre el archivo `.html` generado
- Presiona `Ctrl+F` 
- Busca estos IDs:
  - `txtBusqueda` ← campo de búsqueda
  - `consultar` ← botón buscar
  - `modal-multiples-personas` ← modal si hay conflictos
  - `modalInformation` ← modal de info
  - `multaTable` ← tabla de resultados
  - `resumenEstadoCuenta` ← resumen

**Opción B: En el navegador (más visual)**
- Abre el archivo `.html` en tu navegador (drag & drop)
- Presiona `F12` para abrir DevTools
- Presiona `Ctrl+F` para buscar
- Busca los IDs arriba

### 5. Reportar resultados

Envía el contenido del JSON generado (selectores_*.json) o simplemente confirma:

**Para cada sección, cuál selector encontraste:**

| Elemento | Selector Actual | ¿Encontrado? | Selector Real |
|----------|-----------------|-------------|---------------|
| Campo búsqueda | `#txtBusqueda` | ? | ← completa aquí |
| Botón buscar | `#consultar` | ? | ← completa aquí |
| Modal múltiples | `#modal-multiples-personas` | ? | ← completa aquí |
| Modal información | `#modalInformation` | ? | ← completa aquí |
| Tabla | `#multaTable` | ? | ← completa aquí |
| Resumen | `#resumenEstadoCuenta` | ? | ← completa aquí |

## Selectores actuales en el código

Estos están en: `Aut. Simit x Placa y Cedula\Funciones\f_navegar_simit.py`

```python
# --- Selectores (con alternativas; se prueban en orden) ---
SEL_CAMPO_BUSQUEDA = [
    "#txtBusqueda",                    # ID principal
    "input[placeholder*='laca']",      # Fallback si tiene placeholder
    "input[placeholder*='dula']",      # Fallback para cédula
    "input[type='search']",            # Fallback genérico
]

SEL_BOTON_BUSCAR = [
    "#consultar",                      # ID principal
    "#btnNumDocPlaca",                 # Alternativa
    "button[type='submit']",           # Fallback genérico
]

SEL_RESULTADOS = [
    "#resumenEstadoCuenta",            # Señal de que cargó algo
    "#multaTable"
]

SEL_TABLA_COMP = ["#multaTable"]       # Tabla de comparendos

ID_MODAL_MULTIPLES = "modal-multiples-personas"
ID_MODAL_INFO = "modalInformation"
```

## Si algunos selectores NO existen

Si encuentras que los IDs no existen, me lo confirmas y actualizo el código automáticamente. Por ejemplo:

- Si `#txtBusqueda` no existe pero existe `#inputBusqueda` → lo cambio
- Si `#consultar` no existe pero existe `#btnBuscar` → lo cambio
- Si el modal tiene otro nombre → lo actualizo

## Ejemplo de resultado esperado

Cuando ejecutes el script y presiones ENTER después de buscar, deberías ver algo como:

```
================================================================================
VERIFICACIÓN DE SELECTORES
================================================================================

Campo de búsqueda:
  ✓ Encontrados (1):
    → #txtBusqueda
  ✗ No encontrados (3):
    → input[placeholder*='laca']
    → input[placeholder*='dula']
    → input[type='search']

Botón buscar:
  ✓ Encontrados (2):
    → #consultar
    → button[type='submit']
  ✗ No encontrados (1):
    → #btnNumDocPlaca
```

Esto significa:
- ✓ El selector `#txtBusqueda` **SÍ** existe (está bien)
- ✓ El botón con `#consultar` **SÍ** existe (está bien)
- Los fallbacks no se necesitan pero no causan problema

## ¿Qué pasa si los selectores cambiaron?

Si algunos **no** aparecen en "Encontrados", me envías:

1. La línea del JSON que dice `"#selector": false`
2. O el ID/nombre correcto que viste en el HTML

Entonces actualizo `f_navegar_simit.py` automáticamente.

---

## Resumen del proceso

```
1. Ejecuta capturar_html_simit.py
       ↓
2. Ingresa placa/cédula en el portal
       ↓
3. Presiona ENTER en la terminal
       ↓
4. Se generan 3 archivos en Data/Calibracion/
       ↓
5. Verifica los selectores en el HTML
       ↓
6. Confirma qué selectores encontraste
       ↓
7. Se actualiza f_navegar_simit.py automáticamente
```

---

**¿Listo? Ejecuta el script y cuéntame qué selectores encontraste! 🚀**
