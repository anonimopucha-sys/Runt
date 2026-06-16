# Manual de instalación y ejecución por CMD

Este documento explica paso a paso cómo instalar y correr la automatización de `CE0864_Consulta_Runt_Simit` usando **Windows CMD**.

## 1. Requisitos previos

- Windows 10 o 11.
- Microsoft Edge instalado.
- Python 3.11 o superior.
- El archivo de Excel `Consulta_SIMIT.xlsx` debe estar cerrado mientras se ejecuta el bot.
- Preferible: la carpeta de datos fuera de OneDrive o con sincronización en pausa.

## 2. Copiar la carpeta al equipo

Copia la carpeta completa `CE0864_Consulta_Runt_Simit` a una ruta fija de tu equipo. El ejemplo recomendado es:

```cmd
C:\RPA\CE0864_Consulta_Runt_Simit
```

Ese es el mismo ejemplo que usa el código en `Aut. Simit x Placa y Cedula/Funciones/f_navegar_simit.py`.

## 3. Abrir CMD y moverte al proyecto

Abre el símbolo del sistema de Windows y ejecuta:

```cmd
cd C:\RPA\CE0864_Consulta_Runt_Simit\Aut. Simit x Placa y Cedula
```

Si copiaste el proyecto a otra ruta, ajusta `cd` según corresponda.

## 4. Crear y activar un entorno virtual

Es recomendable usar un entorno virtual para no mezclar dependencias del sistema.

```cmd
python -m venv venv
venv\Scripts\activate
```

Después de activar el entorno, el prompt debe mostrar algo como `(venv)`.

## 5. Instalar dependencias

Usa `pip` para instalar los paquetes necesarios.

```cmd
python -m pip install --upgrade pip
python -m pip install selenium openpyxl pandas pyperclip pillow
```

Si quieres, también puedes instalar desde `requirements.txt`:

```cmd
python -m pip install -r Data\requirements.txt
```

> Nota: si `requirements.txt` incluye paquetes específicos de Windows, el primer comando suele ser más seguro.

## 6. Verificar que Edge funcione con Selenium

El código usa Edge a través de Selenium. La forma más simple de comprobarlo es correr:

```cmd
python -c "from selenium import webdriver; from selenium.webdriver.edge.options import Options; Options(); print('Selenium disponible')"
```

Si ves un error relacionado con EdgeDriver, asegura que Microsoft Edge está instalado y que la versión de Selenium es compatible.

## 7. Ejecutar la automatización

Aun con el entorno virtual activado, ejecuta el script principal:

```cmd
python main.py
```

El script se ejecutará desde la carpeta `Aut. Simit x Placa y Cedula` y registrará la salida en `Data\Log`.

## 8. Configuración importante en el código

El archivo `Funciones\f_navegar_simit.py` contiene estas rutas:

- `ARCHIVO_ENTRADA` y `ARCHIVO_SALIDA` apuntan a:
  - `C:\CE0864_Consulta_RUNT_SIMIT\Consulta_SIMIT.xlsx`
- `REG_EJEC` apunta a:
  - `C:\RPA\CE0864_Consulta_Runt_Simit\Aut. Simit x Placa y Cedula\Data\Reg. Ejecuciones\CE0864_RegEjec.xlsx`

Si tus rutas son distintas, ajusta estos valores.

## 9. Modo calibración

Si necesitas verificar selectores exactos, pon `CALIBRAR = True` en:

```python
Aut. Simit x Placa y Cedula\Funciones\f_navegar_simit.py
```

Luego ejecuta de nuevo. Se generará una carpeta:

```cmd
Aut. Simit x Placa y Cedula\Data\Calibracion\
```

Allí encontrarás archivos `.html` y `.png` con el DOM real y capturas del portal.

## 10. Ejecución recomendada

- Cierra el archivo Excel antes de correr el bot.
- De ser posible, coloca la carpeta de datos fuera de OneDrive.
- Si el portal muestra un modal de varias personas, el bot intentará elegir la primera opción automáticamente.

## 11. Archivos útiles en el repositorio

- `capturar_html_simit.py`: script independiente para capturar HTML y validar selectores.
- `Funciones\f_navegar_simit.py`: código principal del flujo de la automatización.
- `Data\Calibracion\`: carpeta donde se vuelcan los archivos de calibración.

## 12. Solución rápida si da error

1. Revisa que el Excel esté cerrado.
2. Revisa que Python esté en el PATH.
3. Vuelve a activar el entorno virtual.
4. Asegúrate de ejecutar `cd` en la carpeta correcta.
5. Si falla Selenium, instala o actualiza Edge y Edge WebDriver.

---

Con esto tendrás una instalación fácil y explicada por CMD, lista para usarse en Windows.