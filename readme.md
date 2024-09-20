# Extraer Potencia Total Ejecutada de Monitorieo SEIN

Un proyecto de WebScraping

## Instalación

Instalar selenium
```bash
pip install selenium
```

## Uso

Primero, instalar el Navegador Mozilla Firefox

https://www.mozilla.org/en-US/firefox/download/thanks/

Luego, descargar el Gekko Driver, luego extraerlo

https://github.com/mozilla/geckodriver/releases

Si usas Windows, seleccionar el "geckodriver-v0.35.0-win32.zip"

luego extraerlo.

Ahora ejecutas el programa:

![UI](/ui.png)

En el campo  de la ruta del gekko driver, seleccionar la ruta del geckodriver.exe

Después, la ruta del firefox.exe, donde tienes instalado el firefox

Luego el link, ya viene por defecto: https://www.coes.org.pe/Portal/portalinformacion/MonitoreoSEIN
Asi que no hay que modificar eso

Luego el nombre del excel, por defecto "datos_extraidos", no recomiendo modificarlo, si no encunetra un excel con el nombre, el programa crea uno.

Por ultimo, el rango de tiempo que hara las lecturas, esta con un minuto porque estaba verificando, pero lo normal es 30 minutos, como me comunicaron.

