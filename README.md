
# Automatización de Envíos de Emails a Clientes

Este proyecto permite enviar correos electrónicos reales a una lista de empresas desde un archivo CSV. Además, genera un reporte en Excel con los resultados de los envíos, incluyendo gráficas de pastel y estadísticas de éxito/error.

## Instalación de dependencias

Asegúrate de tener Python 3.8 o superior instalado. Luego, instala las dependencias con:

```bash
pip install -r requirements.txt
````

---

## Configuración de variables de entorno

Crea un archivo llamado `.env` en el mismo directorio que el script y define las siguientes variables:

```
EMAIL_USER=tu_correo@example.com
EMAIL_PASS=tu_contraseña_o_token
EMAIL_SMTP=smtp.example.com
EMAIL_PORT=587
```

Notas importantes:

* `EMAIL_USER`: Tu dirección de correo electrónico.
* `EMAIL_PASS`: Contraseña o token de aplicación.
* `EMAIL_SMTP`: Servidor SMTP (por ejemplo: smtp.gmail.com).
* `EMAIL_PORT`: Generalmente 587 para TLS.

Para Gmail, asegúrate de habilitar el acceso mediante contraseña de aplicación.

---

## Ejecución del script

Para correr el programa, simplemente ejecuta:

```bash
python main.py
```

Esto realizará lo siguiente:

1. Cargará los clientes desde el archivo `clientes.csv`.
2. Enviará un correo personalizado a cada uno usando la plantilla.
3. Registrará cada envío en `envios_email.log`.
4. Buscará rebotes en tu cuenta Gmail.
5. Generará el reporte `reporte_envios.xlsx`.

---

## Archivos incluidos en el proyecto

* `main.py`: Código fuente principal del proyecto.
* `clientes.csv`: Lista de empresas y sus correos electrónicos.
* `plantilla.txt`: Plantilla del cuerpo del correo con marcador `{empresa}`.
* `envios_email.log`: Archivo generado automáticamente con el historial de envíos.
* `reporte_envios.xlsx`: Reporte generado con resumen y gráfica de resultados.
* `requirements.txt`: Lista de dependencias de Python.
* `.env`: (No incluido) archivo local con tus credenciales.

---

## Estructura del reporte Excel

El archivo `reporte_envios.xlsx` contiene:

1. **Listado**: Información de cada destinatario y el estado final (SENT o ERROR).
2. **Graficas**: Gráfica de pastel que muestra la distribución de envíos exitosos y con error.
3. **Resumen**: Totales y porcentajes de éxito y error.

-------------

## Detección de Rebotes (opcional - solo Gmail)

Este script incluye una funcionalidad opcional para detectar **rebotes de correos electrónicos** (por ejemplo, direcciones inválidas, buzones llenos, etc.).

Para realizar esta detección de rebotes, el script se conecta a tu cuenta de **Gmail** utilizando la **API de Gmail** (mediante `google-api-python-client`, `google-auth-oauthlib`, y otros componentes).

### ¿Cómo funciona?

* Se utiliza la API de Gmail en modo de solo lectura (`https://www.googleapis.com/auth/gmail.readonly`) para buscar mensajes recientes con asuntos típicos de rebote: `"Mail Delivery Failed"`, `"Returned mail"`, `"Undelivered"`, etc.
* A partir del cuerpo del mensaje, se extrae la dirección de correo que causó el rebote.
* Si se detecta un rebote, se registra una entrada adicional en el log (`envios_email.log`) con estado `ERROR`.
* En el reporte de Excel, ese destinatario aparece marcado como `ERROR`, aunque previamente haya sido registrado como `SENT`.

### Importante

* Esta funcionalidad **solo está disponible si usas una cuenta de Gmail**.
* Debes tener el archivo `credentials.json` y autorizar el acceso una vez (se genera `token.pickle`).
* Si no usas Gmail, el envío de correos funcionará normalmente, pero **no se podrán detectar rebotes automáticos**.


## Notas finales

* El programa funciona con cualquier proveedor de correo electrónico que soporte SMTP.
* Se utiliza la API de Gmail para detectar rebotes, por lo tanto, es necesario autorizar el acceso a tu cuenta la primera vez que se ejecuta el script.


