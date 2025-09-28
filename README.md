# Bot de Relleno Automático para Google Forms

## Descripción
Este proyecto es un bot en Python que automatiza el llenado y envío de formularios de **Google Forms** usando Selenium. Permite:

- Ejecutar múltiples hilos para enviar respuestas con distintos correos.
- Generar listas de correos dinámicamente a partir de nombres, apellidos, años o sufijos numéricos.
- Cargar listas (nombres, apellidos, años, dominios) desde archivos `.txt` mediante diálogos.
- Reutilizar o generar un archivo de emails y llevar control de los ya usados (`*_used.txt`).
- Cancelar en cualquier momento con `Ctrl+Z` (EOF) o `Ctrl+C`.

## Requisitos
- Python 3.9+ (recomendado 3.11)
- Google Chrome instalado
- Dependencias Python:
  - selenium
  - webdriver-manager
  - (Opcional) opencv-python y numpy si quieres detección visual de `enviar.png`

## Configuración recomendada del Formulario de Google
Las siguientes pautas mejoran la tasa de envío exitoso. En la carpeta `ejemplo_config_from/` hay capturas de referencia.

1. No habilites la opción de "Limitar a 1 respuesta" si quieres automatizar múltiples envíos.
2. Evita activar la recolección automática de correos de Google (forzaría login). En su lugar crea una pregunta de texto corto titulada por ejemplo: `Correo electrónico`.
3. Usa etiquetas claras para el campo de email: `Correo`, `Correo electrónico`, `Email` o `E-mail`.
4. Minimiza validaciones estrictas (regex personalizados, longitudes exactas) salvo que sean necesarias.
5. Evita preguntas que requieran interacción humana avanzada: captcha, subir archivos, firma dibujada.
6. Para secciones múltiples, mantén los botones estándar: `Siguiente`, `Continuar`, `Enviar`.
7. Si personalizas el mensaje final, incluye palabras como `Gracias` o `Respuesta registrada` (el bot las usa como señal de confirmación).
8. No ocultes el botón Enviar con CSS hasta que se llene algo específico; el bot rellena pero no dispara eventos complejos de scroll dinámico extra.
9. Mantén opciones de selección (radio / checkbox / desplegable) con al menos 1 opción visible.
10. Si el formulario exige login y `ALLOW_MANUAL_LOGIN=true`, podrás iniciar sesión manualmente en la ventana emergente antes de que siga el bot.

| Captura | Descripción |
|---------|-------------|
| `config.jpeg` | Ajustes generales (sin limitar a 1 respuesta / sin recopilar correos forzados). |
| `config1.jpeg` | Ejemplo de pregunta correcta para correo (texto corto). |
| `config2.jpeg` | Diferentes tipos de preguntas soportadas (radio, checkbox, lista). |
| `config3.jpeg` | Flujo de secciones con botón Siguiente visible. |
| `descargar datos de formulario.png` | Vista de gestión de respuestas / exportación. |

Si algún paso falla, revisa estas condiciones primero.

## Instalación de dependencias
```powershell
python -m pip install --upgrade pip
pip install selenium webdriver-manager
# Opcional (solo si usarás detección visual de la imagen enviar.png)
pip install opencv-python numpy
```

### Instalación rápida con requirements.txt
También puedes instalar todo con el archivo `requirements.txt` incluido:
```powershell
pip install -r requirements.txt
```
Si quieres habilitar la detección visual de `enviar.png`, edita `requirements.txt` y descomenta las líneas de `opencv-python` y `numpy` antes de instalar.


## Ejecución básica
```powershell
python bot.py
```
Al iniciar verás un **menú principal**:
```
1) Usar bot con lista de correos existente
2) Crear lista de correos (generador) y luego usar bot
```

## Flujo Opción 2 (Generar correos)
1. Eliges si cargar cada conjunto (Nombres, Apellidos, Años, Dominios) desde archivo `.txt` o escribirlos manualmente separados por comas.
2. Si no proporcionas años, se generan sufijos numéricos aleatorios (2–4 dígitos) por combinación nombre+apellido.
3. Puedes elegir la ruta y nombre del archivo de salida mediante un diálogo.
4. Se crea un archivo (por defecto `emails_generados.txt`). Luego se te pregunta si lo quieres usar para el bot.

Formato generado:
- Con años: `nombreapellido2004@dominio.com`
- Sin años: `nombreapellido123@dominio.com` (número 2–4 dígitos)

## Uso de un archivo de correos propio
Crea un `.txt` con un correo por línea:
```
correo1@dominio.com
correo2@dominio.com
...
```
Pásalo cuando el script lo solicite o ejecútalo con:
```powershell
python bot.py --emails="ruta\a\correos.txt" --url="https://docs.google.com/forms/..."
```

## Parámetros CLI soportados
(Se pueden combinar con interacción manual.)
- `--url`: URL del formulario de Google Forms.
- `--threads`: Número de hilos (1–10).
- `--max-submissions`: Límite máximo de envíos.
- `--emails`: Ruta al archivo de correos.

Variables de entorno equivalentes:
- `FORM_URL`
- `NUM_THREADS`
- `MAX_SUBMISSIONS`
- `EMAILS_FILE`

Otras variables:
- `REUSE_SINGLE_BROWSER` (true/false)
- `ALLOW_MANUAL_LOGIN` (true/false)
- `CHROME_USER_DATA_DIR` / `CHROME_PROFILE_DIR` para usar un perfil ya logueado.

## Archivos generados
- `emails_generados.txt`: Lista de emails creados (cuando usas el generador).
- `<archivo>_used.txt`: Correos ya utilizados en envíos.
- `page_init.png` / `page_load_timeout.png`: Capturas diagnósticas.

## Cancelación segura
- `Ctrl+Z` (Windows) o `Ctrl+D` (Linux/macOS) -> termina limpio (EOF).
- `Ctrl+C` -> Interrumpe y cierra el navegador si está abierto.

## Detección del botón Enviar
El bot intenta:
1. Buscar botón por texto (Enviar / Submit / Send).
2. Buscar imágenes (`enviar.png`).
3. Click forzado por JavaScript.
4. Enter sobre el formulario.

## Control de campos
- Detecta campo de email mediante `aria-label`, `type=email`, placeholders, etc.
- Rellena campos de texto según heurísticas (nombre, apellido, fecha, hora, teléfono, etc.).
- Selecciona opciones aleatorias en radio buttons, checkboxes y dropdowns.

## Multi-hilo
Cada hilo:
- Abre su propio WebDriver.
- Toma un correo de la cola.
- Intenta enviar el formulario.
- Marca el correo como usado al éxito.

## Recomendaciones
- No abuses de formularios que no te pertenecen.
- Respeta términos de servicio de Google.
- Usa pausas/manual login si el formulario requiere autenticación.

## Posibles Mejoras Futuras
- Soporte JSON estructurado para nombres/apellidos.
- Registro en CSV de resultados (éxito, timestamp).
- Limitar generación por número total pedido.
- Integrar proxy / rotación UA.

## Ejemplo rápido (Generar y usar)
```powershell
python bot.py
# Opción 2 -> generar
# Cargar nombres desde diálogo (nombres_apellidos.txt no se usa aquí; es solo ejemplo de base de datos)
# Dominios: gmail.com,outlook.com
# Años: (Enter)
# Guardar -> emails_generados.txt
# Confirmar uso -> Sí
# Ingresar URL formulario
# Ingresar hilos -> 3
# Envíos -> 50
```

## Archivo de nombres y apellidos
Se incluye `nombres_apellidos.txt` con 1000 combinaciones básicas que puedes reutilizar para construir tus propios correos.

---
**Aviso**: Este proyecto es educativo. El autor no se responsabiliza por usos indebidos.

## ejemplo de pagina:

https://docs.google.com/forms/d/e/1FAIpQLSf-pvFeaTQiwwQVctkp24B47sN5oMri3Q9-pDK0P4jFIUie-Q/viewform?usp=header

