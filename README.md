<b>Título: Asistente de Generación de Presentaciones en PPT</b>

Este programa busca acelerar el proceso de generación de presentaciones completas en PowerPoint, apoyando a los expositores con temas como:

<ul>
  <li>Optimización de recursos</li>
  <li>Tiempo destinado a investigación</li>
  <li>Bloqueos creativos</li>
</ul>

El proceso es el siguiente: el usuario es recibido por el asistente IA, el cual le pide los siguientes datos: nombre, tema sobre el cual se procederá a realizar la presentación en PPT y la cantidad de diapositivas que necesita; luego, tras proponer y validar una estructura definida el asistente procede a entregar un código VBA para su rápida implementación en PowerPoint a través de la pestaña "Programador" que generará inmediatamente la diagramación básica de la presentación, sin considerar formato. Finalmente, el asistente brindará prompts para la generación de imágenes en el Generador de Imágenes predilecto del usuario (Bing, Midjourney, Stable Diffusion, Leonardo, etc). No se incluye la opción de generar las imágenes directamente a través de Dall-E 3 por una cuestión de ahorro de costos ya que la llamada a la API con Python es de hasta $0.080 por imagen.

Como un plus personal, desarrollé en los últimos días una interfaz con HTML, Python y Flask para deployment en vivo, trabajando con render.com y MongoDB como gestor de BD.

En un entorno local con Visual Studio Code, utilizando un `.venv` y la función `flask run` se obtuvieron los resultados mostrador en /coderhouseAppFlaskLive/demoImg/ demostrándose que efectivamente el código fue eficiente y, además de generar un histórico de las últimas generaciones, permite también generar una vista independiente por cada resultado llevando al usuario a una URL del tipo XXX.0.0.1:XXXX/entry/<entry_id>, desde donde podría verificar y copiar toda la información referente a su presentación; los datos se guardan en una BD en MongoDB ya que las generaciones guardadas en la memoria de Python se eliminaban cada vez que se reiniciaba la Flask App.

Pese a que lo que funcionó en el entorno local no terminó de funcionar en vivo, comparto el link donde se aprecia la interfaz: <a href="https://python-web-coderhouse-flask-live.onrender.com">https://python-web-coderhouse-flask-live.onrender.com</a> (este link será retirado automáticamente en 6 días, es decir, el 14/02/2024); fue un gran reto programar la app y aunque todavía no logro entender por qué al momento de enviar la consulta se genera un error interno del tipo 502, seguiré practicando hasta conseguir el objetivo.

Todo lo demás se encuentra explicado dentro del mismo archivo "preEntrega1-2v3.0.ipynb", incluyendo:
- Objetivos
- Metodología
- Implementación

Adicionalmente incluyo una muestra del tipo de resultado (en formato PPT) que se puede obtener con este código (apariencia/formato fueron implementados manualmente)

Gracias!
Un abrazo

Jorge Orihuela Castillo (Perú)
