# Configurar Groq en Windows

En la carpeta que contiene `package.json`, crea un archivo llamado `.env` y agrega:

```env
GROQ_API_KEY=PEGA_AQUI_TU_CLAVE_DE_GROQ
GROQ_TRANSCRIPTION_MODEL=whisper-large-v3-turbo
GROQ_CHAT_MODEL=openai/gpt-oss-120b
LOCAL_OPERATOR_MODE=true
# Opcional: reduce el texto máximo enviado al modelo de acta.
GROQ_MAX_TRANSCRIPT_CHARS=50000
```

No incluyas la clave en el código ni la subas a Git. Reinicia el servidor después de guardarlo:

```cmd
Ctrl + C
pnpm dev
```

`whisper-large-v3-turbo` transcribe el audio. `openai/gpt-oss-120b` genera el acta y los compromisos con JSON estricto. El agente filtra compromisos: solo conserva una acción expresamente asignada, con responsable identificable, fecha o plazo de entrega y evidencia de la transcripción.

Las grabaciones se guardan en `data\recordings`, los documentos en `data\documents` y el índice en `data\recordings-index.json`. La captura de video está limitada en la interfaz a aproximadamente una hora o 180 MB; el modo largo es audio-only con bitrate reducido y está pensado para reuniones de más de tres horas. Si Groq falla, el archivo queda visible en **Grabaciones** y puedes usar **Procesar nuevamente** sin volver a grabarlo.

## Consumo y privacidad

Whisper se factura por duración del audio, no por tokens. El modelo de actas sí consume tokens: recibe la transcripción y devuelve el resumen estructurado. La aplicación ahora solicita respuesta JSON estricta, limita la salida a 1.500 tokens, limita por defecto la transcripción a 50.000 caracteres, elimina timestamps que no utiliza y, para transcripciones muy largas, conserva el inicio, el final y las líneas relacionadas con decisiones, fechas y compromisos. El servidor registra en la consola el consumo `prompt`, `completion`, `total` y `cached` para poder medirlo.

La aplicación usa los endpoints directos de inferencia de Groq; no usa Batch ni Fine-tuning. Según la documentación de Groq, por defecto las solicitudes de inferencia no conservan inputs ni outputs; se conserva metadata de uso. Groq puede retener temporalmente inputs/outputs hasta 30 días para confiabilidad o investigación de abuso. La plataforma ofrece **Zero Data Retention (ZDR)** en **Data Controls** para evitar esa retención. Groq indica que los datos retenidos se alojan en buckets de Google Cloud Platform en Estados Unidos. En este proyecto, la grabación original permanece local en `data\recordings`, y el resumen/compromisos también quedan localmente y, si se vincula Microsoft 365, el contenido de `Nota` se guarda en la lista SharePoint configurada.
