# Puesta en marcha local en Windows

Esta versión trabaja para un solo operador. Las grabaciones, el índice y los documentos se guardan en la carpeta `data` del proyecto; no requiere MySQL ni almacenamiento S3 para probar el flujo local.

## 1. Preparar variables

Copia `.env.example` a `.env` si tu paquete lo incluye. Si no existe, crea un archivo `.env` en la raíz con este contenido y reemplaza únicamente los valores entre comillas:

```dotenv
LOCAL_OPERATOR_MODE=true
LOCAL_OPERATOR_NAME="Operador"
GROQ_API_KEY="pega_aqui_tu_clave_de_groq"
GROQ_TRANSCRIPTION_MODEL="whisper-large-v3-turbo"
GROQ_CHAT_MODEL="openai/gpt-oss-120b"
VITE_MICROSOFT_CLIENT_ID="Application-ID-de-Entra"
VITE_MICROSOFT_TENANT_ID="Directory-ID-o-common"
VITE_SHAREPOINT_SITE_URL="https://abcstorage.sharepoint.com/transformaciondigital"
VITE_SHAREPOINT_LIST_NAME="Reuniones Efectivas"
VITE_SHAREPOINT_COLUMN_NAME="Nota"
```

No pongas la clave de Groq en `Home.tsx`, en la URL ni en un repositorio público. `GROQ_API_KEY` solo se lee en el servidor. Si todavía no quieres usar Microsoft 365, deja vacíos los valores `VITE_MICROSOFT_*`; puedes probar grabación local y generación desde texto.

## 2. Instalar y abrir

Desde CMD, dentro de la carpeta del proyecto:

```bat
pnpm install
pnpm dev
```

El comando ya usa `cross-env`, así que no debe volver a aparecer el error `NODE_ENV no se reconoce`. Abre `http://localhost:3000/`.

## 3. Probar una grabación

En **Reuniones**, pulsa **Grabar localmente** o activa el interruptor **Grabar** de una reunión concreta. El navegador pedirá escoger la ventana/pantalla y el audio; selecciona la ventana de Teams y activa compartir audio del sistema. Informa a los participantes y obtén el consentimiento requerido antes de grabar.

- **Pantalla + audio:** guarda un archivo `.webm` de video y audio, hasta aproximadamente una hora o 180 MB.
- **Modo largo · solo audio:** reduce el tamaño para reuniones de más de tres horas; la atribución de personas depende de que el audio sea claro y de que la transcripción conserve nombres.
- Al detener, el archivo se guarda antes de enviarse a Groq. Si falla Groq, la grabación permanece en **Grabaciones** con el botón **Procesar nuevamente**.
- El archivo descargado conserva su extensión `.webm`; el servidor envía `Content-Type` y `Content-Disposition` de archivo, no HTML.

Los archivos físicos se encuentran en `data\recordings`, los documentos en `data\documents` y el índice en `data\recordings-index.json`. El botón **Eliminar referencia** quita la entrada del historial pero no borra físicamente el archivo, por lo que una grabación no reaparece por una recreación automática.

## 4. Microsoft 365 y SharePoint

En **Vincular Microsoft 365**, el inicio de sesión usa una redirección en la misma pestaña. La aplicación registrada como SPA debe tener el redirect URI `http://localhost:3000/`. Agrega permisos delegados `Calendars.Read` y `Sites.ReadWrite.All`; este último normalmente requiere consentimiento administrativo. Después de autorizar, las reuniones del calendario se muestran en el panel.

Al terminar una transcripción local o un reprocesamiento, el navegador obtiene un token delegado y crea un elemento en la lista **Reuniones Efectivas** del sitio configurado. El cuerpo enviado contiene únicamente:

```json
{ "fields": { "Nota": "resumen ejecutivo, decisiones y compromisos con fecha" } }
```

La aplicación no escribe otras columnas. La lista debe permitir crear elementos sin otros campos obligatorios; si la columna `Title` quedó obligatoria por la plantilla de SharePoint, hazla opcional o configura el valor predeterminado desde SharePoint sin cambiar el agente.

## 5. Verificación rápida

1. Genera un acta desde **Actas → Nueva prueba** pegando una transcripción de al menos 20 caracteres.
2. Confirma que aparecen los documentos en **Actas** y que el enlace descarga `.docx` o `.pdf`.
3. Confirma que **Compromisos** muestra solo acciones expresas con responsable y fecha de entrega.
4. Cambia el estado a pendiente, en progreso, completado o bloqueado; también puedes eliminar un compromiso.
5. Con Microsoft 365 vinculado, revisa el nuevo elemento y el contenido de `Nota` en Power Apps/SharePoint.

## Transcripción alternativa para reuniones largas

Si se configura `DEEPGRAM_API_KEY` en el `.env` local, el agente usa Deepgram automáticamente para transcribir las grabaciones y deja de usar Groq para la transcripción. Se recomienda `DEEPGRAM_MODEL=nova-3`. Deepgram admite archivos pregrabados largos y ofrece créditos iniciales según su plan vigente. La generación del resumen continúa usando el proveedor configurado en `GROQ_API_KEY`; para retirar Groq también de esa etapa se debe configurar otro proveedor de generación compatible.
