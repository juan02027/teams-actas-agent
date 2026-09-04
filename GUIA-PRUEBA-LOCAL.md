# Guía de prueba local — Teams Actas Agent

## Requisitos y arranque

Instala Node.js 22 LTS, un navegador moderno y pnpm (`npm install -g pnpm`). Descomprime el ZIP, abre CMD en la carpeta que contiene `package.json` y ejecuta:

```bat
pnpm install
pnpm dev
```

El script activa automáticamente `LOCAL_OPERATOR_MODE=true` durante desarrollo y usa `cross-env`, por lo que funciona en CMD y PowerShell. Abre `http://localhost:3000/`. El lanzador `INICIAR-AGENTE-WINDOWS.cmd` hace lo mismo con doble clic.

## Configurar Groq

Crea `.env` en la raíz:

```dotenv
GROQ_API_KEY=PEGA_AQUI_TU_CLAVE
GROQ_TRANSCRIPTION_MODEL=whisper-large-v3-turbo
GROQ_CHAT_MODEL=openai/gpt-oss-120b
LOCAL_OPERATOR_MODE=true
```

Reinicia `pnpm dev` después de guardar el archivo. La clave nunca debe ir en el frontend ni compartirse.

## Prueba de texto sin grabar

Entra en **Actas → Nueva prueba**, pega una transcripción de al menos 20 caracteres y pulsa **Generar acta**. El agente crea un acta ejecutiva, un documento de compromisos, un PDF y registros en el índice local. En **Compromisos**, solo deben aparecer compromisos expresos que incluyan responsable y fecha de entrega. El estado se puede cambiar entre pendiente, en progreso, completado y bloqueado, o eliminarse.

## Prueba de grabación local

En **Reuniones**, activa **Grabar** en una reunión o pulsa **Grabar localmente**. El navegador solicitará permiso para compartir pantalla y audio. Selecciona la ventana de Teams y el audio del sistema; informa a los participantes y obtén el consentimiento requerido. El agente no graba ocultamente ni modifica el indicador de Teams.

- **Pantalla + audio:** produce `.webm`, con límite operativo de aproximadamente una hora o 180 MB.
- **Modo largo · solo audio:** usa bitrate reducido para reuniones de más de tres horas; la identificación de participantes depende de la claridad del audio y de que se mencionen los nombres.
- Al detener, la grabación se guarda en disco antes del procesamiento. Si Groq responde con error, aparece en **Grabaciones** y el botón **Procesar nuevamente** permite reintentar sin volver a grabar.
- **Descargar grabación** entrega el archivo `.webm` real con tipo MIME de video/audio; no entrega la ruta del panel ni un archivo HTML.

Ubicaciones locales: `data\recordings`, `data\documents` y `data\recordings-index.json`. **Eliminar referencia** quita la entrada del historial sin borrar el archivo físico.

## Microsoft 365 y SharePoint

Para mostrar calendario y escribir actas, configura las variables siguientes:

```dotenv
VITE_MICROSOFT_CLIENT_ID=Application-ID-de-Entra
VITE_MICROSOFT_TENANT_ID=Directory-ID-o-common
VITE_SHAREPOINT_SITE_URL=https://abcstorage.sharepoint.com/transformaciondigital
VITE_SHAREPOINT_LIST_NAME=Reuniones Efectivas
VITE_SHAREPOINT_COLUMN_NAME=Nota
```

En Entra ID registra la plataforma **Aplicación de página única (SPA)** con `http://localhost:3000/` y permisos delegados `Calendars.Read` y `Sites.ReadWrite.All`. El segundo permiso permite crear elementos en listas de SharePoint y normalmente requiere consentimiento administrativo.

Después de **Vincular Microsoft 365**, las reuniones del calendario aparecen en el panel. Tras una transcripción o reprocesamiento, el navegador crea un elemento en `Reuniones Efectivas` y envía únicamente el campo `Nota`, con resumen ejecutivo, decisiones y compromisos con fecha. No se envían otras columnas. La lista no debe tener otros campos obligatorios; si `Title` es obligatorio, configúralo como opcional o con valor predeterminado desde SharePoint.

## Verificación técnica

```bat
pnpm check
pnpm test
pnpm build
```

La suite final contiene 4 archivos de prueba y 7 casos exitosos.

## Problemas frecuentes

- `pnpm` no se reconoce: ejecuta `npm install -g pnpm` y abre una nueva terminal.
- `GROQ_API_KEY no está configurada`: revisa `.env`, reinicia el servidor y confirma que el archivo está en la misma carpeta que `package.json`.
- `interaction_in_progress` o `block_nested_popups`: cierra pestañas de autenticación antiguas y usa **Vincular Microsoft 365**; la versión actual utiliza redirección en la misma pestaña, no popup.
- `AADSTS900971`: revisa que la URI SPA sea exactamente `http://localhost:3000/`.
- SharePoint devuelve 403: falta consentimiento de `Sites.ReadWrite.All` o la cuenta no tiene permiso de edición en la lista.
- Groq falla al generar: la grabación permanece guardada; abre **Grabaciones** y pulsa **Procesar nuevamente** después de corregir la clave o el servicio.
