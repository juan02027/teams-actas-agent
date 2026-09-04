# Configuración de Teams Actas Agent en Microsoft 365

## Alcance actual: una sola persona

Este agente está diseñado para un único operador. Debe iniciar sesión con su cuenta de Microsoft 365 y utilizar permisos delegados: el agente solo podrá leer el calendario, reuniones, grabaciones y transcripciones a las que esa persona tenga acceso. No se debe solicitar acceso de aplicación para todo el tenant si el objetivo es personal.

## Qué está construido

La aplicación incluye un panel privado para un único operador, controles por reunión para activar o desactivar grabación y procesamiento, una política corporativa, el motor de IA estructurado, generación de Word/PDF y almacenamiento de archivos. En la instalación Windows se usa Groq: `whisper-large-v3-turbo` para transcripción y `openai/gpt-oss-20b` para extracción JSON estricta, con fallback JSON si el workspace no admite strict mode.

El flujo corporativo previsto es: Teams genera la grabación y la transcripción; Microsoft Graph notifica que el artefacto está disponible; Power Automate o un receptor HTTP obtiene el archivo `.vtt`; el endpoint de integración del agente recibe la transcripción; la IA extrae el acta y los compromisos; Word/PDF se guarda en SharePoint/OneDrive; finalmente el operador revisa antes de distribuir. Para una prueba local, el operador también puede capturar pantalla y audio con consentimiento, guardar el WebM en `data/recordings`, procesarlo con Groq y generar documentos en `data/documents`.

## Endpoints preparados

- `GET/POST /api/integrations/teams/notifications`: endpoint de validación y recepción inicial de notificaciones de Microsoft Graph. La validación devuelve `validationToken` y las notificaciones se validan mediante `TEAMS_GRAPH_CLIENT_STATE`.
- `POST /api/integrations/teams/transcript`: endpoint para que Power Automate entregue una transcripción lista para procesar. Requiere el encabezado `x-teams-agent-token` con el valor de `TEAMS_AGENT_INGEST_TOKEN`.

El endpoint de ingestión acepta este JSON:

```json
{
  "graphMeetingId": "id-de-la-reunion-en-graph",
  "title": "Comité de operaciones",
  "organizerName": "Mariana López",
  "organizerEmail": "mariana@empresa.com",
  "scheduledAt": "2026-08-28T15:00:00Z",
  "attendeesCount": 9,
  "transcript": "Mariana: ..."
}
```

## Permisos de Entra ID

Para procesar reuniones organizadas por toda la empresa, registrar una aplicación en Microsoft Entra ID y solicitar consentimiento administrativo para los permisos mínimos que correspondan al alcance:

- `OnlineMeetingTranscript.Read.All` para transcripciones de reuniones programadas.
- `OnlineMeetingRecording.Read.All` para grabaciones de reuniones programadas.
- `CallTranscripts.Read.All` y `CallRecordings.Read.All` si se incluirán llamadas ad hoc o PSTN.
- `OnlineMeetingAiInsight.Read.All` solo si se usará directamente el resumen/insights generados por Teams y existe la licencia correspondiente.

Para que el panel escriba automáticamente el resultado en la lista `Reuniones Efectivas`, la aplicación SPA debe tener el permiso delegado `Sites.ReadWrite.All` y solicitarlo junto con `Calendars.Read`. Microsoft Graph documenta que crear un elemento de lista requiere `Sites.ReadWrite.All` para cuentas laborales o educativas. El panel usa el sitio `https://abcstorage.sharepoint.com/transformaciondigital` por defecto y envía únicamente `{ fields: { Nota: "..." } }`; no modifica otras columnas.

Microsoft documenta que el acceso a transcripciones puede estar bloqueado por un control administrativo del tenant. También debe habilitarse la atribución de hablantes si el acta necesita asignar compromisos por persona.

## Política de Teams

Revisar y asignar las políticas de reunión, eventos y llamadas para:

1. Permitir grabación y transcripción para los grupos autorizados.
2. Definir si se exige consentimiento explícito de los participantes.
3. Configurar el aviso y la URL de privacidad corporativa.
4. Definir caducidad de grabaciones y transcripciones.
5. Decidir si se permite la descarga desde SharePoint/OneDrive.
6. Configurar el valor de grabación automática. En reuniones, Teams puede dejar el interruptor de grabación automática apagado por defecto; para forzarlo por plantilla o etiqueta de sensibilidad puede requerirse Teams Premium.

La aplicación respeta una política conservadora: la grabación está apagada por defecto en la interfaz inicial, el operador puede cambiarla por reunión y la distribución puede requerir revisión humana.

## Flujo recomendado de Power Automate

1. Recibir el evento de Microsoft Graph en un endpoint HTTP público y responder a la validación.
2. Obtener los metadatos de la reunión desde Teams/Graph.
3. Obtener el contenido de la transcripción mediante la acción de Teams equivalente a **Get meeting transcript content**.
4. Invocar `POST /api/integrations/teams/transcript` con el JSON anterior y el encabezado secreto.
5. Guardar los documentos devueltos en la biblioteca de SharePoint definida por el operador, o conservar los enlaces de almacenamiento de la aplicación.
6. Enviar una aprobación al operador; solo después de aprobar, notificar a las personas responsables.
7. Registrar en SharePoint o Dataverse el estado de distribución y cualquier corrección manual.

## Seguridad y operación

- No colocar secretos en el frontend ni en la URL. Configurar `TEAMS_GRAPH_CLIENT_STATE` y `TEAMS_AGENT_INGEST_TOKEN` como secretos del servidor.
- Restringir el acceso del panel al rol `admin` de la aplicación; los usuarios comunes reciben `403`.
- Mantener idempotencia por `graphMeetingId` para evitar duplicar reuniones. Antes de producción, agregar idempotencia de artefactos si el mismo evento puede reintentarse después de generar documentos.
- Conservar solo las grabaciones y transcripciones necesarias, aplicar retención corporativa y documentar el tratamiento de datos personales.
- Validar legalmente el uso de grabaciones, especialmente para participantes externos y llamadas PSTN.

## Prueba inicial

Antes de activar toda la empresa, probar con un grupo piloto y una reunión controlada. Confirmar: grabación visible para participantes, transcripción con nombres, recepción del evento, acta, compromisos por persona, documentos Word/PDF, aprobación del operador y aviso de distribución.
