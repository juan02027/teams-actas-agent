# Vincular una cuenta de Microsoft 365

El panel incluye **Vincular Microsoft 365** e **Importar .ics**. Para sincronizar el calendario real se registra una aplicación en Microsoft Entra ID; una URL de Power Apps o una contraseña no sustituyen ese registro.

## Registro de la aplicación

1. Entra a `https://entra.microsoft.com` con tu cuenta.
2. Abre **App registrations → New registration**.
3. Usa un nombre como `Teams Actas Agent Personal`.
4. Selecciona cuentas de esta organización solamente si usarás tu tenant `abcstorage`; usa `common` únicamente si realmente necesitas cuentas de varios tenants.
5. En **Authentication → Add a platform**, selecciona **Single-page application** y agrega exactamente:

```text
http://localhost:3000/
```

La versión actual usa MSAL con `loginRedirect` y `handleRedirectPromise` en la misma pestaña. No se debe colocar un `client_secret` en el frontend.

## Permisos delegados

Agrega Microsoft Graph → Delegated permissions:

- `Calendars.Read` para leer el calendario de la persona autorizada.
- `Sites.ReadWrite.All` para crear el elemento de la lista de SharePoint con el acta en la columna `Nota`.

`Sites.ReadWrite.All` suele requerir consentimiento de administrador. El administrador puede otorgarlo desde **API permissions → Grant admin consent** o revisar la solicitud pendiente en **Enterprise applications / Admin consent requests**, según la política del tenant.

## Variables del frontend

En el archivo `.env` local configura:

```text
VITE_MICROSOFT_CLIENT_ID=ID_DE_LA_APLICACION
VITE_MICROSOFT_TENANT_ID=ID_DEL_TENANT
VITE_SHAREPOINT_SITE_URL=https://abcstorage.sharepoint.com/transformaciondigital
VITE_SHAREPOINT_LIST_NAME=Reuniones Efectivas
VITE_SHAREPOINT_COLUMN_NAME=Nota
```

El frontend obtiene un token delegado de MSAL y consulta `https://graph.microsoft.com/v1.0/me/calendarView`. Después de procesar una grabación, usa el mismo token para crear un elemento de lista enviando únicamente `fields.Nota`. Nunca pongas `client_secret` ni la clave de Groq en variables `VITE_`.

## Prueba sin Microsoft 365

En **Reuniones** pulsa **Importar .ics**. En Outlook puedes exportar tu calendario como `.ics`; el agente carga los eventos y permite probar el flujo sin credenciales. También puedes usar **Actas → Nueva prueba** para validar la generación de documentos desde texto.

## Errores frecuentes

- `AADSTS900971`: falta la URI SPA; registra `http://localhost:3000/`.
- `interaction_in_progress` o `block_nested_popups`: cierra la autenticación antigua y vuelve a usar el botón; no abras un popup manualmente.
- `403` al escribir SharePoint: falta consentimiento de `Sites.ReadWrite.All` o la cuenta no puede editar la lista.
