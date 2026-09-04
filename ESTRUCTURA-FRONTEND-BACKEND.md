# Estructura del proyecto

El código quedó organizado en dos carpetas de referencia:

- `frontend/client`: interfaz React, páginas, componentes y estilos.
- `backend/server`: API Express/tRPC, grabaciones, transcripción, Groq, documentos y SharePoint.
- `backend/shared`: tipos compartidos entre ambos lados.
- `backend/drizzle`: esquema y migraciones.

La copia raíz `client/` y `server/` se conserva porque Vite, tRPC y el servidor local necesitan resolver rutas compartidas desde una misma raíz. Esto permite que la aplicación siga funcionando inmediatamente con `pnpm dev` y evita romper imports o el proxy `/api`.

## Ejecución local

```cmd
cd /d C:\TeamsActasAgent\teams-actas-agent
pnpm install
pnpm dev
```

Para compartirla mediante un enlace, debes publicar el proyecto en un hosting con HTTPS. El enlace local `http://localhost:3000` solo funciona en tu propio computador. No compartas una API key dentro del frontend ni en un repositorio público.

## Acceso desde otro computador de la misma red

En el computador donde ejecutas la aplicación, abre CMD y ejecuta:

```cmd
ipconfig
```

Busca la dirección **IPv4**, por ejemplo `192.168.1.25`. Inicia la aplicación con `pnpm dev`; el servidor ahora escucha en `0.0.0.0` y muestra la dirección LAN en la consola. La otra persona debe escribir directamente en la barra de direcciones del navegador:

```text
http://192.168.1.25:3000/
```

No lo escribas en Google ni en el buscador. Si Windows muestra una alerta del Firewall, permite Node.js en **redes privadas**. Ambos computadores deben estar conectados a la misma red Wi-Fi o Ethernet. Si el servidor cambia de puerto porque el 3000 está ocupado, usa el puerto que aparezca en la consola.
