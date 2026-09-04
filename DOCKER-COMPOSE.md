# Teams Actas Agent con Docker Compose

## Requisitos

Se necesita un equipo o servidor con Docker Desktop (Windows) o Docker Engine + Compose (Linux), al menos 4 GB de RAM recomendados, y una carpeta persistente para los archivos de grabación y documentos.

## Instalación

Copie el proyecto al equipo servidor y cree `.env` en la misma carpeta que `docker-compose.yml`:

```env
NODE_ENV=production
LOCAL_OPERATOR_MODE=true
LOCAL_OPERATOR_NAME=Operador

DEEPGRAM_API_KEY=su_clave_deepgram
DEEPGRAM_MODEL=nova-3
GROQ_API_KEY=su_clave_de_groq
GROQ_CHAT_MODEL=openai/gpt-oss-120b

VITE_MICROSOFT_CLIENT_ID=su_client_id
VITE_MICROSOFT_TENANT_ID=su_tenant_id

VITE_SHAREPOINT_SITE_URL=https://abcstorage.sharepoint.com/transformaciondigital
VITE_SHAREPOINT_LIST_NAME=Reuniones Efectivas
VITE_SHAREPOINT_COLUMN_NAME=Nota
```

No incluya `.env` en Git ni lo envíe a otras personas. Las claves se inyectan al contenedor mediante `env_file`.

Inicie el servicio:

```bash
docker compose up -d --build
```

Compruebe el estado y los registros:

```bash
docker compose ps
docker compose logs -f teams-actas-agent
```

Abra localmente:

```text
http://localhost:3000
```

Las grabaciones y documentos quedan en el volumen Docker `teams_actas_data`, por lo que sobreviven a reinicios y actualizaciones del contenedor.

## Acceso desde la red local

Obtenga la IP privada del servidor:

```bash
# Linux
hostname -I

# Windows PowerShell
ipconfig
```

Desde otro equipo de la misma red abra:

```text
http://IP_PRIVADA_DEL_SERVIDOR:3000
```

Si Windows Firewall lo solicita, permita el puerto TCP 3000 en la red privada. En PowerShell como administrador:

```powershell
New-NetFirewallRule -DisplayName "Teams Actas Agent 3000" -Direction Inbound -Protocol TCP -LocalPort 3000 -Action Allow -Profile Private
```

## Publicarlo con IP pública

Una IP pública no es suficiente: el router debe reenviar el puerto y el servidor debe aceptar el tráfico.

1. Reserve una IP privada fija para el equipo servidor, por ejemplo `192.168.1.50`.
2. En el router cree un reenvío TCP desde un puerto externo hacia `192.168.1.50:3000`.
3. Abra el puerto correspondiente en el firewall del servidor.
4. Consulte la IP pública desde el servidor con `curl ifconfig.me`.
5. Pruebe desde una red móvil, no desde la misma Wi-Fi:

```text
http://SU_IP_PUBLICA:3000
```

Algunos proveedores bloquean puertos entrantes o utilizan CGNAT; en ese caso el reenvío no funcionará. Use una VPN o un túnel seguro, como Cloudflare Tunnel, en lugar de abrir el puerto directamente.

## Advertencia de seguridad obligatoria

La configuración `LOCAL_OPERATOR_MODE=true` está diseñada para uso local y permite entrar sin el inicio de sesión del servidor. **No exponga esta configuración directamente a Internet**, porque cualquier persona que conozca la URL podría ver grabaciones, actas y operar la integración de SharePoint.

Para una prueba entre personas conocidas, use una VPN, red corporativa o túnel con autenticación. Para producción se debe agregar autenticación real delante del contenedor (proxy con HTTPS y usuario/contraseña o Microsoft Entra ID), y registrar la URL HTTPS exacta como Redirect URI de la aplicación Microsoft.

Microsoft 365 debe tener registrada la dirección desde la que se abrirá la aplicación, por ejemplo:

```text
https://actas.su-dominio.com/
```

No use una IP pública HTTP para un despliegue corporativo definitivo. Use un dominio y HTTPS; además, actualice la Redirect URI SPA en Entra ID y vuelva a iniciar el contenedor.

## Actualizar

```bash
docker compose down
docker compose up -d --build
```

El volumen `teams_actas_data` no se elimina con esos comandos. Para hacer una copia:

```bash
docker run --rm -v teams_actas_data:/data -v "$PWD":/backup alpine tar czf /backup/teams-actas-data.tgz -C /data .
```

No ejecute `docker compose down -v` salvo que quiera borrar permanentemente los datos del volumen.
