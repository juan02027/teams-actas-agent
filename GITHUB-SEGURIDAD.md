# Publicar Teams Actas Agent en GitHub de forma segura

## Antes de publicar

No subas `.env`, grabaciones, documentos generados, carpetas `data`, archivos ZIP ni claves. El repositorio ya excluye estos archivos mediante `.gitignore`.

Copia `ENV-EJEMPLO.txt` como `.env` únicamente en el servidor donde ejecutarás la aplicación. Reemplaza los valores de ejemplo localmente y nunca subas ese archivo.

Si una clave real fue pegada en un chat, commit, captura o repositorio, revócala en el proveedor y genera una nueva antes de continuar.

## Crear el repositorio

En GitHub crea un repositorio nuevo y privado, sin agregar README, `.gitignore` ni licencia si ya existen en este proyecto. Un repositorio privado es obligatorio mientras la aplicación siga usando `LOCAL_OPERATOR_MODE=true`.

## Publicar desde Windows

Desde CMD, dentro de la carpeta que contiene `package.json`:

```cmd
cd /d C:\TeamsActasDocker\teams-actas-agent
git init
git branch -M main
git add .
git status --short
```

Antes de continuar, verifica que la salida de `git status --short` no incluya:

```text
.env
data/
*.webm
*.mp3
*.mp4
*.zip
```

Si todo es correcto:

```cmd
git commit -m "Preparar despliegue Docker Compose"
git remote add origin https://github.com/TU_USUARIO/TU_REPOSITORIO.git
git push -u origin main
```

GitHub abrirá el navegador o solicitará autenticación. Usa un token de acceso personal o GitHub Desktop; no pongas una contraseña de GitHub en la URL del remote.

## Verificación después de publicar

En GitHub revisa manualmente la pestaña de archivos y confirma que no existan `.env`, claves, grabaciones, PDFs privados ni documentos con información de reuniones. Activa también:

- Repositorio privado.
- Autenticación de dos factores en GitHub.
- Secret scanning y push protection si están disponibles.
- Revisión de colaboradores y permisos mínimos.

## Para Render u otro hosting

No subas las claves al repositorio. Configúralas en el panel del proveedor, en Environment Variables. Las variables `VITE_*` se necesitan durante el build del frontend; las demás se necesitan durante la ejecución del contenedor.

El repositorio solo contiene código. Las grabaciones deben almacenarse en un volumen persistente o en almacenamiento externo como Azure Blob, SharePoint/OneDrive o S3.
