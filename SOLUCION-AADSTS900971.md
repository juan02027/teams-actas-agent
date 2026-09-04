# Solución AADSTS900971

En la aplicación de Entra ID, abrir **Authentication → Add a platform → Single-page application (SPA)** y registrar exactamente:

```text
http://localhost:3000/
```

La aplicación local usa MSAL Browser con `redirectUri` y `postLogoutRedirectUri` explícitos, ambos apuntando a `http://localhost:3000/`. No debe usarse únicamente una plataforma Web sin URI registrada para este flujo.

Después de guardar la configuración, cerrar cualquier popup de Microsoft, reiniciar `pnpm dev`, abrir `http://localhost:3000/` y pulsar una sola vez **Vincular Microsoft 365**.
