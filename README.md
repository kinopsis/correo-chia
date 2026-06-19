# correo-chia

Modulo de correo electronico para el sistema de Atencion Ciudadana - Alcaldia de Chia.

## Arquitectura

```
email_settings.provider --> 'smtp' (default) o 'graph_api'
                              |
             +----------------+----------------+
             |                                 |
      imap-poller                        send-reply
      (provider-aware)                   (provider-aware)
             |                                 |
     +-------+-------+               +--------+--------+
     |               |               |                 |
  Graph API        IMAP          Graph API           SMTP
```

## Edge Functions

| Funcion | Proposito |
|---------|-----------|
| `email-config` | Gestion de configuracion y pruebas de conexion SMTP/Graph API |
| `imap-poller` | Recepcion de correos entrantes via IMAP o Graph API |
| `send-reply` | Envio de respuestas a ciudadanos via SMTP o Graph API |
| `manage-users` | Gestion CRUD de usuarios del sistema |

## Configuracion de Secrets

Los secrets se configuran en Supabase Dashboard > Edge Functions > Secrets:

### SMTP/IMAP (Principal)
- `SMTP_PASSWORD` - Contrasena SMTP
- `IMAP_PASSWORD` - Contrasena IMAP

### Graph API (Alternativa)
- `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`

## Tecnologias

- **Runtime**: Deno (Supabase Edge Functions)
- **Protocolos**: SMTP/STARTTLS, IMAPS, Microsoft Graph API
- **Frontend**: React + TypeScript + Tailwind CSS
- **Base de datos**: PostgreSQL (Supabase)
