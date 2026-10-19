# LOGITEx RRHH — versión web (Vercel)

Sistema de gestión de RRHH: nómina, vacaciones LCT, Gantt, actas de EPP y equipos IT,
declaración de domicilio, evaluación de competencias, matriz de sustitución y usuarios.

## Cómo está armado

| Parte | Archivo | Dónde corre |
|---|---|---|
| Interfaz | `index.html` | No contiene datos personales ni contraseñas. |
| API | `app.py` (Flask) | Función de Python en Vercel. Valida sesión y permisos en cada pedido y entrega la interfaz. |
| Datos | PostgreSQL | Base en la nube (Neon), conectada desde el Marketplace de Vercel. |

- Contraseñas guardadas con PBKDF2; sesiones con cookie `HttpOnly` y `SameSite=Strict`.
- Bloqueo de 15 minutos después de 8 intentos fallidos de ingreso por usuario.
- Copia automática diaria de los datos y otra antes de cada restauración (se conservan 30).
- Si dos personas guardan a la vez, la segunda recibe un aviso y no pisa los cambios de la primera.

## Variables de entorno (Vercel → Settings → Environment Variables)

| Variable | Valor |
|---|---|
| `DATABASE_URL` | La crea sola la integración de Neon. |
| `ADMIN_USERNAME` | Usuario administrador inicial, por ejemplo `cchiappero`. |
| `ADMIN_PASSWORD` | Contraseña de ese usuario, mínimo 8 caracteres. |

El administrador inicial se crea solo la primera vez, cuando la base no tiene usuarios.
Después podés borrar `ADMIN_PASSWORD` de Vercel.

## Nunca subir al repositorio

Archivos con datos del personal (`datos_iniciales.json`, backups `LOGITEx_*.json`, bases `.db`).
El `.gitignore` ya los excluye. Los datos se cargan desde la pantalla **⚙ Datos → Restaurar desde un backup**.

## Probar en la PC (opcional)

```
pip install -r requirements.txt
set ADMIN_USERNAME=admin
set ADMIN_PASSWORD=prueba1234
python app.py
```

Abre en http://127.0.0.1:5000 con una base SQLite local (`rrhh_local.db`).
