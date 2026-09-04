# Security and deployment guide

This application is designed so that the Git repository contains **code and non-sensitive template assets and starter data only**. Runtime event data, database files, encryption keys and credentials must never be committed.

## Secrets

Set these in the hosting provider's secret manager (Streamlit Community Cloud: app Secrets) or as environment variables:

- `ADMIN_USERNAME` — defaults to `admin` if omitted.
- `ADMIN_PASSWORD` — required for predictable first-run administration. Use a unique long password.
- `APP_ENCRYPTION_KEY` — Fernet key used to encrypt all JSON application state stored in SQLite.
- `EVENT_DATA_DIR` — optional persistent directory outside the Git checkout.

Generate a Fernet key with:

```python
from cryptography.fernet import Fernet
print(Fernet.generate_key().decode())
```

If `APP_ENCRYPTION_KEY` is not provided during local development, the app creates an ignored key at `.private_data/app_encryption.key`. For production, always supply the key through the platform's secret manager.

If `ADMIN_PASSWORD` is not provided on a fresh local installation, the app creates a random one-time bootstrap password in `.private_data/bootstrap_admin_password.txt`. Never copy this file into Git.

## What is protected

- Passwords are stored as salted PBKDF2-SHA256 hashes.
- Event/application state in SQLite is encrypted using Fernet before storage.
- Database and key files are written under a private directory with restrictive permissions where the operating system supports them.
- Login failures are throttled after repeated attempts.
- Admin-only functionality is role-gated.
- JSON imports are treated as untrusted input and should be size-limited/validated before use.

## Public repository rules

Do not commit:

- `event_material_generator.db`
- `.private_data/`
- `.env`
- `.streamlit/secrets.toml`
- real passwords
- API keys
- encryption keys
- real participant/personal data
- confidential event documents

The bundled Gait Pattern Analysis starter event contains no confidential credentials and should be replaced with institutional data before operational use if required.

## Production recommendations

1. Use HTTPS only.
2. Use a strong unique administrator password and rotate it after deployment.
3. Store `APP_ENCRYPTION_KEY` separately from the database backup.
4. Use persistent storage for the database. Streamlit Community Cloud's local filesystem should not be treated as durable storage.
5. Back up the encrypted database securely.
6. Restrict admin accounts to trusted personnel.
7. Do not expose the SQLite database or private data directory through a static web server.
8. If the application becomes multi-user at scale, move persistence to a managed database and use an enterprise identity provider/OIDC.
9. Do not place confidential event data into GitHub issues, README files, screenshots or sample JSON files.

## Existing database migration

Older local versions may contain plaintext JSON in `app_state`. On startup, the current version attempts to read legacy plaintext values and re-save them encrypted. After verifying migration, securely remove obsolete database backups containing plaintext data.
