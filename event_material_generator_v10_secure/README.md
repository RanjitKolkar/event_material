# Event Material Generator v10

Streamlit application for institutional event planning, persistent master data, authentication and document generation.

## Secure authentication
- Default administrator username is `admin` unless `ADMIN_USERNAME` is configured.
- **No production password is hard-coded in the repository.** Set `ADMIN_PASSWORD` in Streamlit Secrets/environment variables.
- If no password is configured on a fresh local installation, a random bootstrap password is generated in the ignored `.private_data/bootstrap_admin_password.txt`.
- Passwords are stored as salted PBKDF2-SHA256 hashes.
- Repeated failed logins are throttled.

## Persistent and encrypted data
Runtime application state is stored outside the source tree in `EVENT_DATA_DIR` (default `.private_data/`). Application state is encrypted using a Fernet key before being written to SQLite. For production, set `APP_ENCRYPTION_KEY` in the platform secret manager.

The repository `.gitignore` excludes databases, keys, secrets, environment files and generated/private data. See `SECURITY.md` for deployment requirements.

## Preloaded leadership
- Chief Patron: Padmashri Dr. J. M. Vyas
- Chair: Dr. Naveen Kumar Choudhary
- Co-Chair: Dr. Lokesh Chouhan

## Features
- Authentication with admin/user roles
- Persistent SQLite storage
- Preloaded NFSU Goa faculty master with designations
- Event heads and leadership master data
- Download/upload event JSON
- Spreadsheet-style schedule with custom columns and rows
- Budget and automatic approval amount
- AI Note Studio / deterministic institutional Note
- Official workshop/e-workshop proposal
- Chief Guest invitation and inauguration programme
- Pre-Event, In-Event and Post-Event A4 checklists
- Complete event material ZIP
- NFSU Goa letterhead bundled
- 3-day Gait Pattern Analysis demo data as the default event when no event has yet been saved

## Run
```bash
pip install -r requirements.txt
streamlit run app.py
```
