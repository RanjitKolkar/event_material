# Event Material Generator v14

NFSU Goa institutional event planning, event history, checklist, budget and document generation application.

## Access
- Users register with an `@nfsu.ac.in` email address.
- Passwords may be any non-empty value and are stored as salted PBKDF2 hashes.
- Users can view and manage their own saved event history.
- Administrators can manage users, institutional master data and all event records.

## Event workflow
1. Sign in.
2. View saved events, or create a new event / load the sample event when no event exists.
3. Enter and save the event master record.
4. Maintain schedule, experts/guests, budget and checklist.
5. Generate the complete Word document package.

## Persistence
Application state, users and event history are stored in the configured persistent data directory. For Streamlit Cloud production use, provide persistent storage or a managed database rather than relying on ephemeral local storage.

## Secrets
Configure `ADMIN_USERNAME`, `ADMIN_EMAIL`, `ADMIN_PASSWORD` and `APP_ENCRYPTION_KEY` through the deployment secret manager. Never commit secrets to GitHub.

## Footer
Developed by NFSU Goa Coding Club  
Coder: Dr. Ranjit Kolkar
