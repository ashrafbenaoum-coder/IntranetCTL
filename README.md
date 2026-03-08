# Intranet Controle Qualite

Application intranet full-stack pour les controleurs qualite:
- login securise par user/password
- saisie de mouvements (N Support, EAN, code produit, ecarts + / -)
- scan EAN via camera (html5-qrcode)
- vue admin pour consulter tous les mouvements par controleur/date avec detail au click

## Comptes de test
- `admin` / `test`
- `test1` / `test`
- `test2` / `test`
- `test3` / `test`

## Lancer le projet
Depuis le dossier du projet:

```powershell
python -m pip install -r requirements.txt --target .deps
python app.py
```

Ouvrir:
- `http://127.0.0.1:8000`

## Notes Windows / ODBC
- Les actions `Donnes` ouvrent des fenetres Windows locales (save dialog, file picker, ODBC Administrator).
- Pour les connexions `ODBC` et `Access`, installer aussi `pyodbc` via `requirements.txt`.

## Securite incluse
- Hash mots de passe (Werkzeug PBKDF2)
- Session cookie `HttpOnly` + `SameSite=Strict`
- Protection CSRF pour formulaires/API
- Role-based access (`admin`, `controller`)
- Validation stricte cote serveur
- Login rate-limit anti brute-force (verrou 5 min)
- Headers securite (CSP, frame deny, nosniff, referrer policy)

## Notes production
- Definir `INTRANET_SECRET_KEY` avec une cle forte
- Definir `INTRANET_COOKIE_SECURE=1` derriere HTTPS
- Mettre l'app derriere reverse proxy interne (IIS/Nginx)
