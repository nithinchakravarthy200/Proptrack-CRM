# PropTrack CRM — Railway Deployment

## Files in this repo
```
proptrack-crm/
├── app.py              ← Flask backend API
├── index.html          ← Frontend CRM UI
├── requirements.txt    ← Python dependencies
├── Procfile            ← Railway start command
├── railway.json        ← Railway config
└── .gitignore
```

## Environment Variables (set on Railway)
| Key | Value |
|-----|-------|
| `GDRIVE_FILE_ID` | `1LftBxyfZ0aUuRjwUGN16zJq8Tr3W6Anc` |

## Local development
```bash
pip install -r requirements.txt
python app.py
```
Open: http://localhost:5000
