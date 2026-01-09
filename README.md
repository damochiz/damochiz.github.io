Mail Create Flask App

Simple Flask app that provides a UI similar to the attached screenshots:
- Left column: form fields to create mail content
- Right column: large preview panel

Quick start:

1. Create a virtualenv and activate it (Windows PowerShell):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

2. Open http://127.0.0.1:5000 in your browser.

Files:
- `app.py`: Flask server
- `templates/index.html`: UI
- `static/main.js`: client logic
