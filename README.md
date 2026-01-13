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

Azure deploy (Linux Web App with Azure CLI)

1. Login and create a resource group and app plan (replace names):

```bash
az login
az group create --name myResourceGroup --location japaneast
az appservice plan create --name myPlan --resource-group myResourceGroup --sku B1 --is-linux
```

2. Create a Web App for containers (replace names):

```bash
az webapp create --resource-group myResourceGroup --plan myPlan --name <your-app-name> --runtime "PYTHON|3.11"
```

3. Deploy from local git or zip deploy. Quick zip deploy:

```bash
zip -r app.zip .
az webapp deployment source config-zip --resource-group myResourceGroup --name <your-app-name> --src app.zip
```

4. Configure `DISABLE_WINDOWS_AUTOMATION` (recommended) and other app settings:

```bash
az webapp config appsettings set --resource-group myResourceGroup --name <your-app-name> --settings DISABLE_WINDOWS_AUTOMATION=true
```

5. Open the site: `https://<your-app-name>.azurewebsites.net`

Notes:
- This app includes Windows-only Outlook automation; keep `DISABLE_WINDOWS_AUTOMATION=true` on Linux.
- Use the provided `Procfile` which runs `gunicorn`.

Note for CI: Windows-only dependencies are split into `requirements-windows.txt`. CI runners on Linux should install only `requirements.txt` to avoid installing `pywin32`.
