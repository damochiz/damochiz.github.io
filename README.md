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

Windows Outlook automation:
- By default the app will enable Outlook automation when running on Windows. To explicitly disable automation set `DISABLE_WINDOWS_AUTOMATION=true` as an environment variable.
	To enable explicitly (if previously disabled): `setx DISABLE_WINDOWS_AUTOMATION 0` and restart your shell.

Default template directory:
- If no environment variable or `config.json` sets `template_dir`, the app now defaults to `createmailapp/template_files` under the current user's home directory (for example: `C:\Users\<username>\createmailapp\template_files`).

Per-user template directories:
- If the incoming HTTP request contains an authenticated username (in `REMOTE_USER`, `X-Remote-User`, `X-Forwarded-User`, `X-Username`, basic auth username, or a `username` cookie), the app will resolve the template directory for that user as `C:\Users\<username>\createmailapp\template_files` on Windows (or `/home/<username>/createmailapp/template_files` on Unix). The username is sanitized to remove unsafe characters.
- Security note: Only use this feature if your front door (reverse proxy / authentication layer) supplies a trusted username header or authentication context. Do NOT rely on untrusted client-sent headers unless you control the proxy.

Template Dir button on headless hosts (App Service / containers):
- On hosts without a GUI (for example Azure App Service Linux or container instances) the native folder picker cannot run. Using the `Template Dir` button will return a JSON response containing `status: "unavailable"` and the currently resolved `template_dir`.
- To change the template directory on headless hosts, either POST to `/config` with JSON `{ "template_dir": "<path>" }` or set the environment variable `CREATEMAIL_TEMPLATE_DIR` (or `CREATE_MAIL_TEMPLATE_DIR`) and restart the app.

Note about Windows paths on non-Windows hosts:
- If you set a Windows-style path (for example `C:\Users\you\createmailapp\template_files`) from a headless host (Azure App Service), the server will now preserve that exact string in `config.json` rather than attempting to call `os.path.abspath` on it (which previously caused `/tmp/.../C:\...` to appear).
