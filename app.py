from flask import Flask, render_template, request, jsonify
import os
import platform
import json
import re
from datetime import datetime, timedelta, date
from typing import Any, cast
import traceback
import threading
import time
import hashlib
import subprocess
import zipfile
import io
import base64
from functools import lru_cache


def dbg(msg):
    try:
        tf = os.path.join(BASE_DIR, 'render_debug.log')
        with open(tf, 'a', encoding='utf-8') as f:
            f.write(msg + '\n')
    except Exception:
        try:
            import sys
            sys.stderr.write(msg + '\n')
        except Exception:
            pass

app = Flask(__name__)

# Blob storage feature toggle: set AZURE_BLOB_TEMPLATES=1 and AZURE_STORAGE_ACCOUNT to enable
AZURE_BLOB_TEMPLATES = os.environ.get('AZURE_BLOB_TEMPLATES', '0') == '1'
AZURE_STORAGE_ACCOUNT = os.environ.get('AZURE_STORAGE_ACCOUNT')
# Single container name (default mail-templates)
MAIL_TEMPLATES_CONTAINER = os.environ.get('MAIL_TEMPLATES_CONTAINER', 'mail-templates')



def _import_blob_clients():
    try:
        from azure.identity import DefaultAzureCredential
        from azure.storage.blob import BlobServiceClient
        return DefaultAzureCredential, BlobServiceClient
    except Exception:
        return None, None


def _sanitize_token(uid: str) -> str:
    # produce a filesystem/blob-safe token from user id
    if not uid:
        return 'anonymous'
    s = uid.lower()
    import re
    s = re.sub(r'[^a-z0-9-._]', '-', s)
    s = re.sub(r'-+', '-', s).strip('-')
    if len(s) < 1:
        s = 'u'
    # trim to reasonable length
    if len(s) > 64:
        s = s[:64]
    return s


@lru_cache(maxsize=32)
def get_blob_service_client():
    if not blob_mode_enabled():
        return None
    DefaultAzureCredential, BlobServiceClient = _import_blob_clients()
    if not DefaultAzureCredential or not BlobServiceClient:
        raise RuntimeError('Azure Blob client libraries are required when AZURE_BLOB_TEMPLATES=1')
    # allow runtime override of storage account via config.json
    cfg = get_storage_config()
    account = cfg.get('storage_account') or AZURE_STORAGE_ACCOUNT
    if not account:
        raise RuntimeError('AZURE_STORAGE_ACCOUNT not configured for blob access')
    account_url = f"https://{account}.blob.core.windows.net"
    cred = DefaultAzureCredential()
    return BlobServiceClient(account_url=account_url, credential=cred)


def _get_request_user_id():
    # derive a user-identifying string from request context (prefer headers set by Easy Auth or proxy)
    try:
        from flask import has_request_context, request
        if not has_request_context():
            return None
        uid = None
        # common headers
        uid = request.headers.get('X-MS-CLIENT-PRINCIPAL-NAME') or request.headers.get('X-MS-CLIENT-PRINCIPAL-ID')
        if not uid:
            uid = request.headers.get('X-Remote-User') or request.headers.get('X-Forwarded-User') or request.headers.get('X-Username')
        try:
            auth = getattr(request, 'authorization', None)
            if not uid and auth and getattr(auth, 'username', None):
                uid = auth.username
        except Exception:
            pass
        try:
            if not uid:
                uid = request.cookies.get('username')
        except Exception:
            pass
        if isinstance(uid, str) and uid:
            return uid
    except Exception:
        pass
    return None


def _get_shared_container():
    # single shared container for all templates; prefer runtime config if set
    try:
        cfg = get_storage_config()
        cont = cfg.get('container')
        if cont:
            return cont
    except Exception:
        pass
    return MAIL_TEMPLATES_CONTAINER


def blob_mode_enabled():
    # Determine blob mode solely from runtime UI config (do not depend on env)
    try:
        cfg = get_storage_config()
        return (cfg.get('mode') == 'blob')
    except Exception:
        return False


def _blob_list_templates_for_user(user_token: str):
    svc = get_blob_service_client()
    if not svc:
        return []
    container = svc.get_container_client(_get_shared_container())
    prefix = f'template_files/{user_token}/'
    try:
        blobs = container.list_blobs(name_starts_with=prefix)
    except Exception:
        return []
    items = []
    for b in blobs:
        # strip prefix
        name = b.name[len(prefix):] if b.name.startswith(prefix) else b.name
        if not name:
            continue
        if name.lower().endswith('.json'):
            items.append(name)
    return items


def _blob_get_text_for_user(user_token: str, blob_name: str):
    svc = get_blob_service_client()
    if not svc:
        return None
    container = svc.get_container_client(_get_shared_container())
    blob_client = container.get_blob_client(f'template_files/{user_token}/{blob_name}')
    try:
        stream = blob_client.download_blob()
        data = stream.readall()
        return data.decode('utf-8')
    except Exception:
        return None


def _blob_delete_for_user(user_token: str, blob_name: str):
    try:
        svc = get_blob_service_client()
        if not svc:
            return False
        container = svc.get_container_client(_get_shared_container())
        blob_client = container.get_blob_client(f'template_files/{user_token}/{blob_name}')
        try:
            blob_client.delete_blob()
            return True
        except Exception:
            return False
    except Exception:
        return False


def _blob_upload_text_for_user(user_token: str, blob_name: str, text: str):
    try:
        dbg(f'blob upload: user_token={user_token} blob_name={blob_name}')
        svc = get_blob_service_client()
        dbg(f'blob upload: svc={None if svc is None else type(svc).__name__}')
        if not svc:
            raise RuntimeError('blob service not available')
        container_name = _get_shared_container()
        dbg(f'blob upload: container={container_name}')
        container = svc.get_container_client(container_name)
        # create container if not exists (idempotent)
        try:
            container.create_container()
        except Exception:
            pass
        blob_client = container.get_blob_client(f'template_files/{user_token}/{blob_name}')
        blob_client.upload_blob(text.encode('utf-8'), overwrite=True)
        dbg(f'blob upload: success user_token={user_token} blob_name={blob_name}')
        return True
    except Exception as e:
        try:
            tb = traceback.format_exc()
            dbg(f'blob upload failed: {e}')
            dbg(tb)
        except Exception:
            try:
                dbg(f'blob upload failed (no traceback): {e}')
            except Exception:
                pass
        return False


# Short-lived in-memory cache to deduplicate near-duplicate schedule requests.
# Keys are SHA256(title+body+start+end) -> timestamp (time.time()).
recent_schedules = {}
recent_schedules_lock = threading.Lock()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')


def get_storage_config():
    # Returns a dict with keys: mode ('local'|'blob'), storage_account, container
    cfg = {}
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                cfg = json.load(f) or {}
    except Exception:
        cfg = {}
    out = {
        'mode': cfg.get('storage_mode', 'local'),
        'storage_account': cfg.get('storage_account') or os.environ.get('AZURE_STORAGE_ACCOUNT'),
        'container': cfg.get('mail_templates_container') or os.environ.get('MAIL_TEMPLATES_CONTAINER', 'mail-templates')
    }
    return out


def save_storage_config(mode: str, storage_account: str | None, container: str | None):
    try:
        try:
            cfg = {}
            if os.path.exists(CONFIG_PATH):
                with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                    cfg = json.load(f) or {}
        except Exception:
            cfg = {}
        cfg['storage_mode'] = mode if mode in ('local', 'blob') else 'local'
        if storage_account:
            cfg['storage_account'] = storage_account
        else:
            cfg.pop('storage_account', None)
        if container:
            cfg['mail_templates_container'] = container
        else:
            cfg.pop('mail_templates_container', None)
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def get_template_dir():
    # Priority: env var -> config.json -> default workspace/template_files
    env = os.environ.get('CREATEMAIL_TEMPLATE_DIR') or os.environ.get('CREATE_MAIL_TEMPLATE_DIR')
    if env:
        try:
            if isinstance(env, str):
                return os.path.abspath(os.path.expanduser(env))
        except Exception:
            pass
    # try config.json
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
                if isinstance(cfg, dict):
                    td = cfg.get('template_dir')
                    if isinstance(td, str) and td:
                        # If the configured path is a Windows absolute path (e.g. C:\...) but
                        # this process is running on non-Windows, do not call os.path.abspath
                        # (which would prepend the Linux cwd and produce /tmp/.../C:\...).
                        try:
                            import re
                            is_win_abs = bool(re.match(r'^[A-Za-z]:[\\/]|^\\\\', td))
                        except Exception:
                            is_win_abs = False
                        if is_win_abs and platform.system() != 'Windows':
                            return td
                        return os.path.abspath(os.path.expanduser(td))
    except Exception:
        pass
    # fallback default: if there is a request context, try to resolve the visiting
    # user's account name from common headers/environ/cookies/auth and build
    # a per-user path like C:\Users\<username>\createmailapp\template_files
    try:
        from flask import has_request_context, request
        if has_request_context():
            username = None
            try:
                # Common sources for an authenticated username
                username = request.environ.get('REMOTE_USER')
            except Exception:
                username = None
            if not username:
                username = request.headers.get('X-Remote-User') or request.headers.get('X-Forwarded-User') or request.headers.get('X-Username')
            try:
                auth = getattr(request, 'authorization', None)
                if not username and auth and getattr(auth, 'username', None):
                    username = auth.username
            except Exception:
                pass
            try:
                if not username:
                    username = request.cookies.get('username')
            except Exception:
                pass

            if isinstance(username, str) and username:
                # sanitize username to safe filesystem token
                try:
                    import re
                    uname = re.sub(r'[^A-Za-z0-9_.-]', '', username)
                except Exception:
                    uname = username
                if uname:
                    if os.name == 'nt':
                        default_dir = os.path.join('C:\\Users', uname, 'createmailapp', 'template_files')
                    else:
                        default_dir = os.path.join('/home', uname, 'createmailapp', 'template_files')
                    return os.path.abspath(default_dir)
    except Exception:
        pass

    # final fallback: current user's home createmailapp/template_files, then repo template_files
    try:
        home = os.path.expanduser('~')
        if home:
            default_dir = os.path.join(home, 'createmailapp', 'template_files')
            return os.path.abspath(default_dir)
    except Exception:
        pass
    return os.path.join(BASE_DIR, 'template_files')

def atomic_write_json(path, data):
    # write JSON atomically
    tmp = str(path) + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    try:
        os.replace(tmp, path)
    except Exception:
        # best-effort fallback
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

def safe_type_filename(email_type: str) -> str:
    # create a filesystem-safe filename for an email type
    if not email_type:
        return 'default'
    # keep the original text but remove characters illegal in filenames (Windows/most OS)
    name = email_type.strip()
    import re
    # replace runs of whitespace with underscore
    name = re.sub(r'\s+', '_', name)
    # remove characters that are invalid in Windows filenames and control chars
    name = re.sub(r'[<>:"/\\|\?\*\x00-\x1f]', '', name)
    # collapse consecutive underscores or hyphens
    name = re.sub(r'[_\-]+', '_', name)
    # trim leading/trailing separators and spaces
    name = name.strip(' _-')
    # limit length to reasonable filename size (keep 120 chars)
    if len(name) > 120:
        name = name[:120]
    return name or 'template'

def ensure_templates_dir():
    try:
        os.makedirs(get_template_dir(), exist_ok=True)
    except Exception:
        pass

    # If a legacy file named 'fcs_.json' exists (from older safe filename logic),
    # prefer its content and rename it to 'fcs.json' to normalize storage.
    try:
        legacy = os.path.join(get_template_dir(), 'fcs_.json')
        canonical = os.path.join(get_template_dir(), 'fcs.json')
        if os.path.exists(legacy) and os.path.exists(canonical):
            # read legacy and overwrite canonical with legacy content
            try:
                with open(legacy, 'r', encoding='utf-8-sig') as f:
                    data = f.read()
                with open(canonical, 'w', encoding='utf-8-sig') as f:
                    f.write(data)
                try:
                    os.remove(legacy)
                except Exception:
                    pass
            except Exception:
                pass
        elif os.path.exists(legacy) and not os.path.exists(canonical):
            try:
                os.replace(legacy, canonical)
            except Exception:
                try:
                    # fallback to copy+remove
                    with open(legacy, 'r', encoding='utf-8-sig') as f:
                        data = f.read()
                    with open(canonical, 'w', encoding='utf-8-sig') as f:
                        f.write(data)
                    os.remove(legacy)
                except Exception:
                    pass
    except Exception:
        pass

def template_file_path_for(email_type: str) -> str:
    ensure_templates_dir()
    fname = safe_type_filename(email_type) + '.json'
    return os.path.join(get_template_dir(), fname)


def load_templates_file():
    # legacy aggregated templates file (template.json) loader
    try:
        path = os.path.join(get_template_dir(), 'template.json')
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8-sig') as f:
                data = json.load(f)
                if isinstance(data, dict):
                    return data
    except Exception:
        pass
    return {}

def _ensure_templates_dir():
    ensure_templates_dir()
    # Also perform immediate normalization when helper is invoked directly
    try:
        legacy = os.path.join(get_template_dir(), 'fcs_.json')
        canonical = os.path.join(get_template_dir(), 'fcs.json')
        if os.path.exists(legacy):
            try:
                with open(legacy, 'r', encoding='utf-8-sig') as f:
                    data = f.read()
                with open(canonical, 'w', encoding='utf-8-sig') as f:
                    f.write(data)
                try:
                    os.remove(legacy)
                except Exception:
                    pass
            except Exception:
                pass
    except Exception:
        pass

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/preview', methods=['POST'])
def preview():
    data = request.json or {}
    # simple preview payload
    sr = data.get('srNumber', '')
    customer = data.get('customerName', '')
    title = data.get('title', '')
    content = data.get('content', '')
    email_type = data.get('emailType', '')
    preview_text = f"SR: {sr}\nCustomer: {customer}\nType: {email_type}\nTitle: {title}\n\n{content}"
    return jsonify({'preview': preview_text})


DISABLE_WINDOWS_AUTOMATION = os.environ.get('DISABLE_WINDOWS_AUTOMATION')
if DISABLE_WINDOWS_AUTOMATION is None:
    # Default to enabled automation on Windows, disabled elsewhere
    DISABLE_WINDOWS_AUTOMATION = '0' if platform.system() == 'Windows' else '1'


@app.route('/create_mail', methods=['POST'])
def create_mail():
    data = request.json or {}
    # prefer the largePanel content if provided
    body = data.get('content') or data.get('preview') or ''
    subject = data.get('title') or data.get('emailType') or ''

    # Only attempt Outlook automation on Windows and when not disabled by env
    if DISABLE_WINDOWS_AUTOMATION.lower() in ('1', 'true', 'yes') or platform.system() != 'Windows':
        return jsonify({'status': 'error', 'message': 'Outlook automation is disabled or only supported on Windows.'}), 400

    try:
        import win32com.client
        import pythoncom

        # Create the mail item on a background thread so the HTTP request
        # can return immediately instead of blocking while the Outlook
        # window is open. Initialize COM on the new thread and uninitialize
        # when done. Use non-modal Display(False) to avoid blocking the
        # thread while still showing the compose window to the user.
        def _create_mail_item(sbj, bd):
            try:
                pythoncom.CoInitialize()
                try:
                    outlook = win32com.client.Dispatch('Outlook.Application')
                    mail = outlook.CreateItem(0)  # 0: olMailItem
                    if sbj:
                        mail.Subject = sbj
                    mail.Body = bd
                    try:
                        # Open non-modal if supported
                        mail.Display(False)
                    except Exception:
                        # Fallback to default Display (may be modal on some setups)
                        try:
                            mail.Display()
                        except Exception:
                            pass
                finally:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
            except Exception:
                # Swallow exceptions in background thread; nothing to return
                pass

        t = threading.Thread(target=_create_mail_item, args=(subject, body), daemon=True)
        t.start()
        return jsonify({'status': 'ok', 'message': 'Mail creation started.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Failed to create mail: {e}'}), 500


@app.route('/create_schedule', methods=['POST'])
def create_schedule():
    data = request.json or {}
    title = data.get('title', '')
    body = data.get('body', '')
    start_iso = data.get('start_iso')
    end_iso = data.get('end_iso')
    if not start_iso or not end_iso:
        return jsonify({'status': 'error', 'message': 'start_iso and end_iso are required'}), 400

    # Only attempt Outlook automation on Windows and when not disabled by env
    if DISABLE_WINDOWS_AUTOMATION.lower() in ('1', 'true', 'yes') or platform.system() != 'Windows':
        return jsonify({'status': 'error', 'message': 'Outlook scheduling is disabled or only supported on Windows.'}), 400

    try:
        # compute a short-lived dedupe key for this schedule request
        try:
            h_in = (title or '') + '||' + (body or '') + '||' + (start_iso or '') + '||' + (end_iso or '')
            key = hashlib.sha256(h_in.encode('utf-8')).hexdigest()
        except Exception:
            key = None

        if key:
            nowt = time.time()
            with recent_schedules_lock:
                # remove expired entries
                expired = [k for k, v in recent_schedules.items() if nowt - v > 30]
                for ex in expired:
                    try:
                        recent_schedules.pop(ex, None)
                    except Exception:
                        pass
                if key in recent_schedules:
                    dbg(f"create_schedule duplicate blocked key={key}")
                    return jsonify({'status': 'error', 'message': 'Duplicate schedule request blocked'}), 429
                recent_schedules[key] = nowt
        import win32com.client
        import pythoncom

        def _create_appointment(sbj, bd, st_iso, en_iso):
            # Simplified COM lifecycle: initialize, perform actions, always uninitialize
            try:
                pythoncom.CoInitialize()
            except Exception:
                # if CoInitialize fails, bail out
                try:
                    dbg('pythoncom.CoInitialize failed')
                except Exception:
                    pass
                return

            try:
                # Main appointment creation logic encapsulated in a try/except
                try:
                    dbg(f"create_schedule received start_iso={st_iso} end_iso={en_iso}")
                except Exception:
                    pass

                try:
                    outlook = win32com.client.Dispatch('Outlook.Application')
                    appt = outlook.CreateItem(1)  # olAppointmentItem
                except Exception:
                    try:
                        dbg('Failed to create Outlook appointment item')
                    except Exception:
                        pass
                    return

                # Subject
                try:
                    if sbj:
                        s = sbj.replace('\r', '\n')
                        s = s.replace('\n', ' ').strip()
                        appt.Subject = s
                except Exception:
                    try:
                        appt.Subject = (sbj or '').strip()
                    except Exception:
                        pass

                # Body representations
                def _normalize_plain_text(x):
                    try:
                        if not x:
                            return ''
                        t = x.replace('\r\n', '\n').replace('\r', '\n')
                        parts = t.split('\n')
                        while parts and parts[-1].strip() == '':
                            parts.pop()
                        return '\r\n'.join(parts)
                    except Exception:
                        return x or ''

                plain = _normalize_plain_text(bd)
                # Ensure CRLF for plain text to give Outlook the expected line endings
                def _ensure_crlf(x):
                    try:
                        if not x:
                            return ''
                        t = x.replace('\r\n', '\n').replace('\r', '\n')
                        return '\r\n'.join(t.split('\n'))
                    except Exception:
                        return x or ''

                plain_crlf = _ensure_crlf(plain)
                try:
                    dbg(f"plain_crlf repr={repr(plain_crlf)[:400]}")
                except Exception:
                    pass
                html_wrapped = None
                body_html = ''
                try:
                    def html_escape(s):
                        return (s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

                    def _url_to_anchor(m):
                        url = m.group(0)
                        href = url if re.match(r'^[a-zA-Z][a-zA-Z0-9+.-]*://', url) else 'http://' + url
                        safe = href.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        return f'<a href="{safe}">{html_escape(url)}</a>'

                    url_re = re.compile(r'(?:(?:https?://)|(?:www\.)|(?:[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}))(?:[\w\-\./?%&=+#~:,;@!$\(\)\[\]\\]*)')
                    out_lines = []
                    # Use the CRLF-normalized plain and split on CRLF for HTML lines
                    for ln in plain_crlf.split('\r\n'):
                        esc = html_escape(ln)
                        anchored = url_re.sub(_url_to_anchor, esc)
                        out_lines.append(anchored)
                    body_html = '<br/>'.join(out_lines)
                    # Prefer Yu Gothic at 10pt, fall back to Meiryo / MS PGothic / sans-serif
                    font_css = "font-family: 'Yu Gothic', 'YuGothic', 'Yu Gothic UI', 'Meiryo', 'MS PGothic', sans-serif; font-size:10pt;"
                    html_wrapped = f"<div style=\"{font_css} white-space:normal; line-height:1.35; margin:0; padding:0;\">{body_html}</div>"
                except Exception:
                    html_wrapped = None

                # Set plain body first
                try:
                    appt.Body = plain_crlf
                except Exception:
                    try:
                        appt.Body = _ensure_crlf(bd or '')
                    except Exception:
                        pass

                # Date parsing helper
                def parse_iso_to_local_naive(s):
                    if not s:
                        return None
                    try:
                        if s.endswith('Z') or '+' in s[10:] or '-' in s[10:]:
                            s2 = s.replace('Z', '+00:00') if s.endswith('Z') else s
                            dt = datetime.fromisoformat(s2)
                            if dt.tzinfo is not None:
                                try:
                                    dt = dt.astimezone().replace(tzinfo=None)
                                except Exception:
                                    dt = dt.replace(tzinfo=None)
                            return dt
                        try:
                            return datetime.strptime(s, '%Y-%m-%dT%H:%M:%S')
                        except Exception:
                            try:
                                base = s.split('.')[0]
                                return datetime.strptime(base, '%Y-%m-%dT%H:%M:%S')
                            except Exception:
                                return None
                    except Exception:
                        return None

                st_dt = parse_iso_to_local_naive(st_iso)
                en_dt = parse_iso_to_local_naive(en_iso)
                try:
                    dbg(f"parsed start_dt={st_dt} end_dt={en_dt}")
                except Exception:
                    pass

                if st_dt:
                    try:
                        import pywintypes
                        if isinstance(st_dt, datetime):
                            appt.Start = pywintypes.Time(cast(Any, st_dt))
                        try:
                            appt.Start = st_dt.strftime('%Y-%m-%d %H:%M:%S')
                        except Exception:
                            pass
                    except Exception:
                        try:
                            appt.Start = st_dt
                        except Exception:
                            try:
                                appt.Start = st_dt.strftime('%Y-%m-%d %H:%M')
                            except Exception:
                                pass

                if en_dt:
                    try:
                        import pywintypes
                        if isinstance(en_dt, datetime):
                            appt.End = pywintypes.Time(cast(Any, en_dt))
                        try:
                            appt.End = en_dt.strftime('%Y-%m-%d %H:%M:%S')
                        except Exception:
                            pass
                    except Exception:
                        try:
                            appt.End = en_dt
                        except Exception:
                            try:
                                appt.End = en_dt.strftime('%Y-%m-%d %H:%M')
                            except Exception:
                                pass

                # Reminder and display
                try:
                    try:
                        appt.ReminderSet = True
                        appt.ReminderMinutesBeforeStart = 15
                    except Exception:
                        pass

                    try:
                        # Log EntryID before display (may be None/raise if not saved)
                        try:
                            entry_before = getattr(appt, 'EntryID', None)
                            dbg(f"appt.EntryID before display = {entry_before}")
                        except Exception:
                            entry_before = None
                        appt.Display(False)
                        try:
                            try:
                                insp = appt.GetInspector
                                try:
                                    editor = insp.WordEditor
                                    try:
                                        editor.Content.Font.Name = 'Yu Gothic'
                                        editor.Content.Font.Size = 10
                                        dbg('Set WordEditor font to Yu Gothic 10pt')
                                    except Exception:
                                        dbg('Failed to set editor font properties')
                                except Exception:
                                    pass
                            except Exception:
                                pass
                        except Exception:
                            pass
                    except Exception:
                        try:
                            appt.Display()
                        except Exception:
                            pass

                    try:
                        import time as _time
                        _time.sleep(0.25)
                    except Exception:
                        pass

                    if html_wrapped:
                        try:
                            try:
                                # 2 corresponds to olFormatHTML in Outlook constants
                                appt.BodyFormat = 2
                            except Exception:
                                pass
                            # strengthen font specification by wrapping in <font> as Outlook may prefer font tags
                            try:
                                font_wrapper = f"<font face=\"Yu Gothic, YuGothic, 'Yu Gothic UI', Meiryo, 'MS PGothic'\" style=\"font-size:10pt; line-height:1.4;\">{body_html}</font>"
                                full_html = f"<div style=\"font-family: Yu Gothic, YuGothic, 'Yu Gothic UI', Meiryo, 'MS PGothic', sans-serif; font-size:10pt; line-height:1.4; margin:0; padding:0;\">{font_wrapper}</div>"
                            except Exception:
                                full_html = html_wrapped
                            appt.HTMLBody = full_html
                            # Log EntryID after HTML update to help detect duplicates
                            try:
                                entry_after_html = getattr(appt, 'EntryID', None)
                                dbg(f"appt.EntryID after HTMLBody set = {entry_after_html}")
                            except Exception:
                                entry_after_html = None
                            # Do not call Save unconditionally — calling Save after Display can cause duplicate items
                            # If the item appears to be unsaved (EntryID is None or empty), attempt Save once.
                            try:
                                entry_post = getattr(appt, 'EntryID', None)
                            except Exception:
                                entry_post = None
                            try:
                                if not entry_post:
                                    try:
                                        appt.Save()
                                        dbg(f"appt.Save() called for tentative EntryID after save = {getattr(appt, 'EntryID', None)}")
                                    except Exception:
                                        dbg('appt.Save() failed or raised')
                                else:
                                    dbg(f"appt already had EntryID, skipping Save: {entry_post}")
                            except Exception:
                                pass
                        except Exception:
                            pass
                except Exception:
                    try:
                        dbg('Error while displaying or saving appointment')
                    except Exception:
                        pass

            except Exception as e:
                try:
                    dbg(f"_create_appointment unexpected error: {e}")
                except Exception:
                    pass
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                # remove dedupe key so subsequent identical requests are allowed after work is done
                try:
                    if key:
                        with recent_schedules_lock:
                            try:
                                recent_schedules.pop(key, None)
                            except Exception:
                                pass
                except Exception:
                    pass

        t = threading.Thread(target=_create_appointment, args=(title, body, start_iso, end_iso), daemon=True)
        t.start()
        return jsonify({'status': 'ok', 'message': 'Schedule creation started.'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Failed to create schedule: {e}'}), 500


@app.route('/templates', methods=['GET'])
def get_templates():
    # Return all templates: prefer local filesystem (Template Dir). If nothing found locally, fall back to Blob per-user.
    data = {}
    try:
        ensure_templates_dir()
        for fn in os.listdir(get_template_dir()):
            if not fn.lower().endswith('.json'):
                continue
            path = os.path.join(get_template_dir(), fn)
            try:
                with open(path, 'r', encoding='utf-8-sig') as f:
                    obj = json.load(f)
                    if isinstance(obj, dict) and 'template' in obj:
                        key = os.path.splitext(fn)[0]
                        data[key] = obj.get('template')
                    elif isinstance(obj, dict):
                        for k, v in obj.items():
                            if isinstance(v, str):
                                data[k] = v
            except Exception:
                continue
    except Exception:
        pass

    # If local templates were found, return them
    if data:
        return jsonify({'status': 'ok', 'templates': data})

    # Otherwise, if blob mode enabled, attempt to read per-user blobs
    if blob_mode_enabled():
        try:
            uid = _get_request_user_id() or 'anonymous'
            user_token = _sanitize_token(uid)
            blobs = _blob_list_templates_for_user(user_token)
            for bname in blobs:
                txt = _blob_get_text_for_user(user_token, bname)
                if not txt:
                    continue
                try:
                    obj = json.loads(txt)
                except Exception:
                    # treat raw string as single template
                    data[os.path.splitext(bname)[0]] = txt
                    continue
                if isinstance(obj, dict) and 'template' in obj:
                    data[os.path.splitext(bname)[0]] = obj.get('template')
                elif isinstance(obj, dict):
                    for k, v in obj.items():
                        if isinstance(v, str):
                            data[k] = v
            return jsonify({'status': 'ok', 'templates': data})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500

    return jsonify({'status': 'ok', 'templates': data})


@app.route('/templates/keys', methods=['GET'])
def get_template_keys():
    # Return mapping of filename -> list of keys inside that file (for multi-key files like fcs.json)
    result = {}
    # Prefer local filesystem; if no local files exist, fall back to blob storage
    try:
        ensure_templates_dir()
        local_found = False
        for fn in os.listdir(get_template_dir()):
            if not fn.lower().endswith('.json'):
                continue
            local_found = True
            path = os.path.join(get_template_dir(), fn)
            try:
                with open(path, 'r', encoding='utf-8-sig') as f:
                    obj = json.load(f)
                    if isinstance(obj, dict):
                        keys = list(obj.keys())
                        result[fn] = keys
            except Exception:
                continue
        if local_found:
            return jsonify({'status': 'ok', 'files': result})
    except Exception:
        pass

    try:
        # filesystem had no files or failed; try blob
        if blob_mode_enabled():
            uid = _get_request_user_id() or 'anonymous'
            user_token = _sanitize_token(uid)
            blobs = _blob_list_templates_for_user(user_token)
            for b in blobs:
                txt = _blob_get_text_for_user(user_token, b)
                if not txt:
                    continue
                try:
                    obj = json.loads(txt)
                except Exception:
                    result[b] = ['template']
                    continue
                if isinstance(obj, dict):
                    result[b] = list(obj.keys())
            return jsonify({'status': 'ok', 'files': result})
        
    except Exception:
        pass
    return jsonify({'status': 'ok', 'files': result})


@app.route('/templates', methods=['POST'])
def save_templates():
    payload = request.json or {}
    # expect a mapping of type -> template string
    templates = payload.get('templates')
    if not isinstance(templates, dict):
        return jsonify({'status': 'error', 'message': 'templates must be an object/dict'}), 400
    # Save per-type files only (filesystem or blob)
    if blob_mode_enabled():
        try:
            uid = _get_request_user_id() or 'anonymous'
            user_token = _sanitize_token(uid)
            ok = True
            for k, v in templates.items():
                fname = safe_type_filename(k) + '.json'
                # store as {'template': v}
                payload_txt = json.dumps({'template': v}, ensure_ascii=False)
                if not _blob_upload_text_for_user(user_token, fname, payload_txt):
                    ok = False
            if ok:
                return jsonify({'status': 'ok'})
            return jsonify({'status': 'error', 'message': 'failed to upload some templates'}), 500
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500

    # filesystem fallback
    success = True
    for k, v in templates.items():
        try:
            path = template_file_path_for(k)
            atomic_write_json(path, {'template': v})
        except Exception:
            success = False
    if success:
        return jsonify({'status': 'ok'})
    return jsonify({'status': 'error', 'message': 'failed to save some files'}), 500


@app.route('/templates/<email_type>', methods=['GET'])
def get_template_for(email_type: str):
    # Optional query parameter to prefer a particular source file (filename only)
    source_file = None
    try:
        source_file = request.args.get('source_file')
        if source_file:
            # sanitize filename: only allow basename and .json extension
            source_file = os.path.basename(source_file)
            if not source_file.lower().endswith('.json'):
                source_file = None
    except Exception:
        source_file = None

    # First: try filesystem (Template Dir)
    try:
        # Try per-type file first (or the explicitly requested source_file)
        if source_file:
            path = os.path.join(get_template_dir(), source_file)
        else:
            path = template_file_path_for(email_type)
        try:
            if os.path.exists(path):
                with open(path, 'r', encoding='utf-8-sig') as f:
                    obj = json.load(f)
                    # If this file is a single-template file
                    if isinstance(obj, dict) and 'template' in obj:
                        return jsonify({'status': 'ok', 'template': obj.get('template'), 'source_file': os.path.basename(path), 'key': email_type})
                    # If a source_file was explicitly requested and it contains the desired key (multi-key mapping), prefer it
                    if source_file and isinstance(obj, dict) and email_type in obj and isinstance(obj[email_type], str):
                        return jsonify({'status': 'ok', 'template': obj[email_type], 'source_file': os.path.basename(path), 'key': email_type})
        except Exception:
            pass

        # Next: try to find the template inside any per-file mapping (e.g. fcs.json contains multiple named templates)
        try:
            ensure_templates_dir()
            # Prefer group file that matches safe filename for this email_type (e.g., 'FCS' -> 'fcs.json')
            expected_group = safe_type_filename(email_type) + '.json'
            expected_path = os.path.join(get_template_dir(), expected_group)
            if os.path.exists(expected_path):
                try:
                    with open(expected_path, 'r', encoding='utf-8-sig') as f:
                        obj = json.load(f)
                        if isinstance(obj, dict):
                            # direct key match in expected group
                            if email_type in obj and isinstance(obj[email_type], str):
                                return jsonify({'status': 'ok', 'template': obj[email_type], 'source_file': expected_group, 'key': email_type})
                            # try common canonical variants
                            key_variants = [email_type, email_type.strip(), email_type.strip().upper(), email_type.strip().title()]
                            for kv in key_variants:
                                if kv in obj and isinstance(obj[kv], str):
                                    return jsonify({'status': 'ok', 'template': obj[kv], 'source_file': expected_group, 'key': kv})
                except Exception:
                    pass
            for fn in os.listdir(get_template_dir()):
                if not fn.lower().endswith('.json'):
                    continue
                path = os.path.join(get_template_dir(), fn)
                try:
                    with open(path, 'r', encoding='utf-8-sig') as f:
                        obj = json.load(f)
                        if isinstance(obj, dict):
                            # If a source_file was requested, prefer it
                            if source_file and fn == source_file:
                                if email_type in obj and isinstance(obj[email_type], str):
                                    return jsonify({'status': 'ok', 'template': obj[email_type], 'source_file': fn, 'key': email_type})
                            # direct key match (case-sensitive as stored)
                            if email_type in obj and isinstance(obj[email_type], str):
                                return jsonify({'status': 'ok', 'template': obj[email_type], 'source_file': fn, 'key': email_type})
                            # try common canonical keys for variants
                            key_variants = [email_type, email_type.strip(), email_type.strip().upper(), email_type.strip().title()]
                            for kv in key_variants:
                                if kv in obj and isinstance(obj[kv], str):
                                    return jsonify({'status': 'ok', 'template': obj[kv], 'source_file': fn, 'key': kv})
                except Exception:
                    continue
        except Exception:
            pass

    except Exception:
        pass

    # If not found locally, and blob mode enabled, try Blob (single shared container, per-user prefix -> shared)
    if AZURE_BLOB_TEMPLATES:
        try:
            # derive user token and shared container name
            user_id = _get_request_user_id() or 'anonymous'
            user_token = _sanitize_token(user_id)
            shared_container = _get_shared_container()

            # prefer requested source_file in user's prefix
            if source_file:
                txt = _blob_get_text_for_user(user_token, source_file)
                if txt:
                    try:
                        obj = json.loads(txt)
                    except Exception:
                        return jsonify({'status': 'ok', 'template': txt, 'source_file': source_file, 'key': email_type})
                    if isinstance(obj, dict) and 'template' in obj:
                        return jsonify({'status': 'ok', 'template': obj.get('template'), 'source_file': source_file, 'key': email_type})
                    if isinstance(obj, dict) and email_type in obj and isinstance(obj[email_type], str):
                        return jsonify({'status': 'ok', 'template': obj[email_type], 'source_file': source_file, 'key': email_type})

            # try direct filename match in user's prefix
            pref = safe_type_filename(email_type) + '.json'
            txt = _blob_get_text_for_user(user_token, pref)
            if txt:
                try:
                    obj = json.loads(txt)
                except Exception:
                    return jsonify({'status': 'ok', 'template': txt, 'source_file': pref, 'key': email_type})
                if isinstance(obj, dict):
                    if email_type in obj and isinstance(obj[email_type], str):
                        return jsonify({'status': 'ok', 'template': obj[email_type], 'source_file': pref, 'key': email_type})
                    # variants
                    for kv in [email_type, email_type.strip(), email_type.strip().upper(), email_type.strip().title()]:
                        if kv in obj and isinstance(obj[kv], str):
                            return jsonify({'status': 'ok', 'template': obj[kv], 'source_file': pref, 'key': kv})

            # try scanning all user blobs
            blobs = _blob_list_templates_for_user(user_token)
            for b in blobs:
                txt = _blob_get_text_for_user(user_token, b)
                if not txt:
                    continue
                try:
                    obj = json.loads(txt)
                except Exception:
                    if os.path.splitext(b)[0] == email_type:
                        return jsonify({'status': 'ok', 'template': txt, 'source_file': b, 'key': email_type})
                    continue
                if isinstance(obj, dict):
                    if email_type in obj and isinstance(obj[email_type], str):
                        return jsonify({'status': 'ok', 'template': obj[email_type], 'source_file': b, 'key': email_type})
                    for kv in [email_type, email_type.strip(), email_type.strip().upper(), email_type.strip().title()]:
                        if kv in obj and isinstance(obj[kv], str):
                            return jsonify({'status': 'ok', 'template': obj[kv], 'source_file': b, 'key': kv})

            # fallback to shared container (no user prefix)
            blobs = _blob_list_templates(shared_container)
            for b in blobs:
                txt = _blob_get_text(shared_container, b)
                if not txt:
                    continue
                try:
                    obj = json.loads(txt)
                except Exception:
                    if os.path.splitext(b)[0] == email_type:
                        return jsonify({'status': 'ok', 'template': txt, 'source_file': b, 'key': email_type})
                    continue
                if isinstance(obj, dict):
                    if email_type in obj and isinstance(obj[email_type], str):
                        return jsonify({'status': 'ok', 'template': obj[email_type], 'source_file': b, 'key': email_type})
                    for kv in [email_type, email_type.strip(), email_type.strip().upper(), email_type.strip().title()]:
                        if kv in obj and isinstance(obj[kv], str):
                            return jsonify({'status': 'ok', 'template': obj[kv], 'source_file': b, 'key': kv})
        except Exception:
            pass

    # Fallback to legacy aggregated store
    cur = load_templates_file()
    key = safe_type_filename(email_type)
    # aggregated file keys may be original names; try direct
    if email_type in cur:
        return jsonify({'status': 'ok', 'template': cur[email_type]})
    if key in cur:
        return jsonify({'status': 'ok', 'template': cur[key]})
    return jsonify({'status': 'error', 'message': 'not found'}), 404


@app.route('/templates/<email_type>', methods=['POST'])
def save_template_for(email_type: str):
    payload = request.json or {}
    template = payload.get('template')
    if not isinstance(template, str):
        return jsonify({'status': 'error', 'message': 'template must be a string'}), 400
    # Decide target group file. Group FCS-related types into 'fcs.json'.
    try:
        et = (email_type or '').strip()
        lower = et.lower()
        if lower.startswith('fcs'):
            group_base = 'fcs'
        else:
            group_base = safe_type_filename(et)

        group_path = template_file_path_for(group_base)
        # original per-type path (may exist if previously saved separately)
        orig_path = template_file_path_for(email_type)

        # load group existing data
        existing = {}
        if os.path.exists(group_path):
            try:
                with open(group_path, 'r', encoding='utf-8-sig') as f:
                    loaded = json.load(f)
                    if isinstance(loaded, dict):
                        existing = loaded
            except Exception:
                existing = {}

        # if an original per-type file exists and is different from group, merge it in and remove it
        try:
            if orig_path != group_path and os.path.exists(orig_path):
                try:
                    with open(orig_path, 'r', encoding='utf-8-sig') as f:
                        loaded = json.load(f)
                        if isinstance(loaded, dict):
                            # merge keys from orig into existing (orig keys may be 'template' or named keys)
                            for k, v in loaded.items():
                                if isinstance(v, str):
                                    existing[k] = v
                except Exception:
                    pass
                try:
                    os.remove(orig_path)
                except Exception:
                    pass
        except Exception:
            pass

        # set/update the key for the selected email_type (preserve other keys)
        existing[email_type] = template
        if 'template' in existing and len(existing) > 1:
            existing.pop('template', None)

        if blob_mode_enabled():
            # upload group JSON to user's prefix in shared container
            user_id = _get_request_user_id() or 'anonymous'
            user_token = _sanitize_token(user_id)
            fname = os.path.basename(group_path)
            payload_txt = json.dumps(existing, ensure_ascii=False)
            if not _blob_upload_text_for_user(user_token, fname, payload_txt):
                return jsonify({'status': 'error', 'message': 'failed to upload to blob storage'}), 500
        else:
            atomic_write_json(group_path, existing)
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'failed to save: {e}'}), 500
    # per-file saved; no aggregated store
    return jsonify({'status': 'ok'})


@app.route('/templates/save_in_file_explicit', methods=['POST'])
def save_template_in_file():
    payload = request.json or {}
    template = payload.get('template')
    source_file = payload.get('source_file')
    key = payload.get('key')
    if not isinstance(template, str):
        return jsonify({'status': 'error', 'message': 'template must be a string'}), 400
    if not source_file:
        return jsonify({'status': 'error', 'message': 'source_file required to save in a specific file'}), 400
    try:
        ensure_templates_dir()
        path = os.path.join(get_template_dir(), source_file)
        # only allow updating existing files to avoid accidental new-file creation
        if not os.path.exists(path):
            return jsonify({'status': 'error', 'message': f'specified source_file does not exist: {source_file}'}), 400
        try:
            with open(path, 'r', encoding='utf-8-sig') as f:
                obj = json.load(f)
        except Exception:
            obj = {}
        # if file is a dict of named templates and key provided, update that key
        if isinstance(obj, dict) and key:
            obj[key] = template
            atomic_write_json(path, obj)
            return jsonify({'status': 'ok'})
        # if file stores a single template under 'template', overwrite it
        if isinstance(obj, dict) and 'template' in obj:
            obj['template'] = template
            atomic_write_json(path, obj)
            return jsonify({'status': 'ok'})
        # if file exists but doesn't match expected shapes, refuse to overwrite blindly
        return jsonify({'status': 'error', 'message': 'target file not suitable for in-file update'}), 400
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/footer', methods=['GET'])
def get_footer():
    # read footer from template_files/footer.json if exists (or blob storage)
    try:
        if blob_mode_enabled():
            user_id = _get_request_user_id() or 'anonymous'
            user_token = _sanitize_token(user_id)
            txt = _blob_get_text_for_user(user_token, 'footer.json')
            if txt:
                try:
                    obj = json.loads(txt)
                except Exception:
                    return jsonify({'status': 'ok', 'footer': txt})
                if isinstance(obj, dict):
                    footer = obj.get('template') if 'template' in obj else obj.get('footer', '')
                    additional = obj.get('additional_footer', '')
                    fcs = obj.get('fcs_footer', '')
                    return jsonify({'status': 'ok', 'footer': footer or '', 'additional_footer': additional or '', 'fcs_footer': fcs or ''})
            # try shared container fallback
            shared_container = _get_shared_container()
            txt = _blob_get_text(shared_container, 'footer.json')
            if txt:
                try:
                    obj = json.loads(txt)
                except Exception:
                    return jsonify({'status': 'ok', 'footer': txt})
                if isinstance(obj, dict):
                    footer = obj.get('template') if 'template' in obj else obj.get('footer', '')
                    additional = obj.get('additional_footer', '')
                    fcs = obj.get('fcs_footer', '')
                    return jsonify({'status': 'ok', 'footer': footer or '', 'additional_footer': additional or '', 'fcs_footer': fcs or ''})
            return jsonify({'status': 'ok', 'footer': ''})

        # filesystem fallback
        ensure_templates_dir()
        path = os.path.join(get_template_dir(), 'footer.json')
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8-sig') as f:
                obj = json.load(f)
                if isinstance(obj, dict):
                    footer = obj.get('template') if 'template' in obj else obj.get('footer', '')
                    additional = obj.get('additional_footer', '')
                    fcs = obj.get('fcs_footer', '')
                    return jsonify({'status': 'ok', 'footer': footer or '', 'additional_footer': additional or '', 'fcs_footer': fcs or ''})
    except Exception:
        pass
    return jsonify({'status': 'ok', 'footer': ''})


@app.route('/footer', methods=['POST'])
def save_footer():
    payload = request.json or {}
    footer = payload.get('footer', '')
    additional = payload.get('additional_footer', '')
    fcs = payload.get('fcs_footer', '')
    if not isinstance(footer, str) or not isinstance(additional, str) or not isinstance(fcs, str):
        return jsonify({'status': 'error', 'message': 'footer fields must be strings'}), 400
    try:
        if blob_mode_enabled():
            user_id = _get_request_user_id() or 'anonymous'
            user_token = _sanitize_token(user_id)
            payload_txt = json.dumps({'footer': footer, 'additional_footer': additional, 'fcs_footer': fcs}, ensure_ascii=False)
            if not _blob_upload_text_for_user(user_token, 'footer.json', payload_txt):
                return jsonify({'status': 'error', 'message': 'failed to save footer to blob storage'}), 500
        else:
            ensure_templates_dir()
            path = os.path.join(get_template_dir(), 'footer.json')
            atomic_write_json(path, {'footer': footer, 'additional_footer': additional, 'fcs_footer': fcs})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'failed to save footer: {e}'}), 500
    return jsonify({'status': 'ok'})


@app.route('/templates/upload', methods=['POST'])
def upload_templates_zip():
    """Upload a ZIP file containing .json template files and extract them into the template directory.
    POST multipart/form-data with field 'file'. Only .json files are extracted; path traversal is prevented.
    """
    try:
        files = list(request.files.getlist('file'))
        if not files:
            return jsonify({'status': 'error', 'message': 'file field required'}), 400

        saved = []
        ensure_templates_dir()

        dbg(f"upload_templates_zip: blob_mode={blob_mode_enabled()} AZURE_BLOB_TEMPLATES={AZURE_BLOB_TEMPLATES}")
        try:
            cfg = get_storage_config()
            dbg(f"upload_templates_zip: storage_config={json.dumps(cfg, ensure_ascii=False)}")
        except Exception:
            pass

        for f in files:
            filename = getattr(f, 'filename', None) or 'uploaded'
            fname = filename
            data = f.read()
            # detect zip by magic bytes or extension
            is_zip = False
            try:
                if data[:4] == b'PK\x03\x04' or fname.lower().endswith('.zip'):
                    is_zip = True
            except Exception:
                is_zip = False

            if is_zip:
                try:
                    z = zipfile.ZipFile(io.BytesIO(data))
                except Exception:
                    continue
                for member in z.namelist():
                    nm = os.path.normpath(member)
                    if nm.startswith('..') or os.path.isabs(nm):
                        continue
                    if nm.endswith('/') or nm.endswith('\\'):
                        continue
                    if not nm.lower().endswith('.json'):
                        continue
                    base = os.path.basename(nm)
                    try:
                        with z.open(member) as src:
                            txt = src.read().decode('utf-8')
                        if blob_mode_enabled():
                            user_id = _get_request_user_id() or 'anonymous'
                            user_token = _sanitize_token(user_id)
                            _blob_upload_text_for_user(user_token, base, txt)
                        else:
                            dest = os.path.join(get_template_dir(), base)
                            try:
                                parsed = json.loads(txt)
                                atomic_write_json(dest, parsed)
                            except Exception:
                                atomic_write_json(dest, {'template': txt})
                        saved.append(base)
                    except Exception:
                        continue
            else:
                # single file -> store
                try:
                    txt = data.decode('utf-8')
                except Exception:
                    try:
                        txt = data.decode('utf-8', errors='ignore')
                    except Exception:
                        continue
                if not fname.lower().endswith('.json'):
                    fname = fname + '.json'
                if blob_mode_enabled():
                    user_id = _get_request_user_id() or 'anonymous'
                    user_token = _sanitize_token(user_id)
                    _blob_upload_text_for_user(user_token, os.path.basename(fname), txt)
                    saved.append(os.path.basename(fname))
                else:
                    dest = os.path.join(get_template_dir(), os.path.basename(fname))
                    try:
                        parsed = json.loads(txt)
                        atomic_write_json(dest, parsed)
                    except Exception:
                        atomic_write_json(dest, {'template': txt})
                    saved.append(os.path.basename(fname))

        return jsonify({'status': 'ok', 'saved': saved})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/sr/import', methods=['POST'])
def sr_import():
    payload = request.json or {}
    items = payload.get('items')
    if not isinstance(items, list):
        return jsonify({'status': 'error', 'message': 'items must be a list'}), 400
    try:
        # Validate all items first: require case_number, customer_title, contact
        for idx, it in enumerate(items, start=1):
            if not isinstance(it, dict):
                return jsonify({'status': 'error', 'message': f'item at index {idx} is not an object'}), 400
            if not it.get('case_number'):
                return jsonify({'status': 'error', 'message': f'item at index {idx} missing case_number'}), 400
            if not it.get('customer_title'):
                return jsonify({'status': 'error', 'message': f'item at index {idx} missing customer_title'}), 400
            if not it.get('contact'):
                return jsonify({'status': 'error', 'message': f'item at index {idx} missing contact'}), 400
        ensure_templates_dir()
        user_id = _get_request_user_id() or 'anonymous'
        user_token = _sanitize_token(user_id)
        sr_dir = os.path.join(get_template_dir(), 'sr')
        os.makedirs(sr_dir, exist_ok=True)
        # Build set of incoming filenames, save only new ones, and remove orphaned files
        incoming_files = set()
        saved = 0
        for it in items:
            if not isinstance(it, dict):
                continue
            case_no = it.get('case_number') or it.get('case') or it.get('caseNo')
            if not case_no:
                continue
            safe_fn = ''.join([c for c in str(case_no) if c.isalnum() or c in '-_'])
            fname = f'case_{safe_fn}.json'
            incoming_files.add(fname)
            # For blob mode, write as blob_files/{user_token}/sr/{fname}
            if blob_mode_enabled():
                try:
                    # preserve existing replace_name/sympton if present in blob
                    existing_txt = _blob_get_text_for_user(user_token, f'sr/{fname}')
                    if existing_txt:
                        try:
                            existing = json.loads(existing_txt)
                        except Exception:
                            existing = {}
                        if isinstance(existing, dict) and existing.get('replace_name'):
                            it['replace_name'] = existing.get('replace_name')
                        if isinstance(existing, dict) and existing.get('sympton'):
                            it['sympton'] = existing.get('sympton')
                    _blob_upload_text_for_user(user_token, f'sr/{fname}', json.dumps(it, ensure_ascii=False))
                    saved += 1
                except Exception:
                    continue
            else:
                path = os.path.join(sr_dir, fname)
                try:
                    # If existing file has replace_name, preserve it
                    if os.path.exists(path):
                        try:
                            with open(path, 'r', encoding='utf-8-sig') as f:
                                existing = json.load(f)
                        except Exception:
                            existing = {}
                        if isinstance(existing, dict) and existing.get('replace_name'):
                            it['replace_name'] = existing.get('replace_name')
                        # Preserve existing sympton (if present) so import does not wipe it
                        if isinstance(existing, dict) and existing.get('sympton'):
                            it['sympton'] = existing.get('sympton')
                    atomic_write_json(path, it)
                    saved += 1
                except Exception:
                    continue

        # Delete any files in sr_dir / blob that are not present in the incoming list
        deleted = 0
        if blob_mode_enabled():
            # list blobs under template_files/{user_token}/sr/
            try:
                blobs = _blob_list_templates_for_user(user_token)
                # _blob_list_templates_for_user returns names without sr/ prefix for top-level template files,
                # but our sr entries are stored under 'sr/<fname>' so we'll list via container directly.
                svc = get_blob_service_client()
                if svc:
                    container = svc.get_container_client(_get_shared_container())
                    prefix = f'template_files/{user_token}/sr/'
                    try:
                        for b in container.list_blobs(name_starts_with=prefix):
                            name = b.name[len(prefix):] if b.name.startswith(prefix) else b.name
                            if not name:
                                continue
                            if name not in incoming_files:
                                try:
                                    _blob_delete_for_user(user_token, f'sr/{name}')
                                    deleted += 1
                                except Exception:
                                    pass
                    except Exception:
                        pass
            except Exception:
                pass
        else:
            try:
                for fn in os.listdir(sr_dir):
                    if not fn.lower().endswith('.json'):
                        continue
                    if fn not in incoming_files:
                        try:
                            os.remove(os.path.join(sr_dir, fn))
                            deleted += 1
                        except Exception:
                            continue
            except Exception:
                pass
        return jsonify({'status': 'ok', 'saved': saved, 'deleted': deleted})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/sr/list', methods=['GET'])
def sr_list():
    try:
        ensure_templates_dir()
        user_id = _get_request_user_id() or 'anonymous'
        user_token = _sanitize_token(user_id)
        items = []
        if blob_mode_enabled():
            try:
                svc = get_blob_service_client()
                if svc:
                    container = svc.get_container_client(_get_shared_container())
                    prefix = f'template_files/{user_token}/sr/'
                    for b in container.list_blobs(name_starts_with=prefix):
                        name = b.name[len(prefix):] if b.name.startswith(prefix) else b.name
                        if not name or not name.lower().endswith('.json'):
                            continue
                        try:
                            txt = _blob_get_text_for_user(user_token, f'sr/{name}')
                            if not txt:
                                continue
                            obj = json.loads(txt)
                            if isinstance(obj, dict):
                                items.append({
                                    'case_number': obj.get('case_number') or obj.get('case') or obj.get('caseNo'),
                                    'customer_title': obj.get('customer_title') or obj.get('customer') or '',
                                    'internal_title': obj.get('internal_title') or obj.get('internal_title') or '',
                                    'contact': obj.get('contact') or '',
                                    'replace_name': obj.get('replace_name') or '',
                                    'sympton': obj.get('sympton') or ''
                                })
                        except Exception:
                            continue
            except Exception:
                pass
        else:
            sr_dir = os.path.join(get_template_dir(), 'sr')
            if os.path.exists(sr_dir):
                for fn in os.listdir(sr_dir):
                    if not fn.lower().endswith('.json'):
                        continue
                    path = os.path.join(sr_dir, fn)
                    try:
                        with open(path, 'r', encoding='utf-8-sig') as f:
                            obj = json.load(f)
                            if isinstance(obj, dict):
                                items.append({
                                    'case_number': obj.get('case_number') or obj.get('case') or obj.get('caseNo'),
                                    'customer_title': obj.get('customer_title') or obj.get('customer') or '',
                                    'internal_title': obj.get('internal_title') or obj.get('internal_title') or '',
                                    'contact': obj.get('contact') or '',
                                    'replace_name': obj.get('replace_name') or '',
                                    'sympton': obj.get('sympton') or ''
                                })
                    except Exception:
                        continue
        return jsonify({'status': 'ok', 'items': items})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/sr/replace', methods=['POST'])
def sr_replace():
    payload = request.json or {}
    case_no = payload.get('case_number')
    action = payload.get('action')  # 'set' or 'reset'
    replace_name = payload.get('replace_name')
    sympton = payload.get('sympton')
    if not case_no:
        return jsonify({'status': 'error', 'message': 'case_number required'}), 400
    try:
        ensure_templates_dir()
        sr_dir = os.path.join(get_template_dir(), 'sr')
        safe_fn = ''.join([c for c in str(case_no) if c.isalnum() or c in '-_'])
        fname = f'case_{safe_fn}.json'
        path = os.path.join(sr_dir, fname)
        user_id = _get_request_user_id() or 'anonymous'
        user_token = _sanitize_token(user_id)
        if blob_mode_enabled():
            # read existing from blob
            existing_txt = _blob_get_text_for_user(user_token, f'sr/{fname}')
            if not existing_txt:
                return jsonify({'status': 'error', 'message': 'case file not found'}), 404
            try:
                obj = json.loads(existing_txt)
            except Exception:
                obj = {}
            if action == 'set':
                if not isinstance(replace_name, str):
                    return jsonify({'status': 'error', 'message': 'replace_name must be string for set action'}), 400
                obj['replace_name'] = replace_name
                if isinstance(sympton, str):
                    obj['sympton'] = sympton
            elif action == 'reset':
                if 'replace_name' in obj:
                    obj.pop('replace_name', None)
                if 'sympton' in obj:
                    obj.pop('sympton', None)
            else:
                return jsonify({'status': 'error', 'message': "invalid action, use 'set' or 'reset'"}), 400
            _blob_upload_text_for_user(user_token, f'sr/{fname}', json.dumps(obj, ensure_ascii=False))
        else:
            if not os.path.exists(path):
                return jsonify({'status': 'error', 'message': 'case file not found'}), 404
            with open(path, 'r', encoding='utf-8-sig') as f:
                obj = json.load(f)
            if action == 'set':
                if not isinstance(replace_name, str):
                    return jsonify({'status': 'error', 'message': 'replace_name must be string for set action'}), 400
                obj['replace_name'] = replace_name
                # if sympton provided as string, store it as well
                if isinstance(sympton, str):
                    obj['sympton'] = sympton
            elif action == 'reset':
                if 'replace_name' in obj:
                    obj.pop('replace_name', None)
                # reset sympton as well when resetting replace_name
                if 'sympton' in obj:
                    obj.pop('sympton', None)
            else:
                return jsonify({'status': 'error', 'message': "invalid action, use 'set' or 'reset'"}), 400
            atomic_write_json(path, obj)
        return jsonify({'status': 'ok'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/owner', methods=['GET'])
def get_owner():
    try:
        ensure_templates_dir()
        try:
            if blob_mode_enabled():
                user_id = _get_request_user_id() or 'anonymous'
                user_token = _sanitize_token(user_id)
                txt = _blob_get_text_for_user(user_token, 'owner.json')
                if txt:
                    obj = json.loads(txt)
                    return jsonify({'status': 'ok', 'owner': obj})
            else:
                path = os.path.join(get_template_dir(), 'owner.json')
                if os.path.exists(path):
                    with open(path, 'r', encoding='utf-8-sig') as f:
                        obj = json.load(f)
                        return jsonify({'status': 'ok', 'owner': obj})
        except Exception:
            pass
    except Exception:
        pass
    return jsonify({'status': 'ok', 'owner': {}})


@app.route('/owner', methods=['POST'])
def save_owner():
    payload = request.json or {}
    owner = payload.get('owner')
    if not isinstance(owner, dict):
        return jsonify({'status': 'error', 'message': 'owner must be an object'}), 400
    try:
        ensure_templates_dir()
        user_id = _get_request_user_id() or 'anonymous'
        user_token = _sanitize_token(user_id)
        if blob_mode_enabled():
            _blob_upload_text_for_user(user_token, 'owner.json', json.dumps(owner, ensure_ascii=False))
        else:
            path = os.path.join(get_template_dir(), 'owner.json')
            atomic_write_json(path, owner)
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'failed to save owner: {e}'}), 500
    return jsonify({'status': 'ok'})


@app.route('/config', methods=['GET'])
def get_config():
    try:
        cfg = {}
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                    cfg = json.load(f) or {}
            except Exception:
                cfg = {}
        # always include resolved template_dir
        cfg_out = dict(cfg)
        try:
            cfg_out['template_dir'] = get_template_dir()
        except Exception:
            cfg_out['template_dir'] = ''
        return jsonify({'status': 'ok', 'config': cfg_out})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/config', methods=['POST'])
def save_config():
    payload = request.json or {}
    # allow setting 'template_dir' only
    td = payload.get('template_dir')
    if td is None:
        return jsonify({'status': 'error', 'message': 'template_dir required'}), 400
    try:
        # If td looks like a Windows absolute path but we are running on non-Windows,
        # preserve the raw value in config.json without resolving or trying to create it.
        try:
            import re
            is_win_abs = bool(re.match(r'^[A-Za-z]:[\\/]|^\\\\', td))
        except Exception:
            is_win_abs = False

        if is_win_abs and platform.system() != 'Windows':
            td_resolved = td
            try:
                with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                    json.dump({'template_dir': td_resolved}, f, ensure_ascii=False, indent=2)
            except Exception:
                return jsonify({'status': 'error', 'message': 'failed to write config'}), 500
            return jsonify({'status': 'ok', 'template_dir': td_resolved})

        # For non-Windows paths (or when running on Windows), resolve and ensure directory exists
        td_resolved = os.path.abspath(os.path.expanduser(td))
        try:
            os.makedirs(td_resolved, exist_ok=True)
        except Exception:
            pass
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump({'template_dir': td_resolved}, f, ensure_ascii=False, indent=2)
        return jsonify({'status': 'ok', 'template_dir': td_resolved})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/storage_config', methods=['GET'])
def api_get_storage_config():
    try:
        cfg = get_storage_config()
        return jsonify({'status': 'ok', 'config': cfg})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/storage_config', methods=['POST'])
def api_set_storage_config():
    payload = request.json or {}
    mode = payload.get('mode')
    storage_account = payload.get('storage_account')
    container = payload.get('container')
    if mode not in ('local', 'blob'):
        return jsonify({'status': 'error', 'message': "mode must be 'local' or 'blob'"}), 400
    ok = save_storage_config(mode, storage_account, container)
    if not ok:
        return jsonify({'status': 'error', 'message': 'failed to save storage config'}), 500
    return jsonify({'status': 'ok'})


@app.route('/phone_statuses', methods=['GET'])
def phone_statuses():
    try:
        ensure_templates_dir()
        # try blob first when enabled
        try:
            if blob_mode_enabled():
                user_id = _get_request_user_id() or 'anonymous'
                user_token = _sanitize_token(user_id)
                txt = _blob_get_text_for_user(user_token, 'phone.json')
                if not txt:
                    return jsonify({'items': []})
                j = json.loads(txt)
            else:
                path = os.path.join(get_template_dir(), 'phone.json')
                if not os.path.exists(path):
                    return jsonify({'items': []})
                with open(path, 'r', encoding='utf-8') as f:
                    j = json.load(f)
        except Exception:
            return jsonify({'items': []})
        items = []
        if isinstance(j, dict):
            for k, v in j.items():
                items.append({'label': k, 'value': v})
        elif isinstance(j, list):
            for entry in j:
                if isinstance(entry, dict) and 'label' in entry:
                    # Preserve label and value independently. Do not fall back to label when value is empty.
                    val = entry.get('value') if 'value' in entry else ''
                    items.append({'label': entry.get('label'), 'value': val})
        return jsonify({'items': items})
    except Exception as e:
        return jsonify({'items': [], 'error': str(e)}), 500


@app.route('/nextc', methods=['GET'])
def nextc_options():
    try:
        ensure_templates_dir()
        try:
            if blob_mode_enabled():
                user_id = _get_request_user_id() or 'anonymous'
                user_token = _sanitize_token(user_id)
                txt = _blob_get_text_for_user(user_token, 'nextc.json')
                if not txt:
                    # fallback to holidays.json in blob
                    txt = _blob_get_text_for_user(user_token, 'holidays.json')
                if not txt:
                    return jsonify({'items': []})
                j = json.loads(txt)
            else:
                path = os.path.join(get_template_dir(), 'nextc.json')
                if not os.path.exists(path):
                    # fallback to holidays.json if present
                    path2 = os.path.join(get_template_dir(), 'holidays.json')
                    if os.path.exists(path2):
                        path = path2
                    else:
                        return jsonify({'items': []})
                with open(path, 'r', encoding='utf-8') as f:
                    j = json.load(f)
        except Exception:
            return jsonify({'items': []})
        items = []
        if isinstance(j, dict):
            for k, v in j.items():
                items.append({'label': k, 'value': v})
        elif isinstance(j, list):
            for entry in j:
                if isinstance(entry, dict) and 'label' in entry:
                    # Preserve label and value independently. Do not fall back to label when value is empty.
                    val = entry.get('value') if 'value' in entry else ''
                    items.append({'label': entry.get('label'), 'value': val})
        return jsonify({'items': items})
    except Exception as e:
        return jsonify({'items': [], 'error': str(e)}), 500


@app.route('/holidays', methods=['GET'])
def holidays_list():
    try:
        ensure_templates_dir()
        try:
            if blob_mode_enabled():
                user_id = _get_request_user_id() or 'anonymous'
                user_token = _sanitize_token(user_id)
                txt = _blob_get_text_for_user(user_token, 'holidays.json')
                if not txt:
                    return jsonify({'dates': []})
                j = json.loads(txt)
            else:
                path = os.path.join(get_template_dir(), 'holidays.json')
                if not os.path.exists(path):
                    return jsonify({'dates': []})
                with open(path, 'r', encoding='utf-8') as f:
                    j = json.load(f)
        except Exception:
            return jsonify({'dates': []})
        dates = []
        if isinstance(j, dict):
            # keys are ISO dates
            for k in j.keys():
                dates.append(k)
        elif isinstance(j, list):
            for entry in j:
                if isinstance(entry, str):
                    dates.append(entry)
        return jsonify({'dates': dates})
    except Exception as e:
        return jsonify({'dates': [], 'error': str(e)}), 500


def compute_business_day_py(start_date: date, offset: int, holidays_set: set) -> date:
    dt = date(start_date.year, start_date.month, start_date.day)
    if offset <= 0:
        return dt
    remaining = offset
    while remaining > 0:
        dt = dt + timedelta(days=1)
        iso = dt.isoformat()
        # Python weekday: Monday=0 .. Sunday=6
        is_weekend = dt.weekday() in (5, 6)  # Saturday, Sunday
        if (not is_weekend) and (iso not in holidays_set):
            remaining -= 1
    return dt


@app.route('/render_template', methods=['POST'])
def render_template_server():
    """Render template server-side using holidays.json for business-day tokens.
    POST JSON: { 'template': str, optional 'today': 'YYYY-MM-DD' }
    Returns: { 'status':'ok', 'rendered': str }
    """
    payload = request.json or {}
    template = payload.get('template')
    if not isinstance(template, str):
        return jsonify({'status': 'error', 'message': 'template must be string'}), 400
    today_str = payload.get('today')
    try:
        if today_str:
            today = datetime.strptime(today_str, '%Y-%m-%d').date()
        else:
            today = datetime.now().date()
    except Exception:
        today = datetime.now().date()

    # load holidays
    try:
        path = os.path.join(get_template_dir(), 'holidays.json')
        holidays = set()
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                j = json.load(f)
            if isinstance(j, dict):
                holidays = set(j.keys())
            elif isinstance(j, list):
                holidays = set([d for d in j if isinstance(d, str)])
    except Exception:
        holidays = set()

    out = template
    try:
        dbg(f"RENDER_TEMPLATE: original_template_len={len(template)}")
        dbg(f"RENDER_TEMPLATE: original_preview={repr(template)[:400]}")
    except Exception:
        pass
    # simple replacements
    out = out.replace('<YYYY>', str(today.year))
    out = re.sub(r'<\s*MM\s*>', str(today.month), out)
    out = re.sub(r'<\s*DD\s*>', str(today.day), out)
    # Python's date.weekday(): Monday=0 .. Sunday=6
    weekdays = ['月','火','水','木','金','土','日']
    out = re.sub(r'<\s*AA\s*>', weekdays[today.weekday()], out)
    try:
        dbg(f"RENDER_TEMPLATE: after_simple_len={len(out)}")
        dbg(f"RENDER_TEMPLATE: after_simple_preview={repr(out)[:400]}")
    except Exception:
        pass

    # server-side: try to fill <offer_meeting> from meeting.json if present (use first label as fallback)
    try:
        meeting_label = ''
        mj = None
        try:
            if blob_mode_enabled():
                user_id = _get_request_user_id() or 'anonymous'
                user_token = _sanitize_token(user_id)
                txt = _blob_get_text_for_user(user_token, 'meeting.json')
                if txt:
                    try:
                        mj = json.loads(txt)
                    except Exception:
                        mj = None
            # fallback to local file when blob missing or not enabled
            if mj is None:
                mpath = os.path.join(get_template_dir(), 'meeting.json')
                if os.path.exists(mpath):
                    try:
                        with open(mpath, 'r', encoding='utf-8') as mf:
                            mj = json.load(mf)
                    except Exception:
                        mj = None
        except Exception:
            mj = None

        if isinstance(mj, dict):
            # keys are labels
            for k in mj.keys():
                meeting_label = k
                break
        elif isinstance(mj, list) and mj:
            first = mj[0]
            if isinstance(first, dict) and 'label' in first:
                meeting_label = first.get('label') or ''

        if meeting_label and str(meeting_label).strip() != '無し':
            out = out.replace('<offer_meeting>', meeting_label)
        else:
            # replace any line containing the token with a paragraph separator to preserve spacing
            parts = []
            for ln in out.splitlines():
                if '<offer_meeting>' in ln:
                    # insert an explicit blank line separator marker
                    parts.append('')
                else:
                    parts.append(ln)
            out = '\n'.join(parts)
            # collapse repeated blank lines to at most one blank line (i.e., two newlines)
            out = re.sub(r'\n{3,}', '\n\n', out)
        try:
            dbg(f"RENDER_TEMPLATE: after_meeting_len={len(out)}")
            dbg(f"RENDER_TEMPLATE: after_meeting_preview={repr(out)[:400]}")
        except Exception:
            pass
    except Exception:
        pass

    # replace <MM+N>, <DD+N>, <AA+N> with business-day calculation
    def repl_plus(match):
        token = match.group(1)  # MM or DD or AA
        n = int(match.group(2))
        dt = compute_business_day_py(today, n, holidays)
        if token == 'MM':
            return str(dt.month)
        if token == 'DD':
            return str(dt.day)
        if token == 'AA':
            # dt.weekday(): Monday=0 .. Sunday=6
            return weekdays[dt.weekday()]
        return match.group(0)

    # handle variants with spaces and full-width plus
    out = re.sub(r'<\s*(MM|DD|AA)\s*[+＋]\s*(\d+)\s*>', repl_plus, out)
    try:
        dbg(f"RENDER_TEMPLATE: final_len={len(out)}")
        dbg(f"RENDER_TEMPLATE: final_preview={repr(out)[:400]}")
    except Exception:
        pass

    return jsonify({'status': 'ok', 'rendered': out})


@app.route('/config/pick', methods=['POST'])
def pick_config_dir():
    # Run pick_dir.py in subprocess to show native dialog on main thread
    try:
        # Determine whether a native GUI picker can be launched.
        # On Windows we assume a GUI is available when running locally.
        # On non-Windows hosts require DISPLAY/Wayland/XDG session indicators.
        has_gui = False
        try:
            if platform.system() == 'Windows':
                has_gui = True
            else:
                if os.environ.get('DISPLAY') or os.environ.get('WAYLAND_DISPLAY') or os.environ.get('XDG_SESSION_TYPE'):
                    has_gui = True
        except Exception:
            has_gui = False

        if not has_gui:
            try:
                td = get_template_dir()
            except Exception:
                td = ''
            return jsonify({
                'status': 'unavailable',
                'message': 'Directory picker is not available on this host (no GUI). Use /config POST or set environment variable CREATEMAIL_TEMPLATE_DIR to change the template directory.',
                'template_dir': td
            }), 200
        import shutil, sys
        py = sys.executable or (shutil.which('python') or shutil.which('python3'))
        if not py:
            return jsonify({'status': 'error', 'message': 'python executable not found'}), 500
        script = os.path.join(BASE_DIR, 'pick_dir.py')
        if not os.path.exists(script):
            return jsonify({'status': 'error', 'message': 'picker script not found'}), 500
        # pass current template_dir as initial directory if available
        # Allow client to supply initial directory in POST body
        current = None
        try:
            body = request.json or {}
            if isinstance(body, dict):
                init_val = body.get('initial')
                if isinstance(init_val, str) and init_val:
                    # ensure we pass a str to expanduser/abspath for static type checkers
                    current = os.path.abspath(os.path.expanduser(init_val))
            # If no initial provided, try config.json, otherwise fall back to get_template_dir()
            if not current:
                cur = None
                try:
                    if os.path.exists(CONFIG_PATH):
                        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                            cfg = json.load(f) or {}
                            cur = cfg.get('template_dir')
                except Exception:
                    cur = None
                if isinstance(cur, str) and cur:
                    try:
                        cur_abs = os.path.abspath(os.path.expanduser(cur))
                        os.makedirs(cur_abs, exist_ok=True)
                        current = cur_abs
                    except Exception:
                        current = None
                else:
                    # use the new default template dir
                    try:
                        current = get_template_dir()
                        os.makedirs(current, exist_ok=True)
                    except Exception:
                        current = None
        except Exception:
            current = None
        args = [py, script]
        if current:
            args.append(current)
        proc = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.DEVNULL)
        out, err = proc.communicate(timeout=60)
        if err and not out:
            return jsonify({'status': 'error', 'message': err.decode(errors='ignore')}), 500
        try:
            data = json.loads(out.decode('utf-8') if out else '{}')
        except Exception:
            return jsonify({'status': 'error', 'message': 'invalid output from picker'}), 500
        sel = data.get('selected') if isinstance(data, dict) else None
        if not sel:
            return jsonify({'status': 'cancelled', 'message': 'no selection'}), 200
        td_resolved = os.path.abspath(os.path.expanduser(sel))
        try:
            os.makedirs(td_resolved, exist_ok=True)
        except Exception:
            pass
        try:
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump({'template_dir': td_resolved}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            return jsonify({'status': 'error', 'message': f'failed to write config: {e}'}), 500
        return jsonify({'status': 'ok', 'template_dir': td_resolved})
    except subprocess.TimeoutExpired:
        return jsonify({'status': 'error', 'message': 'picker timed out'}), 500
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/meeting_options', methods=['GET'])
def meeting_options():
    try:
        ensure_templates_dir()
        try:
            dbg('meeting_options: blob_mode=' + str(blob_mode_enabled()))
            if blob_mode_enabled():
                user_id = _get_request_user_id() or 'anonymous'
                user_token = _sanitize_token(user_id)
                txt = _blob_get_text_for_user(user_token, 'meeting.json')
                if not txt:
                    return jsonify({'items': []})
                j = json.loads(txt)
            else:
                path = os.path.join(get_template_dir(), 'meeting.json')
                if not os.path.exists(path):
                    return jsonify({'items': []})
                with open(path, 'r', encoding='utf-8') as f:
                    j = json.load(f)
        except Exception as ee:
            dbg('meeting_options: error reading meeting.json: ' + str(ee))
            return jsonify({'items': []})

        items = []
        if isinstance(j, dict):
            for k, v in j.items():
                items.append({'label': k, 'value': v})
        elif isinstance(j, list):
            for entry in j:
                if isinstance(entry, dict) and 'label' in entry:
                    # Preserve label and value independently. Do not fall back to label when value is empty.
                    val = entry.get('value') if 'value' in entry else ''
                    items.append({'label': entry.get('label'), 'value': val})
        return jsonify({'items': items})
    except Exception as e:
        dbg('meeting_options: unexpected error: ' + str(e))
        return jsonify({'items': [], 'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
