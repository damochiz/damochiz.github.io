from flask import Flask, render_template, request, jsonify
import os
import platform
import json
import re
from datetime import datetime, timedelta, date
import traceback
import threading


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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')

def get_template_dir():
    # Priority: env var -> config.json -> default workspace/template_files
    env = os.environ.get('CREATEMAIL_TEMPLATE_DIR') or os.environ.get('CREATE_MAIL_TEMPLATE_DIR')
    if env:
        try:
            return os.path.abspath(os.path.expanduser(env))
        except Exception:
            pass
    # try config.json
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
                if isinstance(cfg, dict) and cfg.get('template_dir'):
                    return os.path.abspath(os.path.expanduser(cfg.get('template_dir')))
    except Exception:
        pass
    # fallback default
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
    # Prefer canonical 01_FCS.json when resolving FCS group files
    try:
        if fname.lower() == 'fcs.json':
            preferred = os.path.join(get_template_dir(), '01_FCS.json')
            if os.path.exists(preferred):
                return preferred
    except Exception:
        pass
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


@app.route('/create_mail', methods=['POST'])
def create_mail():
    data = request.json or {}
    # prefer the largePanel content if provided
    body = data.get('content') or data.get('preview') or ''
    subject = data.get('title') or data.get('emailType') or ''

    # Only attempt Outlook automation on Windows
    if platform.system() != 'Windows':
        return jsonify({'status': 'error', 'message': 'Outlook automation is only supported on Windows.'}), 400

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


@app.route('/templates', methods=['GET'])
def get_templates():
    # Return all templates by reading per-type files
    data = {}
    # read per-type files
    try:
        ensure_templates_dir()
        for fn in os.listdir(get_template_dir()):
            # skip non-json files
            if not fn.lower().endswith('.json'):
                continue
            path = os.path.join(get_template_dir(), fn)
            try:
                with open(path, 'r', encoding='utf-8-sig') as f:
                    obj = json.load(f)
                    # if file contains {"template": "..."} treat as single template
                    if isinstance(obj, dict) and 'template' in obj:
                        key = os.path.splitext(fn)[0]
                        data[key] = obj.get('template')
                    # if file contains a mapping of multiple named templates, merge them
                    elif isinstance(obj, dict):
                        for k, v in obj.items():
                            # only include string templates
                            if isinstance(v, str):
                                data[k] = v
            except Exception:
                continue
    except Exception:
        pass
    return jsonify({'status': 'ok', 'templates': data})


@app.route('/templates/keys', methods=['GET'])
def get_template_keys():
    # Return mapping of filename -> list of keys inside that file (for multi-key files like fcs.json)
    result = {}
    try:
        ensure_templates_dir()
        for fn in os.listdir(get_template_dir()):
            if not fn.lower().endswith('.json'):
                continue
            path = os.path.join(get_template_dir(), fn)
            try:
                with open(path, 'r', encoding='utf-8-sig') as f:
                    obj = json.load(f)
                    if isinstance(obj, dict):
                        keys = list(obj.keys())
                        result[fn] = keys
            except Exception:
                continue
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
    # Save per-type files only
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
        # remove legacy single-key 'template' if we now have named keys (optional)
        if 'template' in existing and len(existing) > 1:
            existing.pop('template', None)

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
    # read footer from template_files/footer.json if exists
    try:
        ensure_templates_dir()
        path = os.path.join(get_template_dir(), 'footer.json')
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8-sig') as f:
                obj = json.load(f)
                if isinstance(obj, dict):
                    # support multiple keys: template (legacy), additional_footer, fcs_footer
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
        ensure_templates_dir()
        path = os.path.join(get_template_dir(), 'footer.json')
        atomic_write_json(path, {'footer': footer, 'additional_footer': additional, 'fcs_footer': fcs})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'failed to save footer: {e}'}), 500
    return jsonify({'status': 'ok'})


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

        # Delete any files in sr_dir that are not present in the incoming list
        deleted = 0
        for fn in os.listdir(sr_dir):
            if not fn.lower().endswith('.json'):
                continue
            if fn not in incoming_files:
                try:
                    os.remove(os.path.join(sr_dir, fn))
                    deleted += 1
                except Exception:
                    continue
        return jsonify({'status': 'ok', 'saved': saved, 'deleted': deleted})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/sr/list', methods=['GET'])
def sr_list():
    try:
        ensure_templates_dir()
        sr_dir = os.path.join(get_template_dir(), 'sr')
        items = []
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
        path = os.path.join(get_template_dir(), 'owner.json')
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8-sig') as f:
                obj = json.load(f)
                return jsonify({'status': 'ok', 'owner': obj})
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
        # ensure directory exists or can be created
        td_resolved = os.path.abspath(os.path.expanduser(td))
        try:
            os.makedirs(td_resolved, exist_ok=True)
        except Exception:
            pass
        # write config
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump({'template_dir': td_resolved}, f, ensure_ascii=False, indent=2)
        return jsonify({'status': 'ok', 'template_dir': td_resolved})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/phone_statuses', methods=['GET'])
def phone_statuses():
    try:
        ensure_templates_dir()
        path = os.path.join(get_template_dir(), 'phone.json')
        if not os.path.exists(path):
            return jsonify({'items': []})
        with open(path, 'r', encoding='utf-8') as f:
            j = json.load(f)
        items = []
        if isinstance(j, dict):
            for k, v in j.items():
                items.append({'label': k, 'value': v})
        elif isinstance(j, list):
            for entry in j:
                if isinstance(entry, dict) and 'label' in entry:
                    items.append({'label': entry.get('label'), 'value': entry.get('value', '')})
        return jsonify({'items': items})
    except Exception as e:
        return jsonify({'items': [], 'error': str(e)}), 500


@app.route('/nextc', methods=['GET'])
def nextc_options():
    try:
        ensure_templates_dir()
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
        items = []
        if isinstance(j, dict):
            for k, v in j.items():
                items.append({'label': k, 'value': v})
        elif isinstance(j, list):
            for entry in j:
                if isinstance(entry, dict) and 'label' in entry:
                    items.append({'label': entry.get('label'), 'value': entry.get('value', '')})
        return jsonify({'items': items})
    except Exception as e:
        return jsonify({'items': [], 'error': str(e)}), 500


@app.route('/holidays', methods=['GET'])
def holidays_list():
    try:
        ensure_templates_dir()
        path = os.path.join(get_template_dir(), 'holidays.json')
        if not os.path.exists(path):
            return jsonify({'dates': []})
        with open(path, 'r', encoding='utf-8') as f:
            j = json.load(f)
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
        mpath = os.path.join(get_template_dir(), 'meeting.json')
        if os.path.exists(mpath):
            with open(mpath, 'r', encoding='utf-8') as mf:
                mj = json.load(mf)
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
        import subprocess, shutil, sys
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
            if isinstance(body, dict) and body.get('initial'):
                current = os.path.abspath(os.path.expanduser(body.get('initial')))
            else:
                if os.path.exists(CONFIG_PATH):
                    with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                        cfg = json.load(f) or {}
                        cur = cfg.get('template_dir')
                        if cur:
                            try:
                                # normalize and ensure exists so filedialog initialdir works
                                cur_abs = os.path.abspath(os.path.expanduser(cur))
                                os.makedirs(cur_abs, exist_ok=True)
                                current = cur_abs
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
        path = os.path.join(get_template_dir(), 'meeting.json')
        if not os.path.exists(path):
            return jsonify({'items': []})
        with open(path, 'r', encoding='utf-8') as f:
            j = json.load(f)
        items = []
        if isinstance(j, dict):
            for k, v in j.items():
                items.append({'label': k, 'value': v})
        elif isinstance(j, list):
            for entry in j:
                if isinstance(entry, dict) and 'label' in entry:
                    items.append({'label': entry.get('label'), 'value': entry.get('value', '')})
        return jsonify({'items': items})
    except Exception as e:
        return jsonify({'items': [], 'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
