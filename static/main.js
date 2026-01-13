async function postJson(url, data){
  const res = await fetch(url, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(data)});
  const ct = res.headers.get('content-type') || '';
  try{
    if (ct.indexOf('application/json') !== -1){
      const j = await res.json();
      return j;
    }
    // If server did not return JSON, fall back to text and return an error-shaped object
    const txt = await res.text();
    return { status: res.ok ? 'ok' : 'error', message: txt };
  }catch(e){
    // JSON parse error or other failure: return text body if possible
    try{
      const txt = await res.text();
      return { status: 'error', message: txt || String(e) };
    }catch(ee){
      return { status: 'error', message: String(e) };
    }
  }
}

// Global error handler to surface unexpected exceptions during initialization
window.addEventListener('error', function(ev){
  try{ console.error('Global error caught:', ev.error || ev.message, ev.filename + ':' + ev.lineno); }catch(e){}
});
window.addEventListener('unhandledrejection', function(ev){
  try{ console.error('Unhandled promise rejection:', ev.reason); }catch(e){}
});

console.info('main.js loaded, readyState=', document.readyState);

const _createMailBtn = document.getElementById('createMailBtn');
if (_createMailBtn){
  _createMailBtn.addEventListener('click', async function(){
    const panelText = document.getElementById('largePanel') ? document.getElementById('largePanel').textContent : '';
    const titleEl = document.getElementById('title');
    const payload = { content: panelText, title: titleEl ? titleEl.value : '' };
    try{
      const res = await postJson('/create_mail', payload);
      if (res && res.status === 'ok') {
        alert('Outlook に新しいメールを作成しました。');
      } else {
        alert('メール作成に失敗しました: ' + (res && res.message ? res.message : 'Unknown'));
      }
    }catch(e){ alert('メール作成に失敗しました: ' + e); }
  });
} else { console.warn('createMailBtn not found'); }

// Create Schedule button handler: requires schedDate and schedTime to be filled.
(function(){
  const csBtn = document.getElementById('createScheduleBtn');
  if (!csBtn) return;
  csBtn.addEventListener('click', async function(){
    try{
      const sdEl = document.getElementById('schedDate');
      const stEl = document.getElementById('schedTime');
      const sdVal = sdEl && sdEl.value ? sdEl.value.trim() : '';
      const stVal = stEl && stEl.value ? stEl.value.trim() : '';
      if (!sdVal || !stVal){ alert('日付と時刻を入力してください'); return; }

      // Load schedule template. Prefer server API /templates which enumerates template files.
      let tpl = null;
      try{
        const r = await fetch('/templates');
        if (r && r.ok){
          const j = await r.json();
          if (j && j.templates && typeof j.templates === 'object'){
            // Some template files (like schedule.json) are exposed as individual keys
            // e.g. { mtg_title: "...", mtg_bofy: "..." } in the flattened templates mapping.
            // If we see mtg_title or mtg_bofy present, assemble tpl from those keys.
            const t = j.templates;
            if ((typeof t.mtg_title === 'string' || typeof t.mtg_bofy === 'string' || typeof t.mtg_body === 'string')){
              tpl = {};
              if (typeof t.mtg_title === 'string') tpl.mtg_title = t.mtg_title;
              if (typeof t.mtg_bofy === 'string') tpl.mtg_bofy = t.mtg_bofy;
              else if (typeof t.mtg_body === 'string') tpl.mtg_bofy = t.mtg_body;
            } else {
              // Otherwise try common filename-based keys
              tpl = t['schedule'] || t['schedule.json'] || null;
            }
            // If tpl is a JSON string, try to parse it
            if (tpl && typeof tpl === 'string'){
              try{ tpl = JSON.parse(tpl); }catch(e){ /* leave as string */ }
            }
          }
        }
      }catch(e){ /* ignore */ }
      // As a last-resort, try fetching the file directly (some deployments may serve static files)
      if (!tpl){
        try{ const r2 = await fetch('/template_files/schedule.json'); if (r2 && r2.ok) tpl = await r2.json(); }catch(e){}
      }
      if (!tpl){ alert('schedule.json を取得できませんでした'); return; }

      const rawTitle = (typeof tpl.mtg_title === 'string') ? tpl.mtg_title : '';
      const rawBody = (typeof tpl.mtg_bofy === 'string') ? tpl.mtg_bofy : (typeof tpl.mtg_body === 'string' ? tpl.mtg_body : '');

      // Use existing synchronous placeholder replacer to substitute tokens
      let title = sanitizeBlankLines(applySyncPlaceholders(rawTitle));
      let body = sanitizeBlankLines(applySyncPlaceholders(rawBody));
      // Ensure any unreplaced angle-bracket tokens (e.g. <customer_name>, <sr_number>)
      // are removed so created appointments do not contain raw placeholders.
      try{
        // Remove only simple placeholder tokens like <customer_name>, <sr_number>, <MM+1>
        // but preserve angle-bracketed URLs such as < https://... > (they contain ':' or '/').
        const placeholderRe = /<\s*[A-Za-z0-9_]+(?:[+\-]\d+)?\s*>/g;
        title = title.replace(placeholderRe, '').trim();
        body = body.replace(placeholderRe, '').trim();
      }catch(e){ /* ignore */ }

      // Compute ISO datetimes (assume local date/time inputs). End = start + 30 minutes
      // Parse components to avoid timezone ambiguity
      const dateMatch = sdVal.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      const timeMatch = stVal.match(/^(\d{1,2}):(\d{2})$/);
      if (!dateMatch || !timeMatch){ alert('日付または時刻の形式が不正です'); return; }
      const yyyy = parseInt(dateMatch[1],10);
      const mm = parseInt(dateMatch[2],10) - 1;
      const dd = parseInt(dateMatch[3],10);
      const hh = parseInt(timeMatch[1],10);
      const mi = parseInt(timeMatch[2],10);
      const start = new Date(yyyy, mm, dd, hh, mi, 0, 0);
      if (!(start instanceof Date) || isNaN(start.getTime())){ alert('開始日時の解析に失敗しました'); return; }
      const end = new Date(start.getTime() + 30*60*1000);
      // Format as local naive ISO-like string (no timezone offset) so server/Outlook
      // interprets the datetime as local time rather than converting from UTC.
      const pad2 = (n)=>String(n).padStart(2,'0');
      const formatWithOffset = (d)=>{
        const base = `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}:00`;
        // getTimezoneOffset() returns minutes difference UTC - local
        const offMin = -d.getTimezoneOffset();
        const sign = offMin >= 0 ? '+' : '-';
        const absMin = Math.abs(offMin);
        const oh = String(Math.floor(absMin / 60)).padStart(2,'0');
        const om = String(absMin % 60).padStart(2,'0');
        return `${base}${sign}${oh}:${om}`;
      };
      const startIso = formatWithOffset(start);
      const endIso = formatWithOffset(end);

      const payload = { title: title, body: body, start_iso: startIso, end_iso: endIso };
      try{ console.debug('create_schedule payload ->', payload); }catch(e){}
      const res = await postJson('/create_schedule', payload);
      if (res && res.status === 'ok'){ alert('Outlook に予定を作成しました。'); }
      else {
        const msg = (res && (res.message || res.error)) ? (res.message || res.error) : JSON.stringify(res);
        alert('予定作成に失敗しました: ' + msg);
      }
    }catch(e){ alert('予定作成に失敗しました: ' + e); }
  });
})();
// Template persistence and modal logic
const STORAGE_KEY = 'mailTemplates_v1';
let templatesStore = {};
// when editing a template, remember the source file (if any) returned by server
let __editingSourceFile = null;
function loadTemplatesFromStorage(){
  try{
    const raw = localStorage.getItem(STORAGE_KEY);
    templatesStore = raw ? JSON.parse(raw) : {};
  }catch(e){ templatesStore = {}; }
}
function saveTemplatesToStorage(){
  try{ localStorage.setItem(STORAGE_KEY, JSON.stringify(templatesStore)); }catch(e){ }
}

loadTemplatesFromStorage();
// Try to load server-side templates and merge into local store
async function loadTemplatesFromServer(){
  try{
    const res = await fetch('/templates');
    if (res.ok){
      const j = await res.json();
      if (j && j.templates){
        templatesStore = Object.assign({}, templatesStore, j.templates);
        saveTemplatesToStorage();
      }
    }
  }catch(e){ /* ignore network errors */ }
}
loadTemplatesFromServer();

// Load footer from server and populate footer textarea
async function loadFooterFromServer(){
  try{
    const res = await fetch('/footer');
      if (res && res.ok){
      const j = await res.json();
      if (j){
        // cache footer JSON for use during preview replacements
        window.__footerJson = j;
        const f = document.getElementById('footer');
        const af = document.getElementById('additionalFooter');
        const ff = document.getElementById('fcsFooter');
        if (f && typeof j.footer === 'string') f.value = j.footer;
        if (af && typeof j.additional_footer === 'string') af.value = j.additional_footer;
        if (ff && typeof j.fcs_footer === 'string') ff.value = j.fcs_footer;
      }
    }
  }catch(e){ }
}
loadFooterFromServer();

// Cached holidays set to reduce fetch latency and ensure availability
let cachedHolidaysSet = null;
async function loadHolidaysToCache(){
  try{
    const r = await fetch('/holidays');
    if (!r.ok) { cachedHolidaysSet = new Set(); return cachedHolidaysSet; }
    const j = await r.json();
    const arr = (j && Array.isArray(j.dates)) ? j.dates : [];
    cachedHolidaysSet = new Set(arr.map(s=>s));
    try{ console.debug('loaded holidays to cache, count', cachedHolidaysSet.size); }catch(e){}
    return cachedHolidaysSet;
  }catch(e){ cachedHolidaysSet = new Set(); return cachedHolidaysSet; }
}
loadHolidaysToCache();

// Load owner (担当) from server
async function loadOwnerFromServer(){
  try{
    const res = await fetch('/owner');
    if (res.ok){
      const j = await res.json();
      if (j && j.owner){
        const o = j.owner;
        // cache owner JSON for preview replacements
        window.__owner = o;
        if (o.family_name) document.getElementById('supportFamily').value = o.family_name;
        if (o.first_name) document.getElementById('supportFirst').value = o.first_name;
        if (o.family_name_furi) document.getElementById('supportFamilyFuri').value = o.family_name_furi;
        if (o.first_name_furi) document.getElementById('supportFirstFuri').value = o.first_name_furi;
      }
    }
  }catch(e){ }
}
loadOwnerFromServer();

// Load template keys (for multi-key files like fcs.json) and populate emailType select
async function loadTemplateKeys(){
  try{
    const res = await fetch('/templates/keys');
    if (!res.ok) return;
    const j = await res.json();
    if (!j || !j.files) return;
    const fileSelect = document.getElementById('templateFileSelect');
    const select = document.getElementById('emailType');
    // Populate templateFileSelect with available files
    if (fileSelect){
      // clear
      while(fileSelect.firstChild) fileSelect.removeChild(fileSelect.firstChild);
      const defaultOpt = document.createElement('option'); defaultOpt.value = ''; defaultOpt.textContent = '[テンプレート]'; fileSelect.appendChild(defaultOpt);
      const files = Object.keys(j.files || {}).sort();
      for (const fn of files){
        const o = document.createElement('option'); o.value = fn; o.textContent = fn; fileSelect.appendChild(o);
      }
    }

    // Helper to populate emailType select for a given filename (or null to show defaults)
    const populateEmailTypeForFile = (fn)=>{
      if (!select) return; // nothing to populate
      // remember existing non-dynamic options (like empty prompt)
      const preserved = [];
      for (let i=0;i<select.options.length;i++){
        const opt = select.options[i];
        if (!opt.dataset.dynamic){ preserved.push({value: opt.value, text: opt.text}); }
      }
      // clear select
      while(select.firstChild) select.removeChild(select.firstChild);
      // re-add preserved
      for (const p of preserved){ const o = document.createElement('option'); o.value = p.value; o.textContent = p.text; select.appendChild(o); }

      if (!fn){
        // No file selected: populate the emailType select with keys from all template files.
        const allFiles = Object.keys(j.files || {}).sort();
        for (const af of allFiles){
          const keys = j.files[af] || [];
          if (!keys || !keys.length) continue;
          const og = document.createElement('optgroup'); og.label = `--- ${af} ---`;
          for (const k of keys){ const o = document.createElement('option'); o.value = k; o.textContent = k; o.dataset.source = af; o.dataset.dynamic = '1'; og.appendChild(o); }
          select.appendChild(og);
        }
        return;
      }

      const keys = j.files[fn] || [];
      if (!keys || !keys.length) return;
      // Populate as options (no optgroup) and mark source
      for (const k of keys){
        const o = document.createElement('option'); o.value = k; o.textContent = k; o.dataset.source = fn; o.dataset.dynamic = '1'; select.appendChild(o);
      }

      };

    // If a templateFileSelect exists, attach change handler to populate emailType
    if (fileSelect){
      fileSelect.onchange = function(){
        const fn = fileSelect.value || null;
        populateEmailTypeForFile(fn);
        // trigger change to render preview for new selection
        try{ select.dispatchEvent(new Event('change')); }catch(e){}
      };
    }

    // initial population: if fileSelect has a current selection, use it; otherwise use defaults
    const initSel = fileSelect ? (fileSelect.value || '') : '';
    populateEmailTypeForFile(initSel || null);
  }catch(e){ console.error('loadTemplateKeys failed', e); }
}
// Delay loading template keys and initial render until holidays are loaded
async function __initAfterHolidays(){
  try{
    await loadHolidaysToCache();
  }catch(e){}
  try{ await loadTemplateKeys(); }catch(e){}
  // load other dependent lists after holidays
  try{ await loadPhoneStatuses(); }catch(e){}
  try{ await loadMeetingOptions(); }catch(e){}
  try{ await loadNextcOptions(); }catch(e){}
  try{ await loadSrList(); }catch(e){}
  // trigger initial render for selected email type if available
  const emailSelect = document.getElementById('emailType');
  if (emailSelect){ try{ emailSelect.dispatchEvent(new Event('change')); }catch(e){} }
}
// Ensure DOM is ready before initializing lists and binding handlers
if (document.readyState === 'loading'){
  document.addEventListener('DOMContentLoaded', function(){
    try{
      console.info('DOMContentLoaded fired');
      console.info('elements:', {
        templateFileSelect: !!document.getElementById('templateFileSelect'),
        emailType: !!document.getElementById('emailType'),
        createMailBtn: !!document.getElementById('createMailBtn'),
        nextC: !!document.getElementById('nextC')
      });
    }catch(e){ console.error('DOMContentLoaded diagnostics failed', e); }
    __initAfterHolidays().then(()=>{ console.info('__initAfterHolidays completed'); }).catch(e=>{ console.error('__initAfterHolidays failed', e); });
  });
} else {
  console.info('document already ready');
  __initAfterHolidays().then(()=>{ console.info('__initAfterHolidays completed (sync path)'); }).catch(e=>{ console.error('__initAfterHolidays failed (sync path)', e); });
}

// Reload page button handler (⌘/Ctrl not modified, performs a full reload)
(function(){
  const reloadBtn = document.getElementById('reloadPageBtn');
  if (!reloadBtn) return;
  reloadBtn.addEventListener('click', function(){
    try{
      // perform a hard reload to ensure latest server-side state
      location.reload(true);
    }catch(e){ location.reload(); }
  });
})();

// Refresh preview button handler (re-render right panel using current inputs)
{
  const refreshBtn = document.getElementById('refreshPreviewBtn');
  if (refreshBtn){
    refreshBtn.addEventListener('click', async function(){
      try{
        const emailSelect = document.getElementById('emailType');
        if (emailSelect && emailSelect.value){
          try{ emailSelect.dispatchEvent(new Event('change')); return; }catch(e){}
        }
        const panel = document.getElementById('largePanel');
        if (!panel) return;
        const tpl = panel.textContent || '';
        const server = await renderWithServer(tpl);
        if (server){ panel.textContent = sanitizeBlankLines(applySyncPlaceholders(server)); }
        else { panel.textContent = sanitizeBlankLines(applyTemplatePlaceholders(tpl)); }
      }catch(e){ console.error('refreshPreview failed', e); }
    });
  }
}

// Copy preview button: copy selected text if any, otherwise copy #largePanel content
(function(){
  const copyBtn = document.getElementById('copyPreviewBtn');
  if (!copyBtn) return;
  copyBtn.addEventListener('click', async function(){
    try{
      // If user has a selection in the document, prefer copying that
      const selection = window.getSelection();
      let textToCopy = '';
      if (selection && selection.toString().trim()){
        textToCopy = selection.toString();
      } else {
        const panel = document.getElementById('largePanel');
        textToCopy = panel ? panel.textContent || '' : '';
      }
      if (!textToCopy){ console.warn('コピーするテキストがありません'); return; }
      // Try navigator.clipboard first
      if (navigator.clipboard && navigator.clipboard.writeText){
        try{ await navigator.clipboard.writeText(textToCopy); console.info('クリップボードにコピーしました'); return; }catch(e){ /* fallthrough to fallback */ }
      }
      // Fallback: create textarea, select, execCommand
      const ta = document.createElement('textarea');
      ta.style.position = 'fixed'; ta.style.left = '-9999px'; ta.style.top = '0'; ta.setAttribute('aria-hidden','true');
      ta.value = textToCopy;
      document.body.appendChild(ta);
      ta.focus(); ta.select();
      const ok = document.execCommand && document.execCommand('copy');
      document.body.removeChild(ta);
      if (ok) console.info('クリップボードにコピーしました'); else console.warn('コピーに失敗しました');
    }catch(e){ console.error('copy failed', e); console.warn('コピーに失敗しました: ' + e); }
  });
})();

// Load phone statuses and populate phoneStatus select
async function loadPhoneStatuses(){
  try{
    const res = await fetch('/phone_statuses');
    if (!res.ok) return;
    const j = await res.json();
      if (!j || !Array.isArray(j.items)) return;
      // cache phone items for preview replacements
      window.__phoneItems = j.items;
      const select = document.getElementById('phoneStatus');
      if (!select) return;
      // clear existing options
      while(select.firstChild) select.removeChild(select.firstChild);
      // populate options from server items and select default when label is exactly '無し'
      const seen = new Set();
      let defaultIndex = -1;
      for (const item of j.items){
        if (!item || typeof item.label !== 'string') continue;
        const label = item.label.trim();
        const rawVal = (typeof item.value === 'string') ? item.value : (item.value != null ? String(item.value) : '');
        if (!label) continue;
        if (seen.has(label)) continue;
        const o = document.createElement('option');
        // set a usable non-empty option.value for browser selection while keeping server-provided raw value
        o.value = rawVal !== '' ? rawVal : label;
        o.textContent = label;
        o.dataset.rawValue = rawVal;
        select.appendChild(o);
        seen.add(label);
        if (label === '無し' && defaultIndex === -1){
          defaultIndex = select.options.length - 1;
        }
      }
      // If a '無し' entry was found, select it by default; otherwise select first option
      if (defaultIndex >= 0){ select.selectedIndex = defaultIndex; }
      else if (select.options.length) { select.selectedIndex = 0; }
  }catch(e){ console.error('loadPhoneStatuses failed', e); }
}
loadPhoneStatuses();

// Load meeting options from meeting.json and populate meetingStatus select
async function loadMeetingOptions(){
  try{
    const res = await fetch('/meeting_options');
    if (!res.ok) return;
    const j = await res.json();
    if (!j || !Array.isArray(j.items)) return;
    // cache meeting options for preview replacements
    window.__meetingItems = j.items;
    const select = document.getElementById('meetingStatus');
    if (!select) return;
    // clear existing options
    while(select.firstChild) select.removeChild(select.firstChild);
    // populate options from server items and select default when label is exactly '無し'
    const seen = new Set();
    let defaultIndex = -1;
    for (const item of j.items){
      if (!item || typeof item.label !== 'string') continue;
      const label = item.label.trim();
      const rawVal = (typeof item.value === 'string') ? item.value : (item.value != null ? String(item.value) : '');
      if (!label) continue;
      if (seen.has(label)) continue;
      const o = document.createElement('option');
      o.value = rawVal !== '' ? rawVal : label;
      o.textContent = label;
      o.dataset.rawValue = rawVal;
      select.appendChild(o);
      seen.add(label);
      if (label === '無し' && defaultIndex === -1){
        defaultIndex = select.options.length - 1;
      }
    }
    // If a '無し' entry was found, select it by default; otherwise select first option
    if (defaultIndex >= 0){ select.selectedIndex = defaultIndex; }
    else if (select.options.length) { select.selectedIndex = 0; }
  }catch(e){ console.error('loadMeetingOptions failed', e); }
}
loadMeetingOptions();

// Load NextC options from nextc.json (or holidays.json fallback) and populate nextC select
async function loadNextcOptions(){
  try{
    const res = await fetch('/nextc');
    if (!res.ok) return;
    const j = await res.json();
    if (!j || !Array.isArray(j.items)) return;
    // cache nextc items for preview replacements
    window.__nextcItems = j.items;
    const select = document.getElementById('nextC');
    if (!select) return;
    // clear existing options
    while(select.firstChild) select.removeChild(select.firstChild);
    let defaultIndex = -1;
    for (const item of j.items){
      if (!item || typeof item.label !== 'string') continue;
      const label = item.label.trim();
      const rawVal = (typeof item.value === 'string') ? item.value : (item.value != null ? String(item.value) : '');
      if (!label) continue;
      const o = document.createElement('option');
      o.value = rawVal !== '' ? rawVal : label;
      o.textContent = label;
      o.dataset.rawValue = rawVal;
      select.appendChild(o);
      if (label === '無し' && defaultIndex === -1){ defaultIndex = select.options.length - 1; }
    }
    // select '無し' by default if present, otherwise keep first option
    if (defaultIndex >= 0){ select.selectedIndex = defaultIndex; }
    // attach change handler to refresh preview when selection changes
    // remove previous handlers by replacing the element's onchange
    select.onchange = function(){
      try{
        const emailSelect = document.getElementById('emailType');
        if (emailSelect){
          // re-render template for current email type
          emailSelect.dispatchEvent(new Event('change'));
        } else {
          const panel = document.getElementById('largePanel');
          if (panel){
            (async ()=>{
              const tpl = panel.textContent || '';
              const server = await renderWithServer(tpl);
              if (server) { panel.textContent = sanitizeBlankLines(applySyncPlaceholders(server)); }
              else { panel.textContent = applyTemplatePlaceholders(tpl); }
            })();
          }
        }
      }catch(e){ console.error('nextC change handler failed', e); }
    };
  }catch(e){ console.error('loadNextcOptions failed', e); }
}
loadNextcOptions();

// When meeting selection changes, refresh the rendered preview on the right
(function(){
  const meetingSelect = document.getElementById('meetingStatus');
  if (!meetingSelect) return;
  meetingSelect.addEventListener('change', function(){
    try{
      const emailSelect = document.getElementById('emailType');
      if (emailSelect){
        // re-render template for current email type
        emailSelect.dispatchEvent(new Event('change'));
      } else {
        const panel = document.getElementById('largePanel');
        if (panel){
          (async ()=>{
            const tpl = panel.textContent || '';
            const server = await renderWithServer(tpl);
            if (server) { panel.textContent = sanitizeBlankLines(applySyncPlaceholders(server)); }
            else { panel.textContent = applyTemplatePlaceholders(tpl); }
          })();
        }
      }
    }catch(e){ console.error('refresh on meeting change failed', e); }
  });
})();

// When phone status changes, refresh the rendered preview on the right
(function(){
  const phoneSelect = document.getElementById('phoneStatus');
  if (!phoneSelect) return;
  phoneSelect.addEventListener('change', function(){
    try{
      const emailSelect = document.getElementById('emailType');
      if (emailSelect){
        // trigger the same rendering code as when emailType changes
        emailSelect.dispatchEvent(new Event('change'));
        } else {
        // fallback: reapply placeholders to current panel content, prefer server render
        const panel = document.getElementById('largePanel');
        if (panel){
          (async ()=>{
            const tpl = panel.textContent || '';
            const server = await renderWithServer(tpl);
            if (server) { panel.textContent = sanitizeBlankLines(applySyncPlaceholders(server)); }
            else { panel.textContent = applyTemplatePlaceholders(tpl); }
          })();
        }
      }
    }catch(e){ console.error('refresh on phone change failed', e); }
  });
})();

// Load SR list and populate SR select (if present)
async function loadSrList(){
  try{
    const input = document.getElementById('srNumberInput');
    const datalist = document.getElementById('srList');
    if (!input || !datalist) return;
    const res = await fetch('/sr/list');
    if (!res.ok) return;
    const j = await res.json();
    if (!j || !j.items) return;
    // clear existing datalist options
    while(datalist.firstChild) datalist.removeChild(datalist.firstChild);
    // Sort items by internal_title ascending (empty titles go last)
    const items = Array.isArray(j.items) ? j.items.slice().sort((a,b)=>{
      const aTitle = (a && a.internal_title) ? String(a.internal_title).toLowerCase() : null;
      const bTitle = (b && b.internal_title) ? String(b.internal_title).toLowerCase() : null;
      if (aTitle === bTitle) return 0;
      if (aTitle === null) return 1;
      if (bTitle === null) return -1;
      return aTitle < bTitle ? -1 : 1;
    }) : [];
    // Replace/Reset handlers are attached once at top-level to avoid duplicate listeners
    // mapping from case_number -> meta
    window.__srMap = {};
    // debug: show first 10 sorted titles
    try{ console.debug('SR sorted preview ->', items.slice(0,10).map(x=>({case_number:x.case_number, internal_title:x.internal_title}))); }catch(e){}
    for (const it of items){
      const v = it.case_number || '';
      if (!v) continue;
      const opt = document.createElement('option');
      // Show internal_title in the option label if available so the dropdown shows extra context
      opt.value = v;
      if (it.internal_title) opt.label = `${v} - ${it.internal_title}`;
      datalist.appendChild(opt);
      window.__srMap[v] = {contact: it.contact || '', customer_title: it.customer_title || '', internal_title: it.internal_title || '', replace_name: it.replace_name || '', sympton: it.sympton || ''};
    }
    // when user types or selects, attempt to autofill if exact match
    input.addEventListener('input', function(){
      const val = input.value && input.value.trim() ? input.value.trim() : '';
      const warningEl = document.getElementById('srWarning');
      const customerEl = document.getElementById('customerName');
      const titleEl = document.getElementById('title');
      const contentEl = document.getElementById('content');

      // If input is empty: clear customer and title and hide warning/info
      if (!val){
        try{ if (customerEl) customerEl.value = ''; }catch(e){}
        try{ if (titleEl) titleEl.value = ''; }catch(e){}
        try{ if (contentEl) contentEl.value = ''; }catch(e){}
        if (warningEl) { warningEl.style.display = 'none'; }
        try{ const infoEl = document.getElementById('srInfo'); if (infoEl) { infoEl.textContent = ''; infoEl.style.display = 'inline-block'; } }catch(e){}
        try{ const contactEl = document.getElementById('contactInfo'); if (contactEl) contactEl.textContent = ''; }catch(e){}
        // refresh preview to reflect cleared fields
        try{ const emailSelect = document.getElementById('emailType'); if (emailSelect && emailSelect.value) emailSelect.dispatchEvent(new Event('change')); else { const panel = document.getElementById('largePanel'); if (panel){ (async ()=>{ const tpl = panel.textContent || ''; const server = await renderWithServer(tpl); if (server){ panel.textContent = sanitizeBlankLines(applySyncPlaceholders(server)); } else { panel.textContent = sanitizeBlankLines(applyTemplatePlaceholders(tpl)); } })(); } } }catch(e){ console.error('refresh on sr clear failed', e); }
        return;
      }

      const meta = window.__srMap[val];
      // If SR not in the datalist / map: show warning and do not overwrite customer/title (to preserve user edits)
      if (!meta){
        // show warning in the info area by hiding info and showing the warning element
        if (warningEl){ warningEl.textContent = '未登録の SR 番号です'; warningEl.style.display = 'inline-block'; }
        try{ const infoEl = document.getElementById('srInfo'); if (infoEl) { infoEl.textContent = ''; infoEl.style.display = 'none'; } }catch(e){}
        // still clear content area since it's SR-specific
        try{ if (contentEl) contentEl.value = ''; }catch(e){}
        // refresh preview to show cleared content
        try{ const emailSelect = document.getElementById('emailType'); if (emailSelect && emailSelect.value) emailSelect.dispatchEvent(new Event('change')); else { const panel = document.getElementById('largePanel'); if (panel){ (async ()=>{ const tpl = panel.textContent || ''; const server = await renderWithServer(tpl); if (server){ panel.textContent = sanitizeBlankLines(applySyncPlaceholders(server)); } else { panel.textContent = sanitizeBlankLines(applyTemplatePlaceholders(tpl)); } })(); } } }catch(e){ console.error('refresh on sr invalid failed', e); }
        return;
      }

      // Valid SR: hide warning, populate fields and update srInfo
      if (warningEl) { warningEl.style.display = 'none'; }
      try{
        // Prefer replace_name when available; otherwise fall back to contact.
        if (meta.replace_name) customerEl.value = meta.replace_name;
        else if (meta.contact) customerEl.value = meta.contact;
        // Title still comes from customer_title
        if (meta.customer_title) titleEl.value = meta.customer_title;
        if (contentEl) contentEl.value = (meta.sympton || '');
      }catch(e){ /* ignore DOM errors */ }
      try{ const infoEl = document.getElementById('srInfo'); if (infoEl) { infoEl.textContent = (meta.internal_title || ''); infoEl.style.display = 'inline-block'; } }catch(e){}
      // update contactInfo display (next to Replace button)
      try{ const contactEl = document.getElementById('contactInfo'); if (contactEl) { contactEl.textContent = (meta.contact || ''); } }catch(e){}

      // Refresh preview on SR change so right panel reflects updated fields
      try{
        const emailSelect = document.getElementById('emailType');
        if (emailSelect && emailSelect.value){
          try{ emailSelect.dispatchEvent(new Event('change')); }
          catch(e){ /* ignore */ }
        } else {
          const panel = document.getElementById('largePanel');
          if (panel){
            (async ()=>{
              const tpl = panel.textContent || '';
              const server = await renderWithServer(tpl);
              if (server){ panel.textContent = sanitizeBlankLines(applySyncPlaceholders(server)); }
              else { panel.textContent = sanitizeBlankLines(applyTemplatePlaceholders(tpl)); }
            })();
          }
        }
      }catch(e){ console.error('refresh on sr input failed', e); }
    });
  }catch(e){ console.error('loadSrList failed', e); }
}
loadSrList();

// Refresh SR button clears SR and related fields
(function(){
  const refreshSrBtn = document.getElementById('refreshSrBtn');
  if (!refreshSrBtn) return;
  refreshSrBtn.addEventListener('click', function(){
    try{
      const srEl = document.getElementById('srNumberInput');
      if (srEl) srEl.value = '';
      const customerEl = document.getElementById('customerName'); if (customerEl) customerEl.value = '';
      const titleEl = document.getElementById('title'); if (titleEl) titleEl.value = '';
      const contentEl = document.getElementById('content'); if (contentEl) contentEl.value = '';
      const contactEl = document.getElementById('contactInfo'); if (contactEl) contactEl.textContent = '';
      const srInfo = document.getElementById('srInfo'); if (srInfo) srInfo.textContent = '';
      const srWarning = document.getElementById('srWarning'); if (srWarning) srWarning.style.display = 'none';
      // trigger preview refresh
      const emailSelect = document.getElementById('emailType');
      if (emailSelect && emailSelect.value) emailSelect.dispatchEvent(new Event('change'));
      else { const panel = document.getElementById('largePanel'); if (panel) panel.textContent = ''; }
    }catch(e){ console.error('refreshSrBtn click failed', e); }
  });
})();

// Replace/Reset handlers are defined further down; no top-level attach here.

const modal = document.getElementById('templateModal');
const editor = document.getElementById('templateEditor');
const templateChangeBtn = document.getElementById('templateChangeBtn');
const saveTemplateBtn = document.getElementById('saveTemplateBtn');
const cancelTemplateBtn = document.getElementById('cancelTemplateBtn');

// SR Import UI
const srImportBtn = document.getElementById('srImportBtn');
const srImportModal = document.getElementById('srImportModal');
const srImportTextarea = document.getElementById('srImportTextarea');
const srImportSaveBtn = document.getElementById('srImportSaveBtn');
const srImportCancelBtn = document.getElementById('srImportCancelBtn');

if (srImportBtn){
  srImportBtn.addEventListener('click', function(){ 
    if (!srImportModal) { console.warn('SR import modal not found'); return; }
    srImportModal.style.display = 'flex';
    srImportModal.setAttribute('aria-hidden','false');
    if (srImportTextarea) srImportTextarea.focus();
  });
}
if (srImportCancelBtn){ 
  srImportCancelBtn.addEventListener('click', function(){ 
    if (!srImportModal) return; 
    srImportModal.style.display='none'; 
    srImportModal.setAttribute('aria-hidden','true'); 
  }); 
}

if (srImportSaveBtn){
  srImportSaveBtn.addEventListener('click', async function(){
    const raw = srImportTextarea.value || '';
    if (!raw.trim()){ alert('入力が空です'); return; }
    // parse lines
    const lines = raw.split(/\r?\n/).map(l=>l.trim()).filter(l=>l.length);
    const items = [];
    for (const line of lines){
      // split by tab
      const cols = line.split('\t');
      // some inputs may have multiple spaces; fallback split
      if (cols.length < 4){
        // try splitting by 2+ spaces
        const sp = line.split(/\s{2,}/);
        if (sp.length >= 4) cols.splice(0, cols.length, ...sp);
      }
      const caseNo = (cols[0] || '').trim();
      const customerTitle = (cols[1] || '').trim();
      const internalTitle = (cols[2] || '').trim();
      const contact = (cols[3] || '').trim();
      if (!caseNo) continue;
      items.push({case_number: caseNo, customer_title: customerTitle, internal_title: internalTitle, contact: contact});
    }
    if (!items.length){ alert('有効な行が見つかりません'); return; }
    try{
      const res = await fetch('/sr/import', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({items})});
      const text = await res.text();
      let j = null;
      try{ j = text ? JSON.parse(text) : null; }catch(e){}
      if (!res.ok){
        const msg = (j && j.message) ? j.message : (`HTTP ${res.status}`);
        alert('取り込みに失敗しました: ' + msg);
        return;
      }
      if (j && j.status && j.status !== 'ok'){
        alert('取り込みに失敗しました: ' + (j.message || JSON.stringify(j)));
        return;
      }
      // success
      try{ await loadSrList(); }catch(e){ console.error('reload sr list failed', e); }
      alert('SRを取り込みました');
      if (srImportModal) { srImportModal.style.display='none'; srImportModal.setAttribute('aria-hidden','true'); }
      if (srImportTextarea) srImportTextarea.value='';
    }catch(e){ alert('取り込みに失敗しました: ' + e); }
  });
}

// DfM and ASC buttons: open target URLs with SR number in new tab
(function(){
  function openWithSr(baseUrlTemplate){
    try{
      const srEl = document.getElementById('srNumberInput');
      const sr = srEl && srEl.value ? srEl.value.toString().trim() : '';
      if (!sr){ console.warn('ボタン操作: SR番号が入力されていません'); return; }
      const encoded = encodeURIComponent(sr);
      const url = baseUrlTemplate.replace(/<sr_number>/g, encoded);
      window.open(url, '_blank');
    }catch(e){ console.error('openWithSr failed', e); }
  }
  const dfm = document.getElementById('dfmBtn');
  if (dfm){ dfm.addEventListener('click', function(){ openWithSr('https://onesupport.crm.dynamics.com/main.aspx?appid=101acb62-8d00-eb11-a813-000d3a8b3117&pagetype=search&searchText=<sr_number>'); }); }
  const asc = document.getElementById('ascBtn');
  if (asc){ asc.addEventListener('click', function(){ openWithSr('https://azuresupportcenter.azure.com/solutionexplorer?srId=<sr_number>'); }); }
})();

function openModalWithContent(content){
  editor.value = content;
  modal.style.display = 'flex';
  modal.setAttribute('aria-hidden','false');
  editor.focus();
}

function closeModal(){
  modal.style.display = 'none';
  modal.setAttribute('aria-hidden','true');
}

function getDefaultRawTemplateFor(type){
  if (type === 'FCS'){
    // Default template removed — rely on per-file templates under template_files/
    return '';
  }
  return '';
}

function getTodayDateString(){
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = today.getMonth()+1;
  const dd = today.getDate();
  return `${yyyy} 年 ${mm} 月 ${dd} 日`;
}

function applyTemplatePlaceholders(template){
  if (!template) return template || '';
  let out = template;
  // date
  out = out.replace(/__TODAY__/g, getTodayDateString());
  out = out.replace(/<YYYY>\s*年\s*<MM>\s*月\s*<DD>\s*日/g, getTodayDateString());
  // basic synchronous placeholder replacements (moved before async holidays replacement)
  try{
    const cname = document.getElementById('customerName') ? document.getElementById('customerName').value.trim() : '';
    const srElem = document.getElementById('srNumberInput');
    const sr = srElem ? (srElem.value ? srElem.value.trim() : '') : '';
    const titleVal = document.getElementById('title') ? document.getElementById('title').value.trim() : '';
    const contentVal = document.getElementById('content') ? document.getElementById('content').value.trim() : '';
    // Determine NextC label: prefer selected option's textContent when available,
    // otherwise prefer the select's value, and finally fall back to server-provided nextc items (look for label '無し').
    let nextc = '';
    try{
      const nextcElem = document.getElementById('nextC');
      const nextcItems = window.__nextcItems || [];
      if (nextcElem){
        if (nextcElem.selectedIndex >= 0){
          const opt = nextcElem.options[nextcElem.selectedIndex];
          const optVal = (opt.value || '').trim();
          const optText = (opt.textContent || '').trim();
          if (optVal){
            nextc = optVal;
          } else if (optText){
            // when option has no value, try to map the visible label to nextcItems
            // Use only the server-provided value when available. Do not fall back to visible label.
            const found = (Array.isArray(nextcItems) ? nextcItems.find(it => it && (it.label === optText || it.value === optText)) : null);
            if (found && typeof found.value === 'string' && found.value.trim() !== '') nextc = found.value;
            else nextc = '';
          }
        } else {
          nextc = (nextcElem.value || '').trim();
        }
      }
      if (!nextc && Array.isArray(nextcItems) && nextcItems.length){
        const found = nextcItems.find(it => it && it.label === '無し');
        if (found) nextc = (found && typeof found.value === 'string' && found.value.trim() !== '') ? found.value : '';
      }
    }catch(e){ nextc = ''; }

    if (cname){ out = out.replace(/<customer_name>/g, cname); }
    if (sr){ out = out.replace(/<sr_number>/g, sr); }
    if (contentVal){ out = out.replace(/<sr_body>/g, contentVal); out = out.replace(/<content>/g, contentVal); }
    // If user selected a schedule date/time, populate <month>, <day>, <weekday>, <time>
    try{
      const schedDateElem = document.getElementById('schedDate');
      const schedTimeElem = document.getElementById('schedTime');
      const sd = schedDateElem && schedDateElem.value ? new Date(schedDateElem.value) : null;
      const st = schedTimeElem && schedTimeElem.value ? schedTimeElem.value : '';
      if (sd instanceof Date && !isNaN(sd.getTime())){
        const months = sd.getMonth() + 1;
        const days = sd.getDate();
        const weekdays = ['日','月','火','水','木','金','土'];
        const w = weekdays[sd.getDay()];
        out = out.replace(/<month>/g, String(months));
        out = out.replace(/<day>/g, String(days));
        out = out.replace(/<weekday>/g, String(w));
        // time handling: support <time> and <time+N> (minutes offset). Format HH:MM
        const formatTimeWithOffset = (base, offsetMinutes)=>{
          if (!base) return '';
          const m = (''+base).toString().match(/^(\d{1,2}):(\d{2})$/);
          if (!m) return base;
          let hh = parseInt(m[1],10); let mm = parseInt(m[2],10);
          let total = hh*60 + mm + (offsetMinutes || 0);
          const dayMinutes = 24*60;
          total = ((total % dayMinutes) + dayMinutes) % dayMinutes;
          const rh = String(Math.floor(total/60)).padStart(2,'0');
          const rm = String(total%60).padStart(2,'0');
          return `${rh}:${rm}`;
        };
        if (st){
          // replace <time+N> and <time-N> (allow spaces and full-width signs)
          out = out.replace(/<\s*time\s*([+＋\-－])\s*(\d+)\s*>/gi, function(_, sign, mins){
            const n = parseInt(mins || '0', 10) || 0;
            const offset = (sign === '+' || sign === '＋') ? n : -n;
            return formatTimeWithOffset(st, offset);
          });
          out = out.replace(/<\s*time\s*>/gi, function(){ return formatTimeWithOffset(st, 0); });
        }
      }
    }catch(e){ /* ignore */ }
    if (nextc){ out = out.replace(/<nextc>/g, nextc); out = out.replace(/<nextC>/g, nextc); }

    // phone status placeholder handling: prefer server-provided raw value (dataset.rawValue).
    // If the raw value is empty (explicit ''), treat as empty and remove lines containing the token.
    try{
      const phoneSelect = document.getElementById('phoneStatus');
      let phoneVal = '';
      if (phoneSelect){
        const selIdx = phoneSelect.selectedIndex >= 0 ? phoneSelect.selectedIndex : 0;
        const opt = phoneSelect.options && phoneSelect.options[selIdx] ? phoneSelect.options[selIdx] : null;
        if (opt && typeof opt.dataset.rawValue === 'string'){
          phoneVal = opt.dataset.rawValue.trim();
        } else if (opt){
          phoneVal = (opt.value || '').toString().trim();
        }
      }
      const norm = (phoneVal || '').toString();
      if (norm){ out = out.replace(/<phone_status>/g, phoneVal); }
      else { out = out.replace(/<phone_status>/g, ''); }
    }catch(e){ /* ignore */ }

    // meeting offer handling: replace <offer_meeting> with selected meetingStatus value.
    // Treat '無し' (trimmed) or empty as no meeting -> remove any lines containing <offer_meeting> and avoid leaving extra blank lines.
    try{
      const meetingSelect = document.getElementById('meetingStatus');
      let meetingVal = '';
      if (meetingSelect){
        // Prefer server-provided raw value stored in dataset.rawValue. If absent, fall back to displayed text.
        const selIdx = meetingSelect.selectedIndex >= 0 ? meetingSelect.selectedIndex : 0;
        const opt = meetingSelect.options && meetingSelect.options[selIdx] ? meetingSelect.options[selIdx] : null;
        let raw = '';
        if (opt && typeof opt.dataset.rawValue === 'string') raw = opt.dataset.rawValue.trim();
        if (raw) meetingVal = raw;
        else if (opt) meetingVal = (opt.textContent || '').toString().trim();
      }
      const norm = (meetingVal || '').toString().trim();
      if (norm && norm !== '無し'){
        out = out.replace(/<\s*offer_meeting\s*>/g, function(){ return '\n' + meetingVal + '\n'; });
      } else {
        // Do not remove entire lines; replace token with empty string
        out = out.replace(/<\s*offer_meeting\s*>/g, '');
      }
    }catch(e){ /* ignore */ }

    // supporter name placeholders
    const family = document.getElementById('supportFamily') ? document.getElementById('supportFamily').value.trim() : '';
    const first = document.getElementById('supportFirst') ? document.getElementById('supportFirst').value.trim() : '';
    const familyFuri = document.getElementById('supportFamilyFuri') ? document.getElementById('supportFamilyFuri').value.trim() : '';
    const firstFuri = document.getElementById('supportFirstFuri') ? document.getElementById('supportFirstFuri').value.trim() : '';
    if (family || first){
      const full = `${family} ${first}`.trim();
      out = out.replace(/<family_name>\s*<first_name>/g, full);
      out = out.replace(/<family_name>/g, family);
      out = out.replace(/<first_name>/g, first);
    }
    if (familyFuri || firstFuri){
      const fullF = `${familyFuri} ${firstFuri}`.trim();
      out = out.replace(/<family_name_furi>\s*<first_name_furi>/g, fullF);
      out = out.replace(/<family_name_furi>/g, familyFuri);
      out = out.replace(/<first_name_furi>/g, firstFuri);
    }
  }catch(e){ /* ignore */ }
  // additional tokens: <MM>, <DD>, <MM+1>.. <MM+5>, <DD+1>.. <DD+5>, <AA+1>.. <AA+5>
  try{
    const today = new Date();
    // We'll compute business-day offsets using holidays fetched from server
    // synchronous helper that returns a Promise resolved set of ISO date strings
    const getHolidays = async ()=>{
      try{
        // Always try to fetch fresh holidays from server; if that fails, fall back to cached value
        const r = await fetch('/holidays');
        if (r && r.ok){
          const j = await r.json();
          const arr = (j && Array.isArray(j.dates)) ? j.dates : [];
          const s = new Set(arr.map(s=>s));
          cachedHolidaysSet = s;
          return s;
        }
        // fetch failed or non-ok -> use cache if available
        return (cachedHolidaysSet || new Set());
      }catch(e){ return (cachedHolidaysSet || new Set()); }
    };
    const computeBusinessDay = (startDate, offset, holidaysSet)=>{
      // startDate may be a Date or an ISO date string 'YYYY-MM-DD'.
      // Normalize to a local-date-only Date to avoid timezone shifts when parsing strings.
      let dt;
      if (typeof startDate === 'string'){
        const m = startDate.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (m){
          const yy = parseInt(m[1],10); const mm = parseInt(m[2],10); const dd = parseInt(m[3],10);
          dt = new Date(yy, mm-1, dd);
        } else {
          const sd = new Date(startDate);
          dt = new Date(sd.getFullYear(), sd.getMonth(), sd.getDate());
        }
      } else if (startDate instanceof Date){
        dt = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
      } else {
        // fallback to today
        const t = new Date(); dt = new Date(t.getFullYear(), t.getMonth(), t.getDate());
      }
      // If offset is zero return same date. If positive, step forward; if negative, step backward.
      if (!offset || offset === 0) return dt;
      let remaining = Math.abs(offset);
      if (offset > 0){
        while(remaining > 0){
          dt.setDate(dt.getDate() + 1);
          const yyyy = dt.getFullYear();
          const mm = String(dt.getMonth() + 1).padStart(2, '0');
          const dd = String(dt.getDate()).padStart(2, '0');
          const iso = `${yyyy}-${mm}-${dd}`;
          const day = dt.getDay();
          const isWeekend = (day === 0 || day === 6);
          const isHoliday = holidaysSet.has(iso);
          if (!isWeekend && !isHoliday){ remaining--; }
        }
        return dt;
      } else {
        while(remaining > 0){
          dt.setDate(dt.getDate() - 1);
          const yyyy = dt.getFullYear();
          const mm = String(dt.getMonth() + 1).padStart(2, '0');
          const dd = String(dt.getDate()).padStart(2, '0');
          const iso = `${yyyy}-${mm}-${dd}`;
          const day = dt.getDay();
          const isWeekend = (day === 0 || day === 6);
          const isHoliday = holidaysSet.has(iso);
          if (!isWeekend && !isHoliday){ remaining--; }
        }
        return dt;
      }
    };
    // Use Monday-first weekday labels to match server-side compute_business_day_py (Monday=0)
    const weekdays = ['月','火','水','木','金','土','日'];
    // If we already have holidays cached, do synchronous replacements so callers get immediate result.
    if (cachedHolidaysSet && (cachedHolidaysSet instanceof Set)){
      try{ console.debug('using cached holidaysSet size', cachedHolidaysSet.size); }catch(e){}
      const holidaysSet = cachedHolidaysSet;
      const baseMM = today.getMonth() + 1;
      const baseDD = today.getDate();
      const baseAA = weekdays[(today.getDay() + 6) % 7];
      out = out.replace(/<MM>/g, String(baseMM));
      out = out.replace(/<DD>/g, String(baseDD));
      out = out.replace(/<AA>/g, baseAA);
      // Replace arbitrary offset tokens like <MM+N>, <MM-N>, <DD+N>, <AA-N> etc.
      const replaceOffsetTokens = (text, startDate, holidaysSet)=>{
        try{
          // Patterns: allow ASCII +/- and full-width +/− (＋, －)
          const signedNum = /([+\uFF0B\-\uFF0D\u2212\u2213\uFF0B\uFF0D])/; // simplified
          // MM offsets
          text = text.replace(/<\s*MM\s*([+＋\-－])\s*(\d+)\s*>/gi, function(_, sign, num){
            const n = parseInt(num,10);
            const offset = (sign === '+' || sign === '＋') ? n : -n;
            const dt = computeBusinessDay(startDate, offset, holidaysSet);
            return String(dt.getMonth() + 1);
          });
          // DD offsets
          text = text.replace(/<\s*DD\s*([+＋\-－])\s*(\d+)\s*>/gi, function(_, sign, num){
            const n = parseInt(num,10);
            const offset = (sign === '+' || sign === '＋') ? n : -n;
            const dt = computeBusinessDay(startDate, offset, holidaysSet);
            return String(dt.getDate());
          });
          // AA offsets (weekday labels, align with Monday-first mapping used elsewhere)
          text = text.replace(/<\s*AA\s*([+＋\-－])\s*(\d+)\s*>/gi, function(_, sign, num){
            const n = parseInt(num,10);
            const offset = (sign === '+' || sign === '＋') ? n : -n;
            const dt = computeBusinessDay(startDate, offset, holidaysSet);
            return weekdays[(dt.getDay() + 6) % 7];
          });
          return text;
        }catch(e){ return text; }
      };
      out = replaceOffsetTokens(out, today, holidaysSet);
      // ensure footer tokens are resolved synchronously as well
      try{
        const footerVal = document.getElementById('footer') ? document.getElementById('footer').value.trim() : '';
        const additionalVal = document.getElementById('additionalFooter') ? document.getElementById('additionalFooter').value.trim() : '';
        const fcsVal = document.getElementById('fcsFooter') ? document.getElementById('fcsFooter').value.trim() : '';
        const footerTokenRegex = /<\s*footer\s*>/i;
        const additionalTokenRegex = /<\s*additional_footer\s*>/i;
        const fcsTokenRegex = /<\s*fcs_footer\s*>/i;
        const hasFooterTokenLocal = footerTokenRegex.test(out) || additionalTokenRegex.test(out) || fcsTokenRegex.test(out);
        const tantoRegexLocal = /- 担当[\s\S]*?営業時間\s*:\s*[^\n]*(?:\n|$)/im;
        const hasTantoLocal = tantoRegexLocal.test(out);
        if ((footerVal || additionalVal || fcsVal) && (hasFooterTokenLocal || hasTantoLocal)){
          if (hasFooterTokenLocal){ out = replaceFooterBlock(out, footerVal, additionalVal, fcsVal); }
          else if (hasTantoLocal){ out = replaceTantoWithFooter(out, tantoRegexLocal, footerVal); }
        } else {
          out = out.replace(new RegExp('<\\s*footer\\s*>','gi'), '');
          out = out.replace(new RegExp('<\\s*additional_footer\\s*>','gi'), '');
          out = out.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), '');
        }
      }catch(e){ }
      return out;
    }
    // async fetch and replacement
    (async ()=>{
      const holidaysSet = await getHolidays();
      try{ console.debug('holidaysSet size', holidaysSet.size, holidaysSet); }catch(e){}
      // base values
      const baseMM = today.getMonth() + 1;
      const baseDD = today.getDate();
      // JavaScript getDay(): Sunday=0 .. Saturday=6
      // Align with server weekday indexing (Monday=0) by rotating index: python_idx = (js_idx + 6) % 7
      const baseAA = weekdays[(today.getDay() + 6) % 7];
      // set replacements for offsets 0..5
      // <MM> and <DD> correspond to offset 0 (today) per original behavior
      out = out.replace(/<MM>/g, String(baseMM));
      out = out.replace(/<DD>/g, String(baseDD));
      out = out.replace(/<AA>/g, baseAA);
        // Support arbitrary offsets like <MM+N>, <MM-N>, <DD+N>, <AA-N> etc. using replaceOffsetTokens
        out = (function(){ try{ return out; }catch(e){ return out; } })();
        out = (function(txt){ try{ return txt; }catch(e){ return txt; } })(out);
        // call helper defined above to replace offset tokens
        try{ out = replaceOffsetTokens(out, today, holidaysSet); }catch(e){}
      // After async replacements, ensure footer tokens are also resolved
      try{
        const footerVal = document.getElementById('footer') ? document.getElementById('footer').value.trim() : '';
        const additionalVal = document.getElementById('additionalFooter') ? document.getElementById('additionalFooter').value.trim() : '';
        const fcsVal = document.getElementById('fcsFooter') ? document.getElementById('fcsFooter').value.trim() : '';
        const footerTokenRegex = /<\s*footer\s*>/i;
        const additionalTokenRegex = /<\s*additional_footer\s*>/i;
        const fcsTokenRegex = /<\s*fcs_footer\s*>/i;
        const hasFooterTokenLocal = footerTokenRegex.test(out) || additionalTokenRegex.test(out) || fcsTokenRegex.test(out);
        const tantoRegexLocal = /- 担当[\s\S]*?営業時間\s*:\s*[^\n]*(?:\n|$)/im;
        const hasTantoLocal = tantoRegexLocal.test(out);
        if ((footerVal || additionalVal || fcsVal) && (hasFooterTokenLocal || hasTantoLocal)){
          if (hasFooterTokenLocal){
            out = replaceFooterBlock(out, footerVal, additionalVal, fcsVal);
          } else if (hasTantoLocal){
            out = replaceTantoWithFooter(out, tantoRegexLocal, footerVal);
          }
        } else {
          out = out.replace(new RegExp('<\\s*footer\\s*>','gi'), '');
          out = out.replace(new RegExp('<\\s*additional_footer\\s*>','gi'), '');
          out = out.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), '');
        }
      }catch(e){ /* ignore footer replacement errors */ }
      // After async replacements, update preview if the panel is present
      const panel = document.getElementById('largePanel');
      if (panel){
        out = ensureFooterSpacing(out, footerVal, additionalVal, fcsVal);
        panel.textContent = sanitizeBlankLines(out);
      }
    })();
  }catch(e){ /* ignore date errors */ }
  // basic fields
  const cname = document.getElementById('customerName') ? document.getElementById('customerName').value.trim() : '';
  const srElem = document.getElementById('srNumberInput');
  const sr = srElem ? (srElem.value ? srElem.value.trim() : '') : '';
  const titleVal = document.getElementById('title') ? document.getElementById('title').value.trim() : '';
  const contentVal = document.getElementById('content') ? document.getElementById('content').value.trim() : '';
    // Determine NextC: prefer selected option's value (which maps to nextc.json value),
    // fall back to selected option text, then to empty.
    let nextc = '';
    try{
      const nc = document.getElementById('nextC');
      if (nc){
        if (nc.selectedIndex >= 0){
          // prefer option.value as the canonical nextc replacement
          nextc = (nc.options[nc.selectedIndex].value || '').toString().trim() || (nc.options[nc.selectedIndex].textContent || '').toString().trim();
        } else {
          nextc = (nc.value || '').toString().trim();
        }
      }
    }catch(e){ nextc = ''; }

  if (cname){ out = out.replace(/<customer_name>/g, cname); }
  if (sr){ out = out.replace(/<sr_number>/g, sr); }
  if (contentVal){ out = out.replace(/<sr_body>/g, contentVal); out = out.replace(/<content>/g, contentVal); }
  if (nextc){ out = out.replace(/<nextc>/g, nextc); out = out.replace(/<nextC>/g, nextc); }

  // phone status placeholder handling: if phone_status is empty, remove lines containing the token
  try{
    const phoneVal = document.getElementById('phoneStatus') ? document.getElementById('phoneStatus').value.trim() : '';
    if (phoneVal){ out = out.replace(/<phone_status>/g, phoneVal); }
    else { out = out.replace(/<phone_status>/g, ''); }
  }catch(e){ /* ignore */ }

  // meeting offer handling: replace <offer_meeting> with selected meetingStatus value.
  // Treat empty or '無し' as no meeting -> replace the token line with a single newline to preserve paragraph spacing.
  try{
    const meetingValRaw = document.getElementById('meetingStatus') ? document.getElementById('meetingStatus').value.trim() : '';
    let meetingVal = meetingValRaw;
    if (!meetingVal){
      const ms = document.getElementById('meetingStatus');
      if (ms && ms.selectedIndex >= 0){
        meetingVal = (ms.options[ms.selectedIndex].textContent || '').toString().trim();
      }
      if (!meetingVal && ms && ms.options && ms.options.length){ meetingVal = (ms.options[0].textContent || '').toString().trim(); }
    }
    const norm = (meetingVal || '').toString().trim();
    if (norm && norm !== '無し'){
      out = out.replace(/<\s*offer_meeting\s*>/g, function(){ return '\n' + meetingVal + '\n'; });
    } else {
      out = out.replace(/<\s*offer_meeting\s*>/g, '');
    }
  }catch(e){ /* ignore */ }

  // supporter name placeholders
  const family = document.getElementById('supportFamily') ? document.getElementById('supportFamily').value.trim() : '';
  const first = document.getElementById('supportFirst') ? document.getElementById('supportFirst').value.trim() : '';
  const familyFuri = document.getElementById('supportFamilyFuri') ? document.getElementById('supportFamilyFuri').value.trim() : '';
  const firstFuri = document.getElementById('supportFirstFuri') ? document.getElementById('supportFirstFuri').value.trim() : '';
  if (family || first){
    const full = `${family} ${first}`.trim();
    out = out.replace(/<family_name>\s*<first_name>/g, full);
    out = out.replace(/<family_name>/g, family);
    out = out.replace(/<first_name>/g, first);
  }
  if (familyFuri || firstFuri){
    const fullF = `${familyFuri} ${firstFuri}`.trim();
    out = out.replace(/<family_name_furi>\s*<first_name_furi>/g, fullF);
    out = out.replace(/<family_name_furi>/g, familyFuri);
    out = out.replace(/<first_name_furi>/g, firstFuri);
  }

  // Final fallback: ensure remaining date tokens are replaced if any escaped earlier
  try{
    const today = new Date();
    const tomorrow = new Date(today.getTime() + 24*60*60*1000);
    const mm = today.getMonth() + 1;
    const dd = today.getDate();
    const weekdays = ['日','月','火','水','木','金','土'];
    // If we have a cached holidays set available, prefer computing business-days synchronously using it.
    // Otherwise, avoid replacing +2..+5 here so the async replacement (which fetches holidays) can apply correct business-days.
    const holidaysAvailable = (cachedHolidaysSet !== null);
    // MM+1 / DD+1: prefer business-day if holidays available, otherwise calendar-tomorrow
    let dt1;
    if (holidaysAvailable){
      try{ dt1 = computeBusinessDay(today, 1, cachedHolidaysSet); }catch(e){ dt1 = tomorrow; }
    } else {
      dt1 = tomorrow;
    }
    const mm1 = dt1.getMonth() + 1;
    const dd1 = dt1.getDate();
    const aa = weekdays[today.getDay()];
    const aa1 = weekdays[dt1.getDay()];
    out = out.replace(/<MM\+1>/g, String(mm1));
    out = out.replace(/<DD\+1>/g, String(dd1));
    out = out.replace(/<MM>/g, String(mm));
    out = out.replace(/<DD>/g, String(dd));
    out = out.replace(/<AA\+1>/g, aa1);
    out = out.replace(/<AA>/g, aa);
    if (holidaysAvailable){
      for (let n=2; n<=5; n++){
        try{
          const dtN = computeBusinessDay(today, n, cachedHolidaysSet);
          const mmn = dtN.getMonth() + 1;
          const ddn = dtN.getDate();
          const aan = weekdays[dtN.getDay()];
          out = out.replace(new RegExp(`<MM\\+${n}>`,`g`), String(mmn));
          out = out.replace(new RegExp(`<\\s*MM\\s*\\+\\s*${n}\\s*>`,`g`), String(mmn));
          out = out.replace(new RegExp(`<MM＋${n}>`,`g`), String(mmn));
          out = out.replace(new RegExp(`<\\s*MM\\s*＋\\s*${n}\\s*>`,`g`), String(mmn));
          out = out.replace(new RegExp(`<DD\\+${n}>`,`g`), String(ddn));
          out = out.replace(new RegExp(`<\\s*DD\\s*\\+\\s*${n}\\s*>`,`g`), String(ddn));
          out = out.replace(new RegExp(`<DD＋${n}>`,`g`), String(ddn));
          out = out.replace(new RegExp(`<\\s*DD\\s*＋\\s*${n}\\s*>`,`g`), String(ddn));
          out = out.replace(new RegExp(`<AA\\+${n}>`,`g`), aan);
          out = out.replace(new RegExp(`<\\s*AA\\s*\\+\\s*${n}\\s*>`,`g`), aan);
          out = out.replace(new RegExp(`<AA＋${n}>`,`g`), aan);
          out = out.replace(new RegExp(`<\\s*AA\\s*＋\\s*${n}\\s*>`,`g`), aan);
        }catch(e){ /* if computation fails for any n, skip */ }
      }
    }
    // also replace standalone YYYY pattern if present
    out = out.replace(/<YYYY>/g, String(today.getFullYear()));
  }catch(e){ /* ignore */ }

  return out;
}

// Apply only synchronous placeholders (customer, SR, title, content, nextc, phone, meeting, supporter)
function applySyncPlaceholders(template){
  if (!template) return template || '';
  let out = template;
  try{
    // Prefer values from cached JSON (server) when available, otherwise fall back to DOM inputs
    const footerJson = window.__footerJson || {};
    const ownerJson = window.__owner || {};
    const phoneItems = window.__phoneItems || [];
    const meetingItems = window.__meetingItems || [];
    const nextcItems = window.__nextcItems || [];

    const cname = document.getElementById('customerName') ? document.getElementById('customerName').value.trim() : '';
    const srElem = document.getElementById('srNumberInput');
    const sr = srElem ? (srElem.value ? srElem.value.trim() : '') : '';
    const titleVal = document.getElementById('title') ? document.getElementById('title').value.trim() : '';
    const contentVal = document.getElementById('content') ? document.getElementById('content').value.trim() : '';
    let nextc = '';
    // if server provided a nextc list, prefer the item's label when selected index is present
    try{
      const nextcElem = document.getElementById('nextC');
      if (nextcElem){
        if (nextcElem.selectedIndex >= 0){
          // prefer option.value (the nextc.json value) then fallback to label text
          nextc = (nextcElem.options[nextcElem.selectedIndex].value || '').trim() || (nextcElem.options[nextcElem.selectedIndex].textContent || '').trim();
        } else nextc = (nextcElem.value || '').trim();
      }
      if (!nextc && Array.isArray(nextcItems) && nextcItems.length){
        const found = nextcItems.find(it => it && it.label === '無し');
        if (found) nextc = (found && typeof found.value === 'string' && found.value.trim() !== '') ? found.value : '';
      }
    }catch(e){}

    if (cname){ out = out.replace(/<customer_name>/g, cname); }
    if (titleVal){ out = out.replace(/<sr_title>/g, titleVal); out = out.replace(/<title>/g, titleVal); }
    if (contentVal){ out = out.replace(/<sr_body>/g, contentVal); out = out.replace(/<content>/g, contentVal); }
    if (nextc){ out = out.replace(/<nextc>/g, nextc); out = out.replace(/<nextC>/g, nextc); }

    // Populate <month>, <day>, <weekday>, <time> from schedule inputs when present
    try{
      const schedDateElem = document.getElementById('schedDate');
      const schedTimeElem = document.getElementById('schedTime');
      const sd = schedDateElem && schedDateElem.value ? new Date(schedDateElem.value) : null;
      const st = schedTimeElem && schedTimeElem.value ? schedTimeElem.value : '';
      if (sd instanceof Date && !isNaN(sd.getTime())){
        const months = sd.getMonth() + 1;
        const days = sd.getDate();
        const weekdays = ['日','月','火','水','木','金','土'];
        const w = weekdays[sd.getDay()];
        out = out.replace(/<month>/g, String(months));
        out = out.replace(/<day>/g, String(days));
        out = out.replace(/<weekday>/g, String(w));
        // time handling: support <time> and <time+N> (minutes offset). Format HH:MM
        const formatTimeWithOffset = (base, offsetMinutes)=>{
          if (!base) return '';
          const m = (''+base).toString().match(/^(\d{1,2}):(\d{2})$/);
          if (!m) return base;
          let hh = parseInt(m[1],10); let mm = parseInt(m[2],10);
          let total = hh*60 + mm + (offsetMinutes || 0);
          const dayMinutes = 24*60;
          total = ((total % dayMinutes) + dayMinutes) % dayMinutes;
          const rh = String(Math.floor(total/60)).padStart(2,'0');
          const rm = String(total%60).padStart(2,'0');
          return `${rh}:${rm}`;
        };
        if (st){
          out = out.replace(/<\s*time\s*([+＋\-－])\s*(\d+)\s*>/gi, function(_, sign, mins){
            const n = parseInt(mins || '0', 10) || 0;
            const offset = (sign === '+' || sign === '＋') ? n : -n;
            return formatTimeWithOffset(st, offset);
          });
          out = out.replace(/<\s*time\s*>/gi, function(){ return formatTimeWithOffset(st, 0); });
        }
      }
    }catch(e){ /* ignore */ }

    // phone status
    try{
      // Prefer server-provided phone label when available
      let phoneVal = '';
      const phoneElem = document.getElementById('phoneStatus');
      if (phoneElem){
        if (phoneElem.selectedIndex >= 0){
          const opt = phoneElem.options[phoneElem.selectedIndex];
          const raw = opt && typeof opt.dataset.rawValue === 'string' ? opt.dataset.rawValue.trim() : '';
          // Use server-provided raw value only. Do not use the visible label as a fallback.
          if (raw) phoneVal = raw;
          else phoneVal = '';
        } else {
          // No selected index: try to match the select's value to a known item value
          const valCandidate = (phoneElem.value || '').trim();
          if (valCandidate && Array.isArray(phoneItems)){
            const match = phoneItems.find(it => it && typeof it.value === 'string' && it.value === valCandidate);
            if (match && typeof match.value === 'string' && match.value.trim() !== '') phoneVal = match.value;
          }
        }
      }
      if (!phoneVal && Array.isArray(phoneItems) && phoneItems.length){
        const found = phoneItems.find(it => it && it.label === '無し');
        if (found && typeof found.value === 'string' && found.value.trim() !== '') phoneVal = found.value;
      }
      if (phoneVal){ out = out.replace(/<phone_status>/g, phoneVal); }
      else { out = out.replace(/<phone_status>/g, ''); }
    }catch(e){ }

    // meeting offer handling (prefer selected value/label)
    // Treat '無し' or empty as no meeting -> remove lines containing <offer_meeting>
    try{
      const meetingSelect = document.getElementById('meetingStatus');
      let meetingVal = '';
      if (meetingSelect){
        const selIdx = meetingSelect.selectedIndex >= 0 ? meetingSelect.selectedIndex : 0;
        const opt = meetingSelect.options && meetingSelect.options[selIdx] ? meetingSelect.options[selIdx] : null;
        let raw = '';
        if (opt && typeof opt.dataset.rawValue === 'string') raw = opt.dataset.rawValue.trim();
        if (raw) meetingVal = raw;
        else if (opt) meetingVal = (opt.textContent || '').toString().trim();
      }
      // If server provided meeting list and no selection, try to use its first non-empty label
      if (!meetingVal && Array.isArray(meetingItems) && meetingItems.length){
        const first = meetingItems[0];
        if (first && typeof first.value === 'string' && first.value.trim() !== '') meetingVal = first.value.toString().trim();
      }
      const norm = (meetingVal || '').toString().trim();
      if (norm && norm !== '無し'){
        out = out.replace(/<offer_meeting>/g, meetingVal);
      } else {
        out = out.replace(/<offer_meeting>/g, '');
      }
    }catch(e){ }

    // supporter name placeholders
    // supporter name placeholders: prefer server-provided owner JSON
    const family = (ownerJson.family_name && ownerJson.family_name.toString().trim()) || (document.getElementById('supportFamily') ? document.getElementById('supportFamily').value.trim() : '');
    const first = (ownerJson.first_name && ownerJson.first_name.toString().trim()) || (document.getElementById('supportFirst') ? document.getElementById('supportFirst').value.trim() : '');
    const familyFuri = (ownerJson.family_name_furi && ownerJson.family_name_furi.toString().trim()) || (document.getElementById('supportFamilyFuri') ? document.getElementById('supportFamilyFuri').value.trim() : '');
    const firstFuri = (ownerJson.first_name_furi && ownerJson.first_name_furi.toString().trim()) || (document.getElementById('supportFirstFuri') ? document.getElementById('supportFirstFuri').value.trim() : '');
    if (family || first){ const full = `${family} ${first}`.trim(); out = out.replace(/<family_name>\s*<first_name>/g, full); out = out.replace(/<family_name>/g, family); out = out.replace(/<first_name>/g, first); }
    if (familyFuri || firstFuri){ const fullF = `${familyFuri} ${firstFuri}`.trim(); out = out.replace(/<family_name_furi>\s*<first_name_furi>/g, fullF); out = out.replace(/<family_name_furi>/g, familyFuri); out = out.replace(/<first_name_furi>/g, firstFuri); }

    // SR-derived fields: prefer SR map entries when SR number matches
    try{
      if (sr && window.__srMap && window.__srMap[sr]){
        const meta = window.__srMap[sr];
        if (meta.replace_name) out = out.replace(/<customer_name>/g, meta.replace_name);
        else if (meta.contact) out = out.replace(/<customer_name>/g, meta.contact);
        if (meta.customer_title) out = out.replace(/<sr_title>/g, meta.customer_title);
        else if (meta.internal_title) out = out.replace(/<sr_title>/g, meta.internal_title);
        if (meta.customer_title) out = out.replace(/<sr_body>/g, meta.customer_title);
      }
    }catch(e){ }

    // Footer tokens: prefer cached footer JSON when available
    try{
      const footerJson = window.__footerJson || {};
      if (footerJson && (typeof footerJson.footer === 'string' || typeof footerJson.additional_footer === 'string' || typeof footerJson.fcs_footer === 'string')){
        if (typeof footerJson.footer === 'string') out = out.replace(new RegExp('<\\s*footer\\s*>','gi'), footerJson.footer + '\n');
        else out = out.replace(new RegExp('<\\s*footer\\s*>','gi'), '');
        if (typeof footerJson.additional_footer === 'string'){
          const av = footerJson.additional_footer;
          if (av === ''){
            out = out.replace(new RegExp('\\n?\\s*<\\s*additional_footer\\s*>\\s*\\n?','gi'), '\n');
          } else {
            out = out.replace(new RegExp('<\\s*additional_footer\\s*>','gi'), '\n' + av + '\n');
          }
        } else {
          out = out.replace(new RegExp('\\n?\\s*<\\s*additional_footer\\s*>\\s*\\n?','gi'), '');
        }
        if (typeof footerJson.fcs_footer === 'string') out = out.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), '\n' + footerJson.fcs_footer + '\n');
        else out = out.replace(new RegExp('\\n?\\s*<\\s*fcs_footer\\s*>\\s*\\n?','gi'), '\n');
      } else {
        out = out.replace(new RegExp('<\\s*footer\\s*>','gi'), '');
        out = out.replace(new RegExp('<\\s*additional_footer\\s*>','gi'), '');
        out = out.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), '');
      }
    }catch(e){ }
  }catch(e){ /* ignore */ }
  try{ return sanitizeBlankLines(out); }catch(e){ return out; }
}

// Try server-side rendering of template (dates/business-day). Returns rendered string or null on failure.
async function renderWithServer(template){
  try{
    const resp = await fetch('/render_template', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({template})});
    if (!resp.ok) return null;
    const j = await resp.json();
    if (j && j.status === 'ok' && typeof j.rendered === 'string') return j.rendered;
  }catch(e){ console.warn('renderWithServer error', e); }
  return null;
}

// Normalize blank lines and trailing spaces before rendering preview
function sanitizeBlankLines(str){
  try{
    // normalize CRLF to LF
    // Work in LF internally for simplicity
    str = str.replace(/\r\n?/g, '\n');
    // trim trailing spaces on each line
    str = str.replace(/[ \t]+$/gm, '');
    // (removed) previously collapsed 3+ consecutive newlines into two — keep original spacing
    // remove leading blank lines
    str = str.replace(/^\s*\n+/, '');
    // reduce trailing newlines to a single newline
    str = str.replace(/\n{2,}$/, '\n');
    // Finally convert normalized LF back to CRLF for output consistency
    try{ str = str.replace(/\n/g, '\r\n'); }catch(e){}
    // Ensure the string ends with a single CRLF
    if (!str.endsWith('\r\n')) str = str + '\r\n';
    return str;
  }catch(e){ return str; }
}

// Replace <footer> token while preserving or ensuring a blank line before it
function replaceFooterToken(out, footerVal){
  try{
    if (!footerVal) return out.replace(new RegExp('<\\s*footer\\s*>','gi'), '');
    return out.replace(new RegExp('<\\s*footer\\s*>','gi'), function(match, offset, string){
      const before = string.slice(0, offset);
      // if there is already a blank line before, insert footer directly
      if (/\n\s*\n$/.test(before)) return footerVal + '\n';
      // if there is a single newline, make it one blank line before footer
      if (/\n$/.test(before)) return '\n' + footerVal + '\n';
      // otherwise ensure at least one newline before footer (do not force two)
      return '\n' + footerVal + '\n';
    });
  }catch(e){ return out; }
}

// Replace a 担当 block (tanto regex) with footer text while preserving/ensuring a blank line before it
function replaceTantoWithFooter(out, regex, footerVal){
  try{
    if (!footerVal) return out.replace(regex, '');
    return out.replace(regex, function(match, offset, string){
      const before = string.slice(0, offset);
      if (/\n\s*\n$/.test(before)) return footerVal + '\n';
      if (/\n$/.test(before)) return '\n' + footerVal + '\n';
      // no trailing newline before match: ensure a single newline only
      return '\n' + footerVal + '\n';
    });
  }catch(e){ return out; }
}

// Replace a sequence of <footer>, <additional_footer>, <fcs_footer> (with optional newlines/space between)
// in one pass, producing a combined footer block while preserving/ensuring a blank line before it.
function replaceFooterBlock(out, footerVal, additionalVal, fcsVal){
  try{
    // If all values empty -> remove the entire token line(s) including surrounding newlines
    if (!footerVal && !additionalVal && !fcsVal){
      const removal = /(?:\r?\n)?\s*<\s*footer\s*>\s*(?:\r?\n)?\s*(?:<\s*additional_footer\s*>\s*(?:\r?\n)?\s*)?(?:<\s*fcs_footer\s*>\s*(?:\r?\n)?\s*)?/gi;
      return out.replace(removal, '');
    }
    // For insertion, match the tokens but DO NOT consume a preceding newline: preserve any existing newline
    const tokenSeq = /<\s*footer\s*>\s*(?:\r?\n)?\s*(?:<\s*additional_footer\s*>\s*(?:\r?\n)?\s*)?(?:<\s*fcs_footer\s*>\s*(?:\r?\n)?\s*)?/gi;
    return out.replace(tokenSeq, function(match, offset, string){
      const before = string.slice(0, offset);
      let prefix = '\n';
      if (/\n\s*\n$/.test(before)) prefix = '';
      else if (/\n$/.test(before)) prefix = '\n';
      const parts = [];
      // Only include parts for tokens that were actually present in the matched substring
      const hasFooterTok = /<\s*footer\s*>/i.test(match);
      const hasAdditionalTok = /<\s*additional_footer\s*>/i.test(match);
      const hasFcsTok = /<\s*fcs_footer\s*>/i.test(match);
      if (hasFooterTok && footerVal) parts.push(footerVal);
      if (hasAdditionalTok){
        if (additionalVal !== undefined && additionalVal !== null){
          if (additionalVal === ''){
            parts.push('');
          } else {
            parts.push('', additionalVal, '');
          }
        }
      }
      if (hasFcsTok && fcsVal) parts.push(fcsVal);
      const block = parts.join('\n') + '\n';
      return prefix + block;
    });
  }catch(e){ return out; }
}

// Ensure the combined footer block appears with at least one blank line before it
function ensureFooterSpacing(out, footerVal, additionalVal, fcsVal){
  try{
    const parts = [];
    if (footerVal) parts.push(footerVal);
    if (additionalVal) parts.push(additionalVal);
    if (fcsVal) parts.push(fcsVal);
    if (parts.length === 0) return out;
    const block = parts.join('\n');
    let idx = out.indexOf(block);
    // if combined block not found, try with footerVal only
    if (idx === -1 && footerVal) idx = out.indexOf(footerVal);
    if (idx === -1) return out;
    // ensure at least one blank line before idx (do not force two)
    const beforeIdx = idx - 1;
    if (beforeIdx < 0) return out; // start of string
    const beforeChar = out[beforeIdx];
    if (beforeChar === '\n'){
      // there's at least one newline before block -> OK
      return out;
    } else {
      // insert a single newline before block
      return out.slice(0, idx) + '\n' + out.slice(idx);
    }
  }catch(e){ return out; }
}

templateChangeBtn.addEventListener('click', function(){
  const emailType = document.getElementById('emailType').value;
  const panel = document.getElementById('largePanel');
  if (emailType){
    // try load per-type template from server first
    fetch(`/templates/${encodeURIComponent(emailType)}`).then(r=>{
      if (r.ok) return r.json();
      throw new Error('no per-type');
    }).then(j=>{
      const raw = (j && j.template) ? j.template : (templatesStore[emailType] || getDefaultRawTemplateFor(emailType) || panel.textContent || '');
      templatesStore[emailType] = raw;
      saveTemplatesToStorage();
      // determine source file: prefer the selected option's dataset.source (optgroup file)
      // to avoid ambiguous key lookups on the server; fall back to server-provided source_file
      let sf = null;
      try{
        const sel = document.getElementById('emailType');
        if (sel && sel.selectedIndex >= 0){
          const opt = sel.options[sel.selectedIndex];
          if (opt && opt.dataset && opt.dataset.source) sf = opt.dataset.source;
        }
      }catch(e){ sf = null; }
      if (!sf) sf = (j && j.source_file) ? j.source_file : null;
      // protect against accidental 'save_in_file...' filenames returned by server
      if (sf && /^save_in_file/i.test(sf)) sf = null;
      __editingSourceFile = sf || null;
      openModalWithContent(raw);
    }).catch(()=>{
      const raw = templatesStore[emailType] || getDefaultRawTemplateFor(emailType) || panel.textContent || '';
      templatesStore[emailType] = raw;
      saveTemplatesToStorage();
      // also ensure footer is loaded before opening modal
      __editingSourceFile = null;
      loadFooterFromServer().then(()=> openModalWithContent(raw));
    });
  } else {
    openModalWithContent(panel.textContent || '');
  }
});

// Footer save button handler
const footerSaveBtn = document.getElementById('footerSaveBtn');
if (footerSaveBtn){
  footerSaveBtn.addEventListener('click', function(){
    const footerVal = document.getElementById('footer').value;
    // POST to server
    fetch('/footer', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({footer: footerVal})}).then(r=>{
      if (r.ok){
          // refresh cached footer from server
          loadFooterFromServer().then(()=>{ alert('Footer saved'); }).catch(()=>{ alert('Footer save failed'); });
      } else {
        alert('Footer save failed');
      }
    }).catch(()=>{ alert('Footer save failed'); });
  });
}

const ownerSaveBtn = document.getElementById('ownerSaveBtn');
if (ownerSaveBtn){
  ownerSaveBtn.addEventListener('click', function(){
    const owner = {
      family_name: document.getElementById('supportFamily').value.trim(),
      first_name: document.getElementById('supportFirst').value.trim(),
      family_name_furi: document.getElementById('supportFamilyFuri').value.trim(),
      first_name_furi: document.getElementById('supportFirstFuri').value.trim()
    };
    // Save owner first
    fetch('/owner', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({owner})}).then(r=>{
      if (r.ok){
        // then save footer content as well
          const footerVal = document.getElementById('footer') ? document.getElementById('footer').value : '';
          const additionalVal = document.getElementById('additionalFooter') ? document.getElementById('additionalFooter').value : '';
          const fcsVal = document.getElementById('fcsFooter') ? document.getElementById('fcsFooter').value : '';
          return fetch('/footer', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({footer: footerVal, additional_footer: additionalVal, fcs_footer: fcsVal})}).then(fr=>{
            if (fr.ok){
              // refresh cached footer after saving
              loadFooterFromServer().then(()=>{ alert('Owner and footers saved'); }).catch(()=>{ alert('Owner and footers saved'); });
            } else { alert('Owner saved but footer save failed'); }
          }).catch(()=>{ alert('Owner saved but footer save failed'); });
      } else {
        alert('Owner save failed');
      }
    }).catch(()=>{ alert('Owner save failed'); });
  });
}

saveTemplateBtn.addEventListener('click', function(){
  const newText = editor.value;
  const emailType = document.getElementById('emailType').value;
  const panel = document.getElementById('largePanel');
  if (emailType){
    templatesStore[emailType] = newText;
    saveTemplatesToStorage();
    // Also attempt to persist to server per-type first, then aggregate as fallback
    try{
      // If we have a source file for this template, prefer saving directly into that file
      if (__editingSourceFile){
        fetch('/templates/save_in_file_explicit', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({template: newText, source_file: __editingSourceFile, key: emailType})}).then(r=>{
          if (!r.ok){
            // fallback to per-type POST
            return fetch(`/templates/${encodeURIComponent(emailType)}`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({template: newText})}).catch(()=>{});
          }
        }).catch(()=>{});
      } else {
        fetch(`/templates/${encodeURIComponent(emailType)}`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({template: newText})}).then(r=>{
          if (!r.ok){
            // fallback to aggregate POST
            return fetch('/templates', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({templates: {[emailType]: newText}})}).catch(()=>{});
          }
        }).catch(()=>{});
      }
    }catch(e){ }
    // apply replacements to show updated preview — prefer server-side render for accurate business-day dates
    let out = newText;
    (async ()=>{
      const panelLocal = panel;
      try{
        const server = await renderWithServer(out);
        if (server){
          panelLocal.textContent = sanitizeBlankLines(applySyncPlaceholders(server));
        } else {
          const clientOut = applyTemplatePlaceholders(out);
          panelLocal.textContent = sanitizeBlankLines(clientOut);
        }
      }catch(e){ const clientOut = applyTemplatePlaceholders(out); panelLocal.textContent = sanitizeBlankLines(clientOut); }
    })();
    // applyTemplatePlaceholders already performs customer/SR/title/content replacements
    // avoid repeating them here to prevent accidental token collisions
    // Replace footer tokens in template if present. Do NOT append footer unconditionally.
    // Use case-insensitive, space-tolerant regexes so templates with varied formatting still match.
    const footerVal = document.getElementById('footer') ? document.getElementById('footer').value.trim() : '';
    const additionalVal = document.getElementById('additionalFooter') ? document.getElementById('additionalFooter').value.trim() : '';
    const fcsVal = document.getElementById('fcsFooter') ? document.getElementById('fcsFooter').value.trim() : '';
    const footerTokenRegex = /<\s*footer\s*>/i;
    const additionalTokenRegex = /<\s*additional_footer\s*>/i;
    const fcsTokenRegex = /<\s*fcs_footer\s*>/i;
    const hasFooterToken = footerTokenRegex.test(out) || additionalTokenRegex.test(out) || fcsTokenRegex.test(out);
    const tantoRegex = /- 担当[\s\S]*?営業時間\s*:\s*[^\n]*(?:\n|$)/im;
    const hasTantoBlock = tantoRegex.test(out);
        if ((footerVal || additionalVal || fcsVal) && (hasFooterToken || hasTantoBlock)){
      if (hasFooterToken){
        out = replaceFooterBlock(out, footerVal, additionalVal, fcsVal);
      } else if (hasTantoBlock){
        out = replaceTantoWithFooter(out, tantoRegex, footerVal);
      }
        } else {
      // remove footer tokens if present but no footer provided or no suitable insertion point
      out = out.replace(new RegExp('<\\s*footer\\s*>','gi'), '');
      out = out.replace(new RegExp('<\\s*additional_footer\\s*>','gi'), '');
      out = out.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), '');
    }
    // ensureFooterSpacing is handled above in server-render path; for client fallback we ensure it here as well
    // (the async renderWithServer call will update panel; this ensures immediate footer tokens are applied for the synchronous path only)
    try{
      let syncOut = applyTemplatePlaceholders(newText);
      syncOut = ensureFooterSpacing(syncOut, footerVal, additionalVal, fcsVal);
      panel.textContent = sanitizeBlankLines(syncOut);
    }catch(e){ /* ignore */ }
    // Save footer back to server as well (fire-and-forget)
    try{ fetch('/footer', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({footer: footerVal, additional_footer: additionalVal, fcs_footer: fcsVal})}).catch(()=>{}); }catch(e){}
  } else {
    panel.textContent = newText;
  }
  // after saving, clear editing source file so subsequent edits default to per-type behavior
  __editingSourceFile = null;
  closeModal();
});

cancelTemplateBtn.addEventListener('click', function(){ closeModal(); });

// Close modal when clicking outside content
modal.addEventListener('click', function(e){ if (e.target === modal){ closeModal(); } });

// Replace/Reset handlers for customer name and sympton
const replaceNameBtn = document.getElementById('replaceNameBtn');
const resetNameBtn = document.getElementById('resetNameBtn');
if (replaceNameBtn){
  replaceNameBtn.addEventListener('click', async function(){
    const caseNo = (document.getElementById('srNumberInput') && document.getElementById('srNumberInput').value) ? document.getElementById('srNumberInput').value.trim() : '';
    if (!caseNo){ alert('Please enter SR number'); return; }
    const newName = (document.getElementById('customerName') && document.getElementById('customerName').value) ? document.getElementById('customerName').value.trim() : '';
    if (!newName){ alert('Please enter a name to replace'); return; }
    const symptonVal = (document.getElementById('content') && document.getElementById('content').value) ? document.getElementById('content').value : '';
    const feedback = document.getElementById('replaceFeedback');
    if (!confirm('Replace name for SR ' + caseNo + '?')) return;
    try{
      const res = await postJson('/sr/replace', { case_number: caseNo, action: 'set', replace_name: newName, sympton: symptonVal });
      if (res && res.status === 'ok'){
        if (!window.__srMap) window.__srMap = {};
        if (!window.__srMap[caseNo]) window.__srMap[caseNo] = {};
        window.__srMap[caseNo].replace_name = newName;
        window.__srMap[caseNo].sympton = symptonVal;
        if (feedback) { feedback.style.display = 'inline'; setTimeout(()=> feedback.style.display='none', 1500); }
      } else {
        alert('Error: ' + (res && res.message ? res.message : 'failed'));
      }
    }catch(e){ alert('Error: ' + e); }
  });
}
if (resetNameBtn){
  resetNameBtn.addEventListener('click', async function(){
    const caseNo = (document.getElementById('srNumberInput') && document.getElementById('srNumberInput').value) ? document.getElementById('srNumberInput').value.trim() : '';
    if (!caseNo){ alert('Please enter SR number'); return; }
    if (!confirm('Reset replaced name for this SR?')) return;
    const feedback = document.getElementById('replaceFeedback');
    try{
      const res = await postJson('/sr/replace', { case_number: caseNo, action: 'reset' });
      if (res && res.status === 'ok'){
        if (window.__srMap && window.__srMap[caseNo]) delete window.__srMap[caseNo].replace_name;
        const item = window.__srMap && window.__srMap[caseNo];
        // Prefer contact for display after reset
        document.getElementById('customerName').value = (item && (item.contact || item.replace_name)) || '';
        try{ const contactEl = document.getElementById('contactInfo'); if (contactEl) contactEl.textContent = (item && item.contact) || ''; }catch(e){}
        if (feedback) { feedback.style.display = 'inline'; setTimeout(()=> feedback.style.display='none', 1200); }
      } else {
        alert('Error: ' + (res && res.message ? res.message : 'failed'));
      }
    }catch(e){ alert('Error: ' + e); }
  });
}

// Show template in largePanel when selecting 'FCS'
const emailSelect = document.getElementById('emailType');
if (emailSelect){
  emailSelect.addEventListener('change', function(){
    try{
      const val = this.value;
      const panel = document.getElementById('largePanel');
      console.debug('emailType changed to', val, 'templatesStore keys:', Object.keys(templatesStore || {}));

      if (!val){ panel.textContent = ''; return; }

      const stored = (templatesStore && templatesStore[val]) ? templatesStore[val] : null;
      // fetch base template per-type and footer, then merge. If per-type fetch returns nothing,
      // fall back to fetching all templates and matching keys (handles multi-key files like fcs.json).
      Promise.all([
        // If the selected option has a source file (dataset.source), prefer that file when fetching
        (async ()=>{
          try{
            const sel = document.getElementById('emailType');
            let src = null;
            if (sel && sel.selectedIndex >= 0){ const opt = sel.options[sel.selectedIndex]; if (opt && opt.dataset && opt.dataset.source) src = opt.dataset.source; }
            const url = src ? `/templates/${encodeURIComponent(val)}?source_file=${encodeURIComponent(src)}` : `/templates/${encodeURIComponent(val)}`;
            const r = await fetch(url);
            return r.ok ? r.json() : null;
          }catch(e){ return null; }
        })(),
        fetch('/footer').then(r=> r.ok ? r.json() : null).catch(()=>null)
      ]).then(async ([tJson, fJson])=>{
        let base = (tJson && tJson.template) ? tJson.template : null;
        // if direct per-type fetch didn't return, try fetching the full templates map and match keys
        if (!base){
          try{
            const allResp = await fetch('/templates');
            if (allResp && allResp.ok){
              const allJson = await allResp.json();
              const map = allJson && allJson.templates ? allJson.templates : null;
              if (map){
                // direct key
                if (map[val]){
                  base = map[val];
                } else {
                  // case-insensitive / trimmed match
                  const normalize = s => (s||'').toString().trim().toLowerCase();
                  const foundKey = Object.keys(map).find(k => normalize(k) === normalize(val));
                  if (foundKey){ base = map[foundKey]; }
                }
              }
            }
          }catch(e){ /* ignore */ }
        }
        // fallback to local templatesStore or default
        base = base || templatesStore[val] || getDefaultRawTemplateFor(val) || '';
        // If no template content is available for this selection, do not inject footer — show empty panel
        if (!base || !base.toString().trim()){
          panel.textContent = '';
          return;
        }
        const footer = (fJson && typeof fJson.footer === 'string') ? fJson.footer : (document.getElementById('footer') ? document.getElementById('footer').value : '');
        const additional_footer = (fJson && typeof fJson.additional_footer === 'string') ? fJson.additional_footer : (document.getElementById('additionalFooter') ? document.getElementById('additionalFooter').value : '');
        const fcs_footer = (fJson && typeof fJson.fcs_footer === 'string') ? fJson.fcs_footer : (document.getElementById('fcsFooter') ? document.getElementById('fcsFooter').value : '');

        let out = applyTemplatePlaceholders(base);
        // Try server-side rendering (business-day date replacements, robust) first
        try{
          const serverRendered = await renderWithServer(out);
          if (serverRendered && typeof serverRendered === 'string'){
            // serverRendered already contains synchronous placeholders resolved server-side,
            // but apply client-side sync placeholders (customer/SR/title/content) to ensure local values (like customer name, SR) are applied.
            let applied = applySyncPlaceholders(serverRendered);
            // Footer handling: prefer server-supplied footer values from fJson above, but ensure local footer fields also applied
            const footerLocal = (window.__footerJson && typeof window.__footerJson.footer === 'string') ? window.__footerJson.footer : (footer || (document.getElementById('footer') ? document.getElementById('footer').value : ''));
            const additionalLocal = (window.__footerJson && typeof window.__footerJson.additional_footer === 'string') ? window.__footerJson.additional_footer : (additional_footer || (document.getElementById('additionalFooter') ? document.getElementById('additionalFooter').value : ''));
            const fcsLocal = (window.__footerJson && typeof window.__footerJson.fcs_footer === 'string') ? window.__footerJson.fcs_footer : (fcs_footer || (document.getElementById('fcsFooter') ? document.getElementById('fcsFooter').value : ''));
            // Use existing footer token logic to inject/clean tokens
            const footerTokenRegex = /<\s*footer\s*>/i;
            const additionalTokenRegex = /<\s*additional_footer\s*>/i;
            const fcsTokenRegex = /<\s*fcs_footer\s*>/i;
            const hasFooterToken = footerTokenRegex.test(applied) || additionalTokenRegex.test(applied) || fcsTokenRegex.test(applied);
            const tantoRegex2 = /- 担当[\s\S]*?営業時間\s*:\s*[^\n]*(?:\n|$)/im;
            const hasTanto = tantoRegex2.test(applied);
            if ((footerLocal || additionalLocal || fcsLocal) && (hasFooterToken || hasTanto)){
              if (hasFooterToken){
                applied = applied.replace(new RegExp('<\\s*footer\\s*>','gi'), footerLocal ? (footerLocal + '\n') : '');
                if (additionalLocal !== undefined && additionalLocal !== null){
                  if (additionalLocal === ''){
                    applied = applied.replace(new RegExp('\\n?\\s*<\\s*additional_footer\\s*>\\s*\\n?','gi'), '\n');
                  } else {
                    applied = applied.replace(new RegExp('<\\s*additional_footer\\s*>','gi'), '\n' + additionalLocal + '\n');
                  }
                } else {
                  applied = applied.replace(new RegExp('\\n?\\s*<\\s*additional_footer\\s*>\\s*\\n?','gi'), '');
                }
                if (fcsLocal){
                  if (additionalLocal || footerLocal){ applied = applied.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), '\n' + fcsLocal + '\n'); }
                  else { applied = applied.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), fcsLocal + '\n'); }
                } else { applied = applied.replace(new RegExp('\\n?\\s*<\\s*fcs_footer\\s*>\\s*\\n?','gi'), '\n'); applied = applied.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), ''); }
              } else if (hasTanto){
                applied = replaceTantoWithFooter(applied, tantoRegex2, footerLocal);
              }
            } else {
              applied = applied.replace(new RegExp('<\\s*footer\\s*>','gi'), '');
              applied = applied.replace(new RegExp('<\\s*additional_footer\\s*>','gi'), '');
              applied = applied.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), '');
            }
            applied = ensureFooterSpacing(applied, footerLocal, additionalLocal, fcsLocal);
            panel.textContent = sanitizeBlankLines(applied);
            return;
          }
        }catch(e){ console.warn('server render failed, falling back to client logic', e); }
        const cname = document.getElementById('customerName').value.trim();
        const srElem3 = document.getElementById('srNumberInput');
        const sr = srElem3 ? (srElem3.value ? srElem3.value.trim() : '') : '';
        const titleVal = document.getElementById('title').value.trim();
        if (cname){ out = out.replace(/<customer_name>/g, cname); }
        if (sr){ out = out.replace(/<sr_number>/g, sr); }
        if (titleVal){ /* preserve marker in template; do not auto-replace bullets */ }
        const contentVal = document.getElementById('content').value.trim();
        if (contentVal){ /* preserve marker in template; do not auto-replace bullets */ }

        // Replace footer tokens in template only when tokens or 担当 block exist; do NOT append footer unconditionally.
        const footerTokenRegex = /<\s*footer\s*>/i;
        const additionalTokenRegex = /<\s*additional_footer\s*>/i;
        const fcsTokenRegex = /<\s*fcs_footer\s*>/i;
        const hasFooterToken = footerTokenRegex.test(out) || additionalTokenRegex.test(out) || fcsTokenRegex.test(out);
        const tantoRegex2 = /- 担当[\s\S]*?営業時間\s*:\s*[^\n]*(?:\n|$)/im;
        const hasTanto = tantoRegex2.test(out);
        if ((footer || additional_footer || fcs_footer) && (hasFooterToken || hasTanto)){
          if (hasFooterToken){
            out = out.replace(new RegExp('<\\s*footer\\s*>','gi'), footer ? (footer + '\n') : '');
            if (additional_footer !== undefined && additional_footer !== null){
              if (additional_footer === ''){
                out = out.replace(new RegExp('\\n?\\s*<\\s*additional_footer\\s*>\\s*\\n?','gi'), '\n');
              } else {
                out = out.replace(new RegExp('<\\s*additional_footer\\s*>','gi'), '\n' + additional_footer + '\n');
              }
            } else {
              out = out.replace(new RegExp('\\n?\\s*<\\s*additional_footer\\s*>\\s*\\n?','gi'), '');
            }
            if (fcs_footer){
              if (additional_footer || footer){
                out = out.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), '\n' + fcs_footer + '\n');
              } else {
                out = out.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), fcs_footer + '\n');
              }
            } else {
              out = out.replace(new RegExp('\\n?\\s*<\\s*fcs_footer\\s*>\\s*\\n?','gi'), '\n');
              out = out.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), '');
            }
          } else if (hasTanto){
            out = replaceTantoWithFooter(out, tantoRegex2, footer);
          }
        } else {
          out = out.replace(new RegExp('<\\s*footer\\s*>','gi'), '');
          out = out.replace(new RegExp('<\\s*additional_footer\\s*>','gi'), '');
          out = out.replace(new RegExp('<\\s*fcs_footer\\s*>','gi'), '');
        }

        console.debug('rendered template length', out.length);
        out = ensureFooterSpacing(out, footer, additional_footer, fcs_footer);
        panel.textContent = sanitizeBlankLines(out);
      }).catch(err=>{ console.error('render combine error', err); panel.textContent = ''; });
    }catch(err){ console.error('render template error', err); }
  });
  // trigger initial render for existing selection
  try{ emailSelect.dispatchEvent(new Event('change')); }catch(e){ }
} else {
  console.error('emailType select not found');
}

// Config modal logic: change template_files directory
const changeDirBtn = document.getElementById('changeTemplateDirBtn');
const configModal = document.getElementById('configModal');
const templateDirInput = document.getElementById('templateDirInput');
const saveConfigBtn = document.getElementById('saveConfigBtn');
const cancelConfigBtn = document.getElementById('cancelConfigBtn');
const pickConfigBtn = document.getElementById('pickConfigBtn');

async function loadConfig(){
  try{
    const res = await fetch('/config');
    if (!res.ok) return null;
    const j = await res.json();
    if (j && j.config && j.config.template_dir) return j.config.template_dir;
  }catch(e){ console.error('loadConfig failed', e); }
  return null;
}

if (changeDirBtn){
  changeDirBtn.addEventListener('click', async function(){
    if (!configModal) return;
    configModal.style.display = 'flex';
    configModal.setAttribute('aria-hidden','false');
    const td = await loadConfig();
    if (td && templateDirInput) templateDirInput.value = td;
    if (templateDirInput) templateDirInput.focus();
  });
}

// Refresh preview button handler (re-render right panel using current inputs)

if (cancelConfigBtn){
  cancelConfigBtn.addEventListener('click', function(){
    if (!configModal) return; configModal.style.display='none'; configModal.setAttribute('aria-hidden','true');
  });
}

if (saveConfigBtn){
  saveConfigBtn.addEventListener('click', async function(){
    const v = templateDirInput ? templateDirInput.value.trim() : '';
    if (!v){ alert('フォルダパスを入力してください'); return; }
    try{
      const res = await fetch('/config', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({template_dir: v})});
      const j = await res.json();
      if (!res.ok || !j || j.status !== 'ok'){
        alert('保存に失敗しました: ' + (j && j.message ? j.message : res.statusText));
        return;
      }
      alert('テンプレート保存先を保存しました: ' + (j.template_dir || v));
      // reload template keys and sr list
      try{ await loadTemplateKeys(); }catch(e){}
      try{ await loadSrList(); }catch(e){}
      if (configModal) { configModal.style.display='none'; configModal.setAttribute('aria-hidden','true'); }
    }catch(e){ alert('保存に失敗しました: ' + e); }
  });
}

if (pickConfigBtn){
  pickConfigBtn.addEventListener('click', async function(){
    try{
      const current = templateDirInput ? templateDirInput.value.trim() : '';
      const res = await fetch('/config/pick', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({initial: current})});
      const j = await res.json();
      if (j && j.status === 'ok' && j.template_dir){
        templateDirInput.value = j.template_dir;
      } else if (j && j.status === 'cancelled'){
        // user cancelled the dialog
      } else {
        alert('Explorer 選択に失敗しました: ' + (j && j.message ? j.message : 'unknown'));
      }
    }catch(e){ alert('Explorer 選択に失敗しました: ' + e); }
  });
}
