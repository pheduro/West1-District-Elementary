/* CID Monitoring Checklist: school-head forms (per-school code) and PSDS summary (PIN).
   All data lives in a private Supabase database; the browser can only call the code/PIN-checked functions. */
(function () {
  'use strict';
  const SUPABASE_URL = 'https://didbinzzgbydutxwptwx.supabase.co';
  const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRpZGJpbnp6Z2J5ZHV0eHdwdHd4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTMxNTk2NDksImV4cCI6MjEwODczNTY0OX0.yLeweCo90AouQ4WKJuaaOkF6mEHpoyPZMAO2AIAlJ6M'; // public anon key, safe in the browser: it can only call the code/PIN-checked functions
  const PSDS_NAME = 'MARY JEAN B. CODIÑERA DevEdD';
  const DEFAULT_MONITOR = 'MARY JEAN B. CODIÑERA';
  const TERMS = ['TERM 1', 'TERM 2', 'TERM 3', 'TERM 4'];

  const S = { forms: null, order: [], code: null, school: null, responses: [], pin: null, psds: null, view: 'login', ed: null, detail: null, busy: false, filter: { term: '', checklist: '' }, flash: '' };
  let root = null, mounted = false;

  const h = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const fmtDate = d => d ? window.CidXlsx.fmtDate(d) : 'No date';
  const fmtStamp = t => { try { return new Date(t).toLocaleString('en-PH', { dateStyle: 'medium', timeStyle: 'short' }); } catch (e) { return ''; } };
  const store = { get: k => { try { return sessionStorage.getItem(k); } catch (e) { return null; } }, set: (k, v) => { try { sessionStorage.setItem(k, v); } catch (e) { } }, del: k => { try { sessionStorage.removeItem(k); } catch (e) { } } };

  async function rpc(fn, args) {
    let res;
    try {
      res = await fetch(SUPABASE_URL + '/rest/v1/rpc/' + fn, { method: 'POST', headers: { apikey: SUPABASE_KEY, Authorization: 'Bearer ' + SUPABASE_KEY, 'Content-Type': 'application/json' }, body: JSON.stringify(args) });
    } catch (e) { throw new Error('Could not reach the server. Please check your internet connection and try again.'); }
    let j = null; try { j = await res.json(); } catch (e) { }
    if (!res.ok) throw new Error((j && j.message) || 'Something went wrong (' + res.status + '). Please try again.');
    if (j && j.error) throw new Error(j.error);
    return j;
  }

  /* ---------- helpers ---------- */
  const formList = () => S.order.map(k => S.forms[k]);
  const totalItems = f => f.sections.reduce((n, s) => n + s.items.length, 0);
  function tally(f, answers) {
    let y = 0, n = 0, t = 0;
    f.sections.forEach((s, si) => s.items.forEach((it, ii) => { t++; const a = (answers || {})[si + '.' + ii]; if (a && a.a === 'Y') y++; else if (a && a.a === 'N') n++; }));
    return { yes: y, no: n, blank: t - y - n, total: t };
  }
  function secTally(f, si, answers) {
    let y = 0, n = 0;
    f.sections[si].items.forEach((it, ii) => { const a = (answers || {})[si + '.' + ii]; if (a && a.a === 'Y') y++; else if (a && a.a === 'N') n++; });
    return { yes: y, no: n };
  }
  const schoolName = id => { const s = (S.psds && S.psds.schools || []).find(x => x.id === id); return s ? s.short_name : id; };
  const schoolInfo = id => (S.psds && S.psds.schools || []).find(x => x.id === id) || (S.school && S.school.school_id === id ? S.school : { full_name: id, head: '' });
  const byDateDesc = (a, b) => String(b.visit_date || '').localeCompare(String(a.visit_date || '')) || String(b.updated_at).localeCompare(String(a.updated_at));

  function toast(msg, bad) {
    S.flash = msg ? { msg, bad: !!bad } : '';
    const el = root.querySelector('.cid-flash'); if (el) { el.outerHTML = flashHtml(); }
    if (msg) setTimeout(() => { if (S.flash && S.flash.msg === msg) { S.flash = ''; const e2 = root.querySelector('.cid-flash'); if (e2) e2.outerHTML = flashHtml(); } }, bad ? 7000 : 4000);
  }
  const flashHtml = () => '<div class="cid-flash" role="status" aria-live="polite">' + (S.flash ? '<div class="cid-toast' + (S.flash.bad ? ' bad' : '') + '">' + h(S.flash.msg) + '</div>' : '') + '</div>';

  /* ---------- downloads ---------- */
  async function download(formKey, items, mode, fname, extra) {
    const f = S.forms[formKey];
    toast('Preparing the Excel file…');
    try {
      const r = await fetch(f.template);
      if (!r.ok) throw new Error('The form template could not be loaded.');
      const bytes = await r.arrayBuffer();
      const out = await window.CidXlsx.build(window.JSZip, bytes, f, items, Object.assign({ mode, psds: PSDS_NAME }, extra || {}));
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = fname; document.body.appendChild(a); a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
      toast('Downloaded ' + fname);
    } catch (e) { toast(e.message || 'The Excel file could not be created.', true); }
  }
  const safe = s => String(s).replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '');
  function downloadResponse(r) {
    const f = S.forms[r.checklist]; const sc = schoolInfo(r.school_id);
    const sheet = f.sheets[r.school_id];
    const fname = f.key.toUpperCase() + '_' + safe(sc.full_name || r.school_id) + '_' + safe(r.term || 'TERM') + (r.visit_date ? '_' + r.visit_date : '') + '.xlsx';
    return download(f.key, [{ sheet, response: r, school: sc, psds: PSDS_NAME }], 'single', fname);
  }
  function downloadConsolidated(formKey) {
    const f = S.forms[formKey]; const term = S.filter.term;
    const latest = {};
    S.psds.responses.filter(r => r.checklist === formKey && (!term || r.term === term)).sort(byDateDesc).forEach(r => { if (!latest[r.school_id]) latest[r.school_id] = r; });
    const items = Object.keys(latest).map(id => ({ sheet: f.sheets[id], response: latest[id], school: schoolInfo(id), psds: PSDS_NAME })).filter(x => x.sheet);
    const blank = {}; Object.keys(f.sheets).forEach(id => { blank[f.sheets[id]] = schoolInfo(id); });
    if (!items.length) { toast('No responses to include yet.', true); return; }
    return download(formKey, items, 'all', f.key.toUpperCase() + '_all-schools' + (term ? '_' + safe(term) : '') + '.xlsx', { blankSheets: blank });
  }

  /* ---------- views ---------- */
  function render() {
    if (!root) return;
    let body = '';
    if (S.view === 'login') body = viewLogin();
    else if (S.view === 'school') body = viewSchool();
    else if (S.view === 'form') body = viewForm();
    else if (S.view === 'psds') body = viewPsds();
    else if (S.view === 'detail') body = viewDetail();
    root.innerHTML = flashHtml() + body;
    window.scrollTo({ top: 0 });
  }

  function viewLogin() {
    return '<div class="card cid-intro"><h2>CID Monitoring Checklists</h2><p class="sub">School Heads answer each checklist every time the PSDS visits the school. Your school code opens only your own school, and no other school can see your responses. The PSDS summary needs a PIN.</p></div>' +
      '<div class="cid-two">' +
      '<form class="card" data-form="school-login" autocomplete="off"><h3>School Head</h3><p class="sub">Enter the code given to your school.</p>' +
      '<label class="cid-lab">School code<input class="cid-in" name="code" required maxlength="20" placeholder="XXXX-XXXX" autocapitalize="characters" spellcheck="false" autocomplete="off"></label>' +
      '<button class="btn" type="submit">Open my school</button></form>' +
      '<form class="card" data-form="psds-login" autocomplete="off"><h3>PSDS summary</h3><p class="sub">District view of all schools. Enter the PSDS PIN.</p>' +
      '<label class="cid-lab">PIN<input class="cid-in" name="pin" type="password" required maxlength="20" inputmode="numeric" autocomplete="off" placeholder="••••••"></label>' +
      '<button class="btn" type="submit">Open summary</button></form></div>';
  }

  function viewSchool() {
    const sc = S.school;
    let out = '<div class="cid-bar"><div><h2>' + h(sc.full_name) + '</h2><p class="sub">School Head: ' + h(sc.head || '') + ' · Only your school can see these responses.</p></div><button class="btn ghost" data-act="logout" type="button">Log out</button></div>';
    formList().forEach(f => {
      const list = S.responses.filter(r => r.checklist === f.key).sort(byDateDesc);
      out += '<section class="card cid-cl"><div class="cid-clh"><div><h3>' + h(f.short) + '</h3><p class="sub">' + totalItems(f) + ' indicators in ' + f.sections.length + ' sections · answer once for every PSDS visit</p></div><button class="btn" data-act="new" data-cl="' + f.key + '" type="button">New response</button></div>';
      if (!list.length) out += '<p class="cid-empty">No response yet. Choose “New response” when the PSDS visits.</p>';
      else {
        out += '<div class="tw"><table class="cid-tbl"><thead><tr><th class="l">Visit date</th><th>Term</th><th>Yes</th><th>No</th><th>Not answered</th><th class="l">Last saved</th><th></th></tr></thead><tbody>';
        list.forEach(r => {
          const t = tally(f, r.answers);
          out += '<tr><td class="l"><b>' + h(fmtDate(r.visit_date)) + '</b>' + (r.note ? '<div class="cid-note">' + h(r.note) + '</div>' : '') + '</td><td>' + h(r.term || '') + '</td><td>' + t.yes + '</td><td>' + t.no + '</td><td>' + t.blank + '</td><td class="l">' + h(fmtStamp(r.updated_at)) + '</td>' +
            '<td class="cid-acts"><button class="btn sm" data-act="edit" data-id="' + r.id + '" type="button">Edit</button><button class="btn sm ghost" data-act="dl" data-id="' + r.id + '" type="button">Excel</button><button class="btn sm ghost danger" data-act="del" data-id="' + r.id + '" type="button">Delete</button></td></tr>';
        });
        out += '</tbody></table></div>';
      }
      out += '</section>';
    });
    return out;
  }

  function viewForm() {
    const ed = S.ed, f = S.forms[ed.checklist], sc = S.school;
    let out = '<div class="cid-bar"><div><div class="eyebrow-s">' + h(sc.full_name) + '</div><h2>' + h(f.title) + '</h2></div><button class="btn ghost" data-act="back" type="button">← Back</button></div>';
    out += '<div class="card cid-head">' +
      '<label class="cid-lab">Date of PSDS visit<input class="cid-in" type="date" data-f="visit_date" value="' + h(ed.visit_date || '') + '"></label>' +
      '<label class="cid-lab">Quarter / Term<select class="cid-in" data-f="term">' + TERMS.map(t => '<option' + (ed.term === t ? ' selected' : '') + '>' + t + '</option>').join('') + '</select></label>' +
      '<label class="cid-lab cid-wide">Monitored by<input class="cid-in" data-f="monitored_by" value="' + h(ed.monitored_by || '') + '" maxlength="120"></label></div>';
    f.sections.forEach((sec, si) => {
      out += '<section class="card cid-sec"><div class="cid-sech"><h3>' + h(sec.name) + '</h3><span class="pill" id="cnt-' + si + '"></span></div><div class="cid-items">';
      sec.items.forEach((it, ii) => {
        const k = si + '.' + ii, a = ed.answers[k] || {};
        out += '<div class="cid-item" data-k="' + k + '"><div class="cid-q">' + h(it.text) + '</div>' +
          '<div class="cid-yn" role="group" aria-label="Answer"><button type="button" data-act="ans" data-k="' + k + '" data-v="Y" aria-pressed="' + (a.a === 'Y') + '">Yes</button><button type="button" data-act="ans" data-k="' + k + '" data-v="N" aria-pressed="' + (a.a === 'N') + '">No</button></div>' +
          '<input class="cid-rem" data-rem="' + k + '" placeholder="Remarks / evidence" value="' + h(a.r || '') + '" maxlength="500"></div>';
      });
      out += '</div></section>';
    });
    out += '<section class="card cid-sum"><h3>Summary of findings</h3>' +
      '<label class="cid-lab">Strengths<textarea class="cid-in" rows="3" data-f="strengths" maxlength="5000">' + h(ed.strengths) + '</textarea></label>' +
      '<label class="cid-lab">Areas for improvement<textarea class="cid-in" rows="3" data-f="areas" maxlength="5000">' + h(ed.areas) + '</textarea></label>' +
      '<label class="cid-lab">Agreed next action (next quarter)<textarea class="cid-in" rows="3" data-f="next_action" maxlength="5000">' + h(ed.next_action) + '</textarea></label></section>';
    out += '<div class="cid-actbar"><span class="cid-prog" id="cid-prog"></span><button class="btn ghost" data-act="dl-draft" type="button">Download Excel</button><button class="btn" data-act="save" type="button">' + (ed.id ? 'Save changes' : 'Save response') + '</button></div>';
    setTimeout(refreshCounts, 0);
    return out;
  }
  function refreshCounts() {
    const ed = S.ed; if (!ed || S.view !== 'form') return;
    const f = S.forms[ed.checklist];
    f.sections.forEach((s, si) => { const c = secTally(f, si, ed.answers), el = document.getElementById('cnt-' + si); if (el) el.textContent = c.yes + ' Yes · ' + c.no + ' No'; });
    const t = tally(f, ed.answers), p = document.getElementById('cid-prog'); if (p) p.textContent = (t.yes + t.no) + ' of ' + t.total + ' answered';
  }

  function viewPsds() {
    const P = S.psds, term = S.filter.term;
    let out = '<div class="cid-bar"><div><h2>PSDS summary</h2><p class="sub">All schools, latest visit per checklist. Only the PSDS can open this page.</p></div><button class="btn ghost" data-act="logout" type="button">Log out</button></div>';
    const terms = [...new Set(P.responses.map(r => r.term).filter(Boolean))].sort();
    out += '<div class="card cid-filters"><label class="cid-lab">Term<select class="cid-in" data-filter="term"><option value="">All terms</option>' + terms.map(t => '<option' + (t === term ? ' selected' : '') + '>' + h(t) + '</option>').join('') + '</select></label></div>';
    // matrix
    out += '<section class="card"><h3>Responses by school</h3><p class="sub">Each cell shows the latest visit: date, then Yes and No counts. Select a cell to read the answers.</p><div class="tw"><table class="cid-tbl cid-matrix"><thead><tr><th class="l">School</th>' + formList().map(f => '<th title="' + h(f.short) + '">' + h(f.abbr || f.short) + '</th>').join('') + '</tr></thead><tbody>';
    P.schools.forEach(sc => {
      out += '<tr><td class="l"><b>' + h(sc.short_name) + '</b><div class="cid-note">' + h(sc.head || '') + '</div></td>';
      formList().forEach(f => {
        const rs = P.responses.filter(r => r.school_id === sc.id && r.checklist === f.key && (!term || r.term === term)).sort(byDateDesc);
        if (!rs.length) out += '<td class="cid-none">No response</td>';
        else { const r = rs[0], t = tally(f, r.answers); out += '<td><button class="cid-cell" data-act="detail" data-id="' + r.id + '" type="button"><b>' + h(fmtDate(r.visit_date)) + '</b><span>' + t.yes + ' Yes · ' + t.no + ' No' + (t.blank ? ' · ' + t.blank + ' blank' : '') + '</span>' + (rs.length > 1 ? '<span class="cid-note">' + rs.length + ' visits</span>' : '') + '</button></td>'; }
      });
      out += '</tr>';
    });
    out += '</tbody></table></div></section>';
    // gaps and downloads per checklist
    formList().forEach(f => {
      const latest = {};
      P.responses.filter(r => r.checklist === f.key && (!term || r.term === term)).sort(byDateDesc).forEach(r => { if (!latest[r.school_id]) latest[r.school_id] = r; });
      const ids = Object.keys(latest);
      out += '<section class="card cid-cl"><div class="cid-clh"><div><h3>' + h(f.short) + '</h3><p class="sub">' + ids.length + ' of ' + P.schools.length + ' schools have a response' + (term ? ' for ' + h(term) : '') + '.</p></div><button class="btn ghost" data-act="dl-all" data-cl="' + f.key + '" type="button"' + (ids.length ? '' : ' disabled') + '>Download all schools (Excel)</button></div>';
      if (ids.length) {
        const rows = [];
        f.sections.forEach((sec, si) => sec.items.forEach((it, ii) => {
          const k = si + '.' + ii; const nos = ids.filter(id => (latest[id].answers[k] || {}).a === 'N').map(schoolName);
          if (nos.length) rows.push({ text: it.text, sec: sec.name, nos });
        }));
        const totalY = ids.map(id => tally(f, latest[id].answers));
        out += '<div class="tw"><table class="cid-tbl"><thead><tr><th class="l">School</th><th>Yes</th><th>No</th><th>Not answered</th><th>% Yes</th></tr></thead><tbody>' +
          ids.sort((a, b) => schoolName(a).localeCompare(schoolName(b))).map(id => { const t = tally(f, latest[id].answers), ans = t.yes + t.no; return '<tr><td class="l">' + h(schoolName(id)) + '</td><td>' + t.yes + '</td><td>' + t.no + '</td><td>' + t.blank + '</td><td>' + (ans ? Math.round(t.yes / ans * 100) + '%' : '–') + '</td></tr>'; }).join('') + '</tbody></table></div>';
        out += '<details class="cid-gaps"><summary>Indicators answered No (' + rows.length + ')</summary>' + (rows.length ? '<ul class="l">' + rows.map(r => '<li>' + h(r.text) + ' <span class="cid-note">' + h(r.nos.join(', ')) + '</span></li>').join('') + '</ul>' : '<p class="cid-empty">No indicator was answered No.</p>') + '</details>';
      }
      out += '</section>';
    });
    return out;
  }

  function viewDetail() {
    const r = S.detail, f = S.forms[r.checklist], sc = schoolInfo(r.school_id);
    const all = S.psds.responses.filter(x => x.school_id === r.school_id && x.checklist === r.checklist).sort(byDateDesc);
    let out = '<div class="cid-bar"><div><div class="eyebrow-s">' + h(sc.full_name) + '</div><h2>' + h(f.title) + '</h2></div><button class="btn ghost" data-act="back" type="button">← Back to summary</button></div>';
    out += '<div class="card cid-head"><label class="cid-lab">Visit<select class="cid-in" data-detail-pick="1">' + all.map(x => '<option value="' + x.id + '"' + (x.id === r.id ? ' selected' : '') + '>' + h(fmtDate(x.visit_date)) + ' · ' + h(x.term || '') + '</option>').join('') + '</select></label>' +
      '<div class="cid-kv"><span>Term</span><b>' + h(r.term || '') + '</b></div><div class="cid-kv"><span>Monitored by</span><b>' + h(r.monitored_by || '') + '</b></div><div class="cid-kv"><span>Last saved</span><b>' + h(fmtStamp(r.updated_at)) + '</b></div><button class="btn" data-act="dl-detail" type="button">Download Excel</button></div>';
    f.sections.forEach((sec, si) => {
      const c = secTally(f, si, r.answers);
      out += '<section class="card cid-sec"><div class="cid-sech"><h3>' + h(sec.name) + '</h3><span class="pill">' + c.yes + ' Yes · ' + c.no + ' No</span></div><div class="cid-items">';
      sec.items.forEach((it, ii) => { const a = r.answers[si + '.' + ii] || {}; out += '<div class="cid-item ro"><div class="cid-q">' + h(it.text) + '</div><div class="cid-ans ' + (a.a === 'Y' ? 'y' : a.a === 'N' ? 'n' : '') + '">' + (a.a === 'Y' ? 'Yes' : a.a === 'N' ? 'No' : 'Not answered') + '</div><div class="cid-remro">' + h(a.r || '') + '</div></div>'; });
      out += '</div></section>';
    });
    out += '<section class="card cid-sum"><h3>Summary of findings</h3><h4>Strengths</h4><p class="cid-pre">' + h(r.strengths || '–') + '</p><h4>Areas for improvement</h4><p class="cid-pre">' + h(r.areas || '–') + '</p><h4>Agreed next action</h4><p class="cid-pre">' + h(r.next_action || '–') + '</p></section>';
    return out;
  }

  /* ---------- actions ---------- */
  async function schoolLogin(code) {
    S.busy = true;
    try {
      const info = await rpc('cid_login', { p_code: code });
      const list = await rpc('cid_list', { p_code: code });
      S.flash = ''; S.code = code; S.school = info; S.responses = list.responses; S.view = 'school'; store.set('cid.code', code); store.del('cid.pin');
    } finally { S.busy = false; }
  }
  async function psdsLogin(pin) {
    const data = await rpc('cid_psds_all', { p_pin: pin });
    S.flash = ''; S.pin = pin; S.psds = data; S.view = 'psds'; store.set('cid.pin', pin); store.del('cid.code');
  }
  function logout() { S.code = S.pin = S.school = S.psds = S.ed = S.detail = null; S.responses = []; S.view = 'login'; store.del('cid.code'); store.del('cid.pin'); render(); }

  function newEd(cl, r) {
    return r ? { id: r.id, checklist: r.checklist, visit_date: r.visit_date || '', term: r.term || 'TERM 1', monitored_by: r.monitored_by || DEFAULT_MONITOR, answers: JSON.parse(JSON.stringify(r.answers || {})), strengths: r.strengths || '', areas: r.areas || '', next_action: r.next_action || '', dirty: false }
      : { id: null, checklist: cl, visit_date: '', term: 'TERM 1', monitored_by: DEFAULT_MONITOR, answers: {}, strengths: '', areas: '', next_action: '', dirty: false };
  }
  async function save() {
    const ed = S.ed, f = S.forms[ed.checklist];
    if (!ed.visit_date) { toast('Please enter the date of the PSDS visit.', true); const el = root.querySelector('[data-f=visit_date]'); if (el) el.focus(); return; }
    const t = tally(f, ed.answers);
    if (t.blank && !confirm(t.blank + ' indicator(s) are not answered yet. Save anyway? You can edit this response later.')) return;
    const clean = {}; Object.keys(ed.answers).forEach(k => { const a = ed.answers[k]; if (a && (a.a || a.r)) clean[k] = { a: a.a || null, r: a.r || '' }; });
    try {
      const res = await rpc('cid_save', { p_code: S.code, p_id: ed.id, p_checklist: ed.checklist, p_visit_date: ed.visit_date, p_term: ed.term, p_monitored_by: ed.monitored_by, p_answers: clean, p_strengths: ed.strengths, p_areas: ed.areas, p_next: ed.next_action });
      const row = res.response; const i = S.responses.findIndex(x => x.id === row.id);
      if (i >= 0) S.responses[i] = row; else S.responses.unshift(row);
      ed.id = row.id; ed.dirty = false; toast('Saved. You can come back and edit this response any time.');
      const b = root.querySelector('[data-act=save]'); if (b) b.textContent = 'Save changes';
    } catch (e) { toast(e.message, true); }
  }

  function onClick(e) {
    const b = e.target.closest('[data-act]'); if (!b || !root.contains(b)) return;
    const a = b.dataset.act;
    if (a === 'logout') return logout();
    if (a === 'new') { S.ed = newEd(b.dataset.cl); S.view = 'form'; return render(); }
    if (a === 'edit') { const r = S.responses.find(x => x.id === b.dataset.id); S.ed = newEd(null, r); S.view = 'form'; return render(); }
    if (a === 'dl') { const r = S.responses.find(x => x.id === b.dataset.id); return downloadResponse(r); }
    if (a === 'del') {
      const r = S.responses.find(x => x.id === b.dataset.id);
      if (!confirm('Delete this response (' + fmtDate(r.visit_date) + ')? This cannot be undone.')) return;
      rpc('cid_delete', { p_code: S.code, p_id: r.id }).then(() => { S.responses = S.responses.filter(x => x.id !== r.id); render(); toast('Response deleted.'); }).catch(err => toast(err.message, true));
      return;
    }
    if (a === 'back') {
      if (S.view === 'form' && S.ed.dirty && !confirm('You have unsaved changes. Leave without saving?')) return;
      S.view = S.psds && S.pin ? 'psds' : 'school'; S.ed = null; S.detail = null; return render();
    }
    if (a === 'ans') {
      const k = b.dataset.k, v = b.dataset.v, ed = S.ed; const cur = ed.answers[k] || {};
      ed.answers[k] = { a: cur.a === v ? null : v, r: cur.r || '' }; ed.dirty = true;
      const row = root.querySelector('.cid-item[data-k="' + k + '"]');
      row.querySelectorAll('[data-act=ans]').forEach(x => x.setAttribute('aria-pressed', String(ed.answers[k].a === x.dataset.v)));
      refreshCounts(); return;
    }
    if (a === 'save') return save();
    if (a === 'dl-draft') {
      const ed = S.ed, sc = S.school;
      return downloadResponse({ id: ed.id, school_id: sc.school_id, checklist: ed.checklist, visit_date: ed.visit_date, term: ed.term, monitored_by: ed.monitored_by, answers: ed.answers, strengths: ed.strengths, areas: ed.areas, next_action: ed.next_action });
    }
    if (a === 'detail') { S.detail = S.psds.responses.find(x => x.id === b.dataset.id); S.view = 'detail'; return render(); }
    if (a === 'dl-detail') return downloadResponse(S.detail);
    if (a === 'dl-all') return downloadConsolidated(b.dataset.cl);
  }
  function onInput(e) {
    const t = e.target;
    if (t.dataset.f && S.ed) { S.ed[t.dataset.f] = t.value; S.ed.dirty = true; }
    else if (t.dataset.rem && S.ed) { const k = t.dataset.rem; const cur = S.ed.answers[k] || {}; S.ed.answers[k] = { a: cur.a || null, r: t.value }; S.ed.dirty = true; }
  }
  function onChange(e) {
    const t = e.target;
    if (t.dataset.filter === 'term') { S.filter.term = t.value; render(); }
    else if (t.dataset.detailPick) { S.detail = S.psds.responses.find(x => x.id === t.value); render(); }
    else if (t.dataset.f && S.ed) { S.ed[t.dataset.f] = t.value; S.ed.dirty = true; }
  }
  async function onSubmit(e) {
    const f = e.target.closest('form[data-form]'); if (!f) return;
    e.preventDefault();
    const btn = f.querySelector('button[type=submit]'); btn.disabled = true; const label = btn.textContent; btn.textContent = 'Checking…';
    try {
      if (f.dataset.form === 'school-login') await schoolLogin(f.code.value.trim());
      else await psdsLogin(f.pin.value.trim());
      render();
    } catch (err) { btn.disabled = false; btn.textContent = label; toast(err.message, true); }
  }

  async function mount() {
    if (mounted) return; mounted = true;
    root = document.getElementById('pane-cid');
    root.classList.add('cid');
    root.innerHTML = '<p class="cid-empty">Loading…</p>';
    try { const r = await fetch('forms.json'); S.forms = await r.json(); S.order = Object.keys(S.forms); } catch (e) { root.innerHTML = '<p class="cid-empty">The checklists could not be loaded. Please refresh the page.</p>'; mounted = false; return; }
    root.addEventListener('click', onClick); root.addEventListener('input', onInput); root.addEventListener('change', onChange); root.addEventListener('submit', onSubmit);
    window.addEventListener('beforeunload', ev => { if (S.view === 'form' && S.ed && S.ed.dirty) { ev.preventDefault(); ev.returnValue = ''; } });
    const code = store.get('cid.code'), pin = store.get('cid.pin');
    try { if (code) await schoolLogin(code); else if (pin) await psdsLogin(pin); } catch (e) { store.del('cid.code'); store.del('cid.pin'); }
    render();
  }

  window.CidApp = { mount };
})();
