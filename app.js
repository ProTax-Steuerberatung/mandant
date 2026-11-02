/* ProTax Mandant – Oberfläche, Anmeldung, Speicherung (Demo / SharePoint) */
'use strict';
const VERSION = 'ProTax Mandant v.1.2';
const C = window.CONFIG || {}, SC = window.SCHEMA, LG = window.LOGIC;
const LIST_NAME = 'MA_Aufnahmen', LIB_NAME = 'MA_Unterlagen';

/* ================= Helfer ================= */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const nowIso = () => new Date().toISOString();
const fmtDT = iso => iso ? new Date(iso).toLocaleString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '';
const fmtT = iso => iso ? new Date(iso).toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }) : '';
const safeName = s => String(s || '').replace(/[\\/:*?"<>|#%{}~&]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80);
const statusLabel = k => (SC.STATUS.find(s => s.k === k) || (k === 'abgebrochen' ? SC.STATUS_ABGEBROCHEN : { l: k })).l;
const initials = n => String(n || '?').split(/[\s.@]+/).filter(Boolean).slice(0, 2).map(x => x[0].toUpperCase()).join('');
function toast(msg, kind = '') {
  const el = document.createElement('div'); el.className = 't ' + kind; el.textContent = msg;
  $('#toast').appendChild(el); setTimeout(() => el.remove(), kind === 'err' ? 7000 : 3500);
}
const I = {
  back: '<svg class="icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true"><path d="M10 3 5 8l5 5"/></svg>',
  mail: '<svg class="icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><rect x="2" y="3.5" width="12" height="9" rx="1.5"/><path d="m2.5 4.5 5.5 4 5.5-4"/></svg>',
  print: '<svg class="icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M4 6V2h8v4M4 12H2.5V6.5h11V12H12"/><rect x="4" y="10" width="8" height="4"/></svg>',
  plus: '<svg class="icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.7" aria-hidden="true"><path d="M8 3v10M3 8h10"/></svg>',
  up: '<svg class="icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M8 11V3M5 6l3-3 3 3M3 12v1.5h10V12"/></svg>',
  file: '<svg class="icon" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.4" aria-hidden="true"><path d="M4 1.5h5l3 3v10H4z"/><path d="M9 1.5v3h3"/></svg>'
};

/* ================= Speicher: Demo ================= */
const Demo = {
  key: 'protax-mandant-demo-v1', mem: null, blobs: {},
  read() { try { const s = localStorage.getItem(this.key); return s ? JSON.parse(s) : null; } catch (e) { return this.mem; } },
  write(list) { this.mem = list; try { localStorage.setItem(this.key, JSON.stringify(list)); } catch (e) { } },
  async init() { if (!this.read()) this.write(demoSeed()); return true; },
  async loadAll() { return structuredClone(this.read() || []); },
  async create(rec) { const all = this.read() || []; rec.rev = 1; all.push(structuredClone(rec)); this.write(all); return rec; },
  async save(rec, force) {
    const all = this.read() || [], i = all.findIndex(r => r.id === rec.id);
    if (i >= 0 && !force && (all[i].rev || 0) !== (rec.rev || 0)) { const e = new Error('Konflikt'); e.conflict = true; throw e; }
    const next = (i >= 0 ? (all[i].rev || 0) : 0) + 1;
    const copy = structuredClone({ ...rec, rev: next });
    if (i >= 0) all[i] = copy; else all.push(copy);
    this.write(all); rec.rev = next; return rec;
  },
  async reload(rec) { return structuredClone((this.read() || []).find(r => r.id === rec.id)); },
  async upload(rec, docKey, file) { const id = uid(); this.blobs[id] = URL.createObjectURL(file); return { id, name: file.name, size: file.size, url: '' }; },
  href(f) { return this.blobs[f.id] || f.url || ''; },
  async mailToken() { return 'demo'; },
  async createDraft() { return null; },
  reset() { try { localStorage.removeItem(this.key); } catch (e) { } this.mem = null; }
};

/* ================= Anmeldung (Microsoft 365) ================= */
const Auth = {
  pca: null, account: null, scopes: ['User.Read', 'Sites.ReadWrite.All', 'Sites.Manage.All'],
  redirect() { return C.redirectUri || location.href.split(/[?#]/)[0].replace(/index\.html$/, ''); },
  async init() {
    if (!window.msal) throw new Error('Anmeldebibliothek (lib/msal-browser.min.js) fehlt');
    this.pca = new msal.PublicClientApplication({
      auth: { clientId: C.clientId, authority: 'https://login.microsoftonline.com/' + C.tenantId, redirectUri: this.redirect(), postLogoutRedirectUri: this.redirect() },
      cache: { cacheLocation: 'localStorage' }
    });
    await this.pca.initialize();
    const r = await this.pca.handleRedirectPromise();
    if (r && r.account) this.pca.setActiveAccount(r.account);
    this.account = this.pca.getActiveAccount() || this.pca.getAllAccounts()[0] || null;
    if (this.account) this.pca.setActiveAccount(this.account);
    return this.account;
  },
  login() { return this.pca.loginRedirect({ scopes: this.scopes, prompt: 'select_account' }); },
  logout() { return this.pca.logoutRedirect({ account: this.account }); },
  async token(scopes) {
    const sc = scopes || this.scopes;
    try { return (await this.pca.acquireTokenSilent({ scopes: sc, account: this.account })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) {
        if (scopes) return (await this.pca.acquireTokenPopup({ scopes: sc, account: this.account })).accessToken;
        await this.pca.acquireTokenRedirect({ scopes: sc });
      }
      throw e;
    }
  }
};

/* ================= Speicher: SharePoint (Microsoft Graph) ================= */
const SP = {
  siteId: null, listId: null, libId: null, driveId: null, webUrl: '',
  async g(path, opt = {}) {
    const url = path.startsWith('http') ? path : 'https://graph.microsoft.com/v1.0' + path;
    const headers = { Authorization: 'Bearer ' + (opt.token || await Auth.token()), ...(opt.headers || {}) };
    let body = opt.body;
    if (opt.json !== undefined) { headers['Content-Type'] = 'application/json'; body = JSON.stringify(opt.json); }
    const r = await fetch(url, { method: opt.method || 'GET', headers, body });
    if (!r.ok) {
      let m = 'HTTP ' + r.status;
      try { const j = await r.json(); m = (j.error && j.error.message) || m; } catch (e) { }
      const err = new Error(m); err.status = r.status; throw err;
    }
    return r.status === 204 ? null : r.json();
  },
  async lists() { return (await this.g(`/sites/${this.siteId}/lists?$select=id,displayName&$top=999`)).value; },
  async init() {
    const sp = C.sharepoint || {};
    const site = await this.g(`/sites/${sp.hostname}:${sp.sitePath}?$select=id,webUrl`);
    this.siteId = site.id; this.webUrl = site.webUrl;
    const ls = await this.lists();
    const a = ls.find(l => l.displayName === LIST_NAME), b = ls.find(l => l.displayName === LIB_NAME);
    if (!a || !b) return false;
    this.listId = a.id; this.libId = b.id;
    this.driveId = (await this.g(`/sites/${this.siteId}/lists/${b.id}/drive?$select=id`)).id;
    return true;
  },
  async setup() {
    const ls = await this.lists();
    if (!ls.find(l => l.displayName === LIST_NAME)) {
      await this.g(`/sites/${this.siteId}/lists`, {
        method: 'POST', json: {
          displayName: LIST_NAME, list: { template: 'genericList' }, columns: [
            { name: 'Typ', text: {} }, { name: 'Status', text: {} }, { name: 'Vollstaendigkeit', number: {} },
            { name: 'DatevNr', text: {} }, { name: 'Rev', number: {} },
            { name: 'Daten', text: { allowMultipleLines: true, linesForEditing: 6, textType: 'plain' } }]
        }
      });
    }
    if (!ls.find(l => l.displayName === LIB_NAME)) {
      await this.g(`/sites/${this.siteId}/lists`, { method: 'POST', json: { displayName: LIB_NAME, list: { template: 'documentLibrary' } } });
    }
    return this.init();
  },
  base() { return `/sites/${this.siteId}/lists/${this.listId}/items`; },
  fromItem(it) {
    let rec; try { rec = JSON.parse(it.fields.Daten || '{}'); } catch (e) { return null; }
    if (!rec || !rec.id) return null;
    rec.spId = it.id; rec.rev = Number(it.fields.Rev) || 0; return rec;
  },
  fields(rec, rev) {
    const copy = { ...rec }; delete copy.spId; delete copy.rev;
    const json = JSON.stringify(copy);
    if (json.length > 60000) throw new Error('Datensatz zu groß für SharePoint (über 60.000 Zeichen) – bitte Notizen kürzen');
    return { Title: rec.title, Typ: rec.typ, Status: rec.status, Vollstaendigkeit: rec.pct || 0, DatevNr: rec.datevNr || '', Rev: rev, Daten: json };
  },
  async loadAll() {
    let url = `${this.base()}?expand=fields(select=Title,Typ,Status,Vollstaendigkeit,Daten,DatevNr,Rev)&$top=500`;
    const out = [];
    while (url) { const j = await this.g(url); j.value.forEach(it => { const r = this.fromItem(it); if (r) out.push(r); }); url = j['@odata.nextLink']; }
    return out;
  },
  async create(rec) { const it = await this.g(this.base(), { method: 'POST', json: { fields: this.fields(rec, 1) } }); rec.spId = it.id; rec.rev = 1; return rec; },
  async save(rec, force) {
    const cur = await this.g(`${this.base()}/${rec.spId}?expand=fields(select=Rev)`);
    const r = Number(cur.fields.Rev) || 0;
    if (!force && r !== (rec.rev || 0)) { const e = new Error('Konflikt'); e.conflict = true; throw e; }
    await this.g(`${this.base()}/${rec.spId}/fields`, { method: 'PATCH', json: this.fields(rec, r + 1) });
    rec.rev = r + 1; return rec;
  },
  async reload(rec) { return this.fromItem(await this.g(`${this.base()}/${rec.spId}?expand=fields(select=Title,Typ,Status,Vollstaendigkeit,Daten,DatevNr,Rev)`)); },
  async upload(rec, docKey, file) {
    const path = [rec.folder, `${docKey}_${safeName(file.name) || 'datei'}`].map(encodeURIComponent).join('/');
    const it = await this.g(`/drives/${this.driveId}/root:/${path}:/content?@microsoft.graph.conflictBehavior=rename`, { method: 'PUT', body: file, headers: { 'Content-Type': file.type || 'application/octet-stream' } });
    return { id: it.id, name: it.name, size: it.size, url: it.webUrl };
  },
  href(f) { return f.url || ''; },
  MAIL: ['Mail.ReadWrite'],
  async mailToken() {
    try { return await Auth.token(this.MAIL); }
    catch (e) { const err = new Error('Keine Berechtigung für Outlook-Entwürfe. In Entra bei der App „Mandant ProTax“ die Berechtigung Mail.ReadWrite hinzufügen und die Administratorzustimmung erteilen.'); err.cause = e; throw err; }
  },
  async createDraft(token, m) {
    const msg = await this.g('/me/messages', { method: 'POST', token, json: {
      subject: m.subject, body: { contentType: 'HTML', content: m.html },
      toRecipients: m.to ? [{ emailAddress: { address: m.to } }] : [],
      attachments: [{ '@odata.type': '#microsoft.graph.fileAttachment', name: m.fileName, contentType: 'application/pdf', contentBytes: m.pdfB64 }]
    } });
    return msg.webLink;
  }
};

/* ================= Zustand ================= */
const S = {
  store: null, user: null, recs: [], view: 'boot', id: null, sec: null, filter: 'aktiv', q: '',
  saveTimer: null, saving: false, saveAgain: false, saveState: '', conflict: false, error: ''
};
const cur = () => S.recs.find(r => r.id === S.id);

/* ================= Start ================= */
(async function start() {
  bindEvents();
  try {
    if (C.demo) {
      S.store = Demo; S.user = { name: 'Demo-Benutzer', email: 'demo@protax.de' };
      await Demo.init();
    } else {
      if (!C.tenantId || !C.clientId) { S.view = 'error'; S.error = 'config.js ist noch nicht ausgefüllt (tenantId / clientId) oder demo steht noch auf true.'; return render(); }
      const acc = await Auth.init();
      if (!acc) { S.view = 'login'; return render(); }
      S.user = { name: acc.name || acc.username, email: acc.username };
      S.store = SP;
      S.view = 'boot'; render();
      const ok = await SP.init();
      if (!ok) { S.view = 'setup'; return render(); }
    }
    await loadAll();
    route();
  } catch (e) {
    console.error(e);
    S.view = 'error';
    S.error = e.status === 403 || e.status === 401 ? 'Keine Berechtigung für die SharePoint-Website. Bitte prüfen, ob Sie Mitglied der Website „Mandant“ sind.' : e.status === 404 ? 'SharePoint-Website nicht gefunden – Adresse in config.js prüfen.' : (e.message || String(e));
    render();
  }
})();

async function loadAll() { S.recs = await S.store.loadAll(); S.recs.forEach(r => { r.pct = LG.evaluate(r).pct; }); }
const PREVIEW = !!window.__PREVIEW;
function route() {
  const h = PREVIEW ? (S.hash || '') : location.hash;
  const m = h.match(/^#\/a\/([^/]+)(?:\/(\w+))?/);
  if (m && S.recs.find(r => r.id === m[1])) { S.view = 'detail'; S.id = m[1]; S.sec = m[2] || S.sec; }
  else { S.view = 'liste'; S.id = null; }
  render();
}
function go(hash) { if (PREVIEW) { S.hash = hash; route(); window.scrollTo(0, 0); return; } if (location.hash === hash) route(); else location.hash = hash; }
window.addEventListener('hashchange', () => { if (S.store) route(); });
window.addEventListener('beforeunload', e => { if (S.saveTimer || S.saving) { e.preventDefault(); e.returnValue = ''; } });

/* ================= Rahmen ================= */
const IC = {
  liste: '<path d="M8 6h12M8 12h12M8 18h12"/><circle cx="4" cy="6" r="1"/><circle cx="4" cy="12" r="1"/><circle cx="4" cy="18" r="1"/>',
  mandant: '<rect x="3" y="5.5" width="18" height="13" rx="2"/><path d="m3.5 7 8.5 6 8.5-6"/>',
  pruef: '<path d="M9 4h6l1 2h3v14H5V6h3z"/><path d="m9 13 2 2 4-4"/>',
  check: '<circle cx="12" cy="12" r="8.5"/><path d="m8.5 12.5 2.3 2.3 4.7-5"/>',
  datev: '<ellipse cx="12" cy="6" rx="7" ry="2.5"/><path d="M5 6v12c0 1.4 3.1 2.5 7 2.5s7-1.1 7-2.5V6M5 12c0 1.4 3.1 2.5 7 2.5s7-1.1 7-2.5"/>',
  archiv: '<rect x="3.5" y="4.5" width="17" height="4" rx="1"/><path d="M5 8.5v10a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-10M10 12.5h4"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  refresh: '<path d="M19 12a7 7 0 1 1-2.05-4.95"/><path d="M19 5v4h-4"/>',
  logout: '<path d="M14 5h4a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2h-4"/><path d="M10 16l-4-4 4-4M6 12h9"/>',
  user: '<circle cx="12" cy="8.5" r="3.5"/><path d="M5 19.5c.9-3.4 3.7-5.5 7-5.5s6.1 2.1 7 5.5"/>',
  lock: '<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>'
};
const ico = k => `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${IC[k]}</svg>`;
const NAV = [['aktiv', 'In Arbeit', 'liste'], ['mandant', 'Beim Mandanten', 'mandant'], ['eingereicht', 'Zur Prüfung', 'pruef'], ['geprueft', 'Geprüft', 'check'], ['angelegt', 'In DATEV angelegt', 'datev'], ['abgebrochen', 'Abgebrochen', 'archiv'], ['alle', 'Alle Aufnahmen', 'liste']];
const navLabel = k => (NAV.find(n => n[0] === k) || [0, 'Aufnahmen'])[1];
function sideHtml() {
  const n = k => S.recs.filter(r => k === 'aktiv' ? AKTIV.includes(r.status) : k === 'alle' ? true : r.status === k).length;
  const cur = S.view === 'liste' ? S.filter : null;
  return `<div class="user"><div class="avatar">${esc(initials(S.user.name))}</div><div><div class="n">${esc(S.user.name)}</div><div class="r">Mandantenaufnahme</div></div></div>
    <button class="nav nav-new" data-act="new">${ico('plus')}<span>Neue Aufnahme</span></button>
    <div class="side-t">Aufnahmen</div>
    ${NAV.map(([k, l, i]) => { const c = n(k); return `<button class="nav ${cur === k ? 'on' : ''}" data-filter="${k}" ${cur === k ? 'aria-current="page"' : ''}>${ico(i)}<span>${l}</span>${c ? `<span class="cnt ${k === 'eingereicht' ? 'hot' : ''}">${c}</span>` : ''}</button>`; }).join('')}
    <div class="ver">${esc(VERSION)}</div>`;
}
function shell(inner, crumb) {
  const K = C.kanzlei || {};
  document.body.classList.toggle('login-mode', !S.user || !['liste', 'detail'].includes(S.view));
  const who = S.user ? `${C.demo ? '<span class="badge b-warn">Demo</span><button class="btn sm" data-act="demo-reset">Demo zurücksetzen</button>' : ''}
    <button class="iconbtn" data-act="refresh" title="Daten neu laden" aria-label="Daten neu laden">${ico('refresh')}</button>
    ${C.demo ? '' : `<button class="iconbtn" data-act="logout" title="Abmelden" aria-label="Abmelden">${ico('logout')}</button>`}` : '';
  const side = S.user && ['liste', 'detail'].includes(S.view);
  return `<header class="topbar"><a class="brand" href="#/" data-act="home" aria-label="Startseite"><span class="logo"><img src="${esc(C.logoHell || 'icons/logo.png')}" alt="${esc(K.name || 'ProTax Steuerberatung')}"></span></a><div class="who">${who}</div></header>
  <div class="shell">${side ? `<nav class="side" aria-label="Hauptnavigation">${sideHtml()}</nav>` : ''}
  <main>${crumb !== undefined ? `<div class="crumb">${esc(K.name || 'ProTax')}<i></i>${esc(crumb)}</div>` : ''}${inner}</main></div>`;
}
function loginCard(title, sub, body) {
  return `<div class="login"><div class="login-ico">${ico('user')}</div><h1>${title}</h1><p class="login-sub">${sub}</p>${body}<p class="login-ver">${esc(VERSION)}</p></div>`;
}

function render() {
  const app = $('#app');
  const v = S.view, K = C.kanzlei || {};
  if (v === 'boot') app.innerHTML = shell('<div class="login"><div class="empty">Verbinde mit SharePoint …</div></div>');
  else if (v === 'login') {
    const ms = '<svg viewBox="0 0 21 21" aria-hidden="true"><rect x="1" y="1" width="9" height="9" fill="#f25022"/><rect x="11" y="1" width="9" height="9" fill="#7fba00"/><rect x="1" y="11" width="9" height="9" fill="#00a4ef"/><rect x="11" y="11" width="9" height="9" fill="#ffb900"/></svg>';
    const feat = [['liste', 'Neue Mandanten im Erstgespräch erfassen'], ['mandant', 'Fehlende Angaben beim Mandanten anfordern'], ['check', 'Prüfen, Stammblatt drucken, nach DATEV übergeben']];
    app.innerHTML = shell(loginCard('Willkommen bei Mandant', `Aufnahme neuer Mandanten für ${esc(K.name || 'die Kanzlei')}`,
      `<ul class="login-feat">${feat.map(([i, t]) => `<li>${ico(i)}<span>${t}</span></li>`).join('')}</ul>
      <button class="btn primary login-btn" data-act="login">${ms}<span>Mit Microsoft anmelden</span></button>
      <p class="login-note">${ico('lock')}Anmeldung mit dem Kanzlei-Konto. Die Daten bleiben im Microsoft 365 der Kanzlei.</p>`));
  }
  else if (v === 'setup') app.innerHTML = shell(loginCard('Einrichtung nicht abgeschlossen', `In der SharePoint-Website fehlen die Liste „${LIST_NAME}“ und/oder die Bibliothek „${LIB_NAME}“. Sie werden jetzt einmalig angelegt.`,
      '<button class="btn primary login-btn" data-act="setup">Listen jetzt anlegen</button>'));
  else if (v === 'error') app.innerHTML = shell(loginCard('Hinweis', esc(S.error), '<button class="btn login-btn" data-act="reload">Neu laden</button>'));
  else if (v === 'liste') app.innerHTML = shell(viewListe(), navLabel(S.filter));
  else if (v === 'detail') { app.innerHTML = shell(viewDetail(), (cur() || {}).title || ''); afterDetail(); }
}

/* ================= Liste ================= */
const AKTIV = ['entwurf', 'mandant', 'eingereicht', 'geprueft'];
function viewListe() {
  const q = S.q.trim().toLowerCase();
  const list = S.recs.filter(r => (S.filter === 'aktiv' ? AKTIV.includes(r.status) : S.filter === 'alle' ? true : r.status === S.filter))
    .filter(r => !q || (r.title + ' ' + (r.data.email || '') + ' ' + (r.datevNr || '')).toLowerCase().includes(q))
    .sort((a, b) => (b.updatedAt || '').localeCompare(a.updatedAt || ''));
  const rows = list.map(r => {
    const ev = LG.evaluate(r);
    return `<tr>
      <td><button class="name" data-open-rec="${r.id}">${esc(r.title)}</button><div class="sm muted">${esc(SC.TYPES[r.typ].label)}${r.datevNr ? ' · DATEV ' + esc(r.datevNr) : ''}</div></td>
      <td><span class="badge st-${r.status}">${esc(statusLabel(r.status))}</span></td>
      <td><span class="bar"><span style="width:${ev.pct}%"></span></span><span class="sm">${ev.pct} %</span></td>
      <td class="sm">${ev.openList.length ? `<span class="badge st-mandant">${ev.openList.length} offen</span>` : '<span class="muted">–</span>'}</td>
      <td class="sm">${esc(r.createdBy || '')}<div class="muted">${fmtDT(r.createdAt).slice(0, 10)}</div></td>
      <td class="sm muted">${fmtDT(r.updatedAt)}</td></tr>`;
  }).join('');
  return `<div class="card">
    <div class="card-h"><div><div class="card-t">${esc(navLabel(S.filter))} <span class="muted" style="font-weight:400">(${list.length})</span></div><div class="sm muted">Erfassung im Erstgespräch, Nachforderung beim Mandanten, Prüfung und Übergabe an DATEV</div></div>
      <div class="row"><label class="sr" for="q">Suche</label><input id="q" class="inp search" type="search" placeholder="Name, E-Mail, DATEV-Nr. …" value="${esc(S.q)}">
      <button class="btn primary" data-act="new">${I.plus}Neue Aufnahme</button></div></div>
    ${list.length ? `<div class="tblwrap"><table class="tbl"><thead><tr><th>Mandant</th><th>Status</th><th>Vollständigkeit</th><th>Beim Mandanten</th><th>Erfasst</th><th>Geändert</th></tr></thead><tbody>${rows}</tbody></table></div>`
      : `<div class="empty">${S.recs.length ? 'Keine Aufnahmen in dieser Ansicht.' : 'Noch keine Aufnahmen. Legen Sie die erste über „Neue Aufnahme“ an.'}</div>`}
  </div>`;
}

/* ================= Neue Aufnahme ================= */
function dlgNew() {
  openDlg(`<form method="dialog" id="fnew">
    <div class="dlg-h">Neue Mandantenaufnahme</div>
    <div class="dlg-b">
      <div class="f"><span class="lbl">Mandantentyp</span><div class="typcards">${Object.entries(SC.TYPES).map(([k, t], i) => `<label class="typcard"><input type="radio" name="typ" value="${k}" ${i === 0 ? 'checked' : ''}>${esc(t.label)}</label>`).join('')}</div></div>
      <div class="grid" id="newperson">
        <div class="f"><label for="n-anrede">Anrede</label><select id="n-anrede"><option value="">–</option>${['Herr', 'Frau', 'ohne Anrede'].map(o => `<option>${o}</option>`).join('')}</select></div><div></div>
        <div class="f"><label for="n-vor">Vorname</label><input id="n-vor"></div>
        <div class="f"><label for="n-nach">Nachname<span class="req">*</span></label><input id="n-nach"></div>
      </div>
      <div class="f" id="newfirma" hidden><label for="n-firma">Firma<span class="req">*</span></label><input id="n-firma" placeholder="z. B. Muster GmbH"></div>
      <div class="f"><label for="n-mail">E-Mail</label><input id="n-mail" type="email" placeholder="für Rückfragen und Nachforderungen"></div>
    </div>
    <div class="dlg-f"><button type="button" class="btn" data-close>Abbrechen</button><button class="btn primary" value="ok" id="newok">Anlegen und erfassen</button></div></form>`);
  const sync = () => { const g = document.querySelector('input[name=typ]:checked').value === 'gmbh'; $('#newperson').hidden = g; $('#newfirma').hidden = !g; };
  document.querySelectorAll('input[name=typ]').forEach(x => x.addEventListener('change', sync));
  $('#newok').addEventListener('click', async e => {
    e.preventDefault();
    const typ = document.querySelector('input[name=typ]:checked').value;
    const data = {};
    if (typ === 'gmbh') { data.firma = $('#n-firma').value.trim(); if (!data.firma) return $('#n-firma').focus(); }
    else { data.anrede = $('#n-anrede').value; data.vorname = $('#n-vor').value.trim(); data.nachname = $('#n-nach').value.trim(); if (!data.nachname) return $('#n-nach').focus(); }
    const mail = $('#n-mail').value.trim(); if (mail) data.email = mail;
    closeDlg();
    await createRec(typ, data);
  });
}
async function createRec(typ, data) {
  const at = nowIso();
  const rec = { id: uid(), typ, status: 'entwurf', data, open: {}, docs: {}, history: [{ at, by: S.user.name, t: 'Aufnahme angelegt' }], createdAt: at, createdBy: S.user.name, updatedAt: at, updatedBy: S.user.name };
  rec.title = LG.computeTitle(rec);
  rec.folder = `${at.slice(0, 10)} ${safeName(rec.title)} ${rec.id.slice(-4)}`;
  rec.pct = LG.evaluate(rec).pct;
  try { await S.store.create(rec); S.recs.push(rec); S.sec = null; go('#/a/' + rec.id); toast('Aufnahme angelegt', 'ok'); }
  catch (e) { toast('Anlegen fehlgeschlagen: ' + e.message, 'err'); }
}

/* ================= Detail ================= */
function sectionList(rec) { return LG.sectionsFor(rec.typ); }
function viewDetail() {
  const rec = cur(); if (!rec) return '<div class="empty">Nicht gefunden.</div>';
  const secs = sectionList(rec);
  if (!secs.find(s => s.k === S.sec)) S.sec = secs[0].k;
  const idx = SC.STATUS.findIndex(s => s.k === rec.status);
  return `<div class="card head">
    <div class="head-row">
      <div style="min-width:0">
        <button class="btn link sm" data-act="home">${I.back}Alle Aufnahmen</button>
        <h1 id="rtitle">${esc(rec.title)}</h1>
        <div class="sm muted">Erfasst von ${esc(rec.createdBy || '')} am ${fmtDT(rec.createdAt)}${rec.datevNr ? ` · DATEV-Mandantennummer <b>${esc(rec.datevNr)}</b>` : ''} · <span id="saveinfo" class="saveinfo">${saveText()}</span></div>
      </div>
      <div class="actions">
        <label class="sr" for="typsel">Mandantentyp</label>
        <select id="typsel" class="typsel" data-act-change="typ">${Object.entries(SC.TYPES).map(([k, t]) => `<option value="${k}" ${k === rec.typ ? 'selected' : ''}>${esc(t.label)}</option>`).join('')}</select>
        <button class="btn" data-act="print">${I.print}Stammblatt</button>
        <button class="btn primary" data-act="mail">${I.mail}E-Mail an Mandanten</button>
      </div>
    </div>
    ${S.conflict ? `<div class="banner err">Diese Aufnahme wurde zwischenzeitlich von jemand anderem geändert. Ihre letzten Änderungen sind noch nicht gespeichert.
      <button class="btn sm" data-act="conflict-reload">Neu laden (meine Änderungen verwerfen)</button><button class="btn sm danger" data-act="conflict-force">Meine Version speichern</button></div>` : ''}
    ${rec.status === 'abgebrochen' ? `<div class="banner err">Aufnahme abgebrochen${rec.abbruchGrund ? ': ' + esc(rec.abbruchGrund) : ''}. <button class="btn sm" data-act="reactivate">Wieder aufnehmen</button></div>` : `
    <ol class="steps" aria-label="Status">${SC.STATUS.map((s, i) => `<li><button data-status="${s.k}" class="${i === idx ? 'cur' : i < idx ? 'done' : ''}" ${i === idx ? 'aria-current="step"' : ''}>${i + 1} ${esc(s.l)}</button></li>`).join('')}</ol>`}
  </div>
  <div class="cols">
    <section class="card formcard" aria-label="Erfassung">
      <nav class="secnav" id="secnav" aria-label="Abschnitte"></nav>
      <div class="secbody" id="secbody"></div>
    </section>
    <aside class="panel" id="side"></aside>
  </div>`;
}
function afterDetail() { renderNav(); renderSection(); renderSide(); }

function renderNav() {
  const rec = cur(), ev = LG.evaluate(rec), el = $('#secnav'); if (!el) return;
  const secs = sectionList(rec), btns = [...el.querySelectorAll('button[data-sec]')];
  if (btns.length === secs.length && btns.every((b, i) => b.dataset.sec === secs[i].k)) {
    // vorhandene Knöpfe aktualisieren statt ersetzen (Klicks bleiben erhalten)
    btns.forEach(b => {
      const st = ev.sections[b.dataset.sec], on = b.dataset.sec === S.sec;
      b.classList.toggle('on', on); if (on) b.setAttribute('aria-current', 'true'); else b.removeAttribute('aria-current');
      b.querySelector('.dot').className = 'dot ' + st.state;
      const c = b.querySelector('.cnt'); if (c) c.textContent = `${st.done}/${st.total}`;
    });
    return;
  }
  el.innerHTML = sectionList(rec).map(s => {
    const st = ev.sections[s.k];
    return `<button data-sec="${s.k}" class="${s.k === S.sec ? 'on' : ''}" ${s.k === S.sec ? 'aria-current="true"' : ''}><span class="dot ${st.state}" aria-hidden="true"></span>${esc(s.l)}${st.total ? `<span class="cnt">${st.done}/${st.total}</span>` : ''}</button>`;
  }).join('');
}

function fieldBy(rec, path) {
  const parts = path.split('.');
  for (const sec of LG.sectionsFor(rec.typ)) {
    for (const f of (sec.fields || [])) {
      if (!LG.appliesTo(f, rec.typ) || f.k !== parts[0]) continue;
      if (parts.length === 1) return f;
      if (f.t === 'group') return f.fields.find(x => x.k === parts[2]);
    }
  }
  return null;
}
function getVal(rec, path) {
  const p = path.split('.'), d = rec.data;
  if (p.length === 1) return d[p[0]];
  return ((d[p[0]] || [])[+p[1]] || {})[p[2]];
}
function setValAt(rec, path, v) {
  const p = path.split('.'), d = rec.data;
  if (p.length === 1) { d[p[0]] = v; return; }
  if (!Array.isArray(d[p[0]])) d[p[0]] = [];
  d[p[0]][+p[1]] = d[p[0]][+p[1]] || {};
  d[p[0]][+p[1]][p[2]] = v;
}
const fid = path => 'f-' + path.replace(/\./g, '-');

function inputHtml(f, val, path) {
  const id = fid(path), ph = f.ph ? ` placeholder="${esc(f.ph)}"` : '';
  const v = val ?? '';
  switch (f.t) {
    case 'select': return `<select id="${id}" data-p="${path}"><option value="">– bitte wählen –</option>${f.o.map(o => `<option ${o === v ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
    case 'textarea': return `<textarea id="${id}" data-p="${path}" rows="3"${ph}>${esc(v)}</textarea>`;
    case 'date': return `<input id="${id}" type="date" data-p="${path}" value="${esc(v)}" max="9999-12-31">`;
    case 'email': return `<input id="${id}" type="email" data-p="${path}" value="${esc(v)}"${ph} autocomplete="off">`;
    case 'tel': return `<input id="${id}" type="tel" data-p="${path}" value="${esc(v)}"${ph}>`;
    case 'num': return `<input id="${id}" inputmode="decimal" data-p="${path}" value="${esc(v)}"${ph}>`;
    case 'stid': return `<input id="${id}" inputmode="numeric" data-p="${path}" value="${esc(v)}" placeholder="${esc(f.ph || '11 Ziffern')}" autocomplete="off">`;
    case 'plz': return `<input id="${id}" inputmode="numeric" data-p="${path}" value="${esc(v)}" maxlength="10">`;
    case 'iban': return `<input id="${id}" data-p="${path}" value="${esc(v)}" placeholder="DE00 0000 0000 0000 0000 00" autocomplete="off" spellcheck="false">`;
    default: return `<input id="${id}" data-p="${path}" value="${esc(v)}"${ph} autocomplete="off">`;
  }
}
function fieldHtml(rec, f, path, opts = {}) {
  const val = getVal(rec, path), isOpen = !opts.inGroup && rec.open[path] && LG.isEmpty(val);
  const err = LG.checkValue(f, val, rec.data);
  const req = opts.inGroup ? f.req : LG.isReq(f, rec.typ);
  const tog = (!opts.inGroup && !f.internal && !f.kanzlei) ? `<button type="button" tabindex="-1" class="mtog" data-mopen="${path}" aria-pressed="${rec.open[path] ? 'true' : 'false'}" title="Beim Mandanten anfragen">→ Mandant</button>` : '';
  if (f.t === 'check') {
    return `<div class="f fcheck ${f.w === 'full' ? 'full' : ''}"><div class="lab"><label class="chk"><input type="checkbox" id="${fid(path)}" data-p="${path}" ${val ? 'checked' : ''}>${esc(f.l)}${req ? '<span class="req">*</span>' : ''}</label>${tog}</div></div>`;
  }
  if (f.t === 'multi') {
    const arr = Array.isArray(val) ? val : [];
    return `<div class="f fmulti full ${isOpen ? 'isopen' : ''}" role="group" aria-labelledby="${fid(path)}-l"><div class="lab"><span class="lbl" id="${fid(path)}-l">${esc(f.l)}${req ? '<span class="req">*</span>' : ''}</span>${tog}</div>
      <div class="multi">${f.o.map(o => `<label class="chk"><input type="checkbox" data-p="${path}" data-multi="${esc(o)}" ${arr.includes(o) ? 'checked' : ''}>${esc(o)}</label>`).join('')}</div></div>`;
  }
  return `<div class="f ${f.w === 'full' || f.t === 'textarea' ? 'full' : ''} ${isOpen ? 'isopen' : ''} ${err ? 'bad' : ''}" data-fwrap="${path}">
    <div class="lab"><label for="${fid(path)}">${esc(f.l)}${req ? '<span class="req" aria-label="Pflichtfeld">*</span>' : ''}</label>${tog}</div>
    ${inputHtml(f, val, path)}
    <div class="ferr" id="${fid(path)}-e">${esc(err)}</div>${f.hint ? `<div class="fhint">${esc(f.hint)}</div>` : ''}</div>`;
}
function groupHtml(rec, g) {
  const rows = Array.isArray(rec.data[g.k]) ? rec.data[g.k] : [];
  const req = LG.isReq(g, rec.typ);
  return `<div class="group full" style="grid-column:1/-1">
    <div class="group-h"><span class="t">${esc(g.l)}${req ? '<span class="req">*</span>' : ''} <span class="sm muted">(${rows.length})</span></span>
      <button type="button" tabindex="-1" class="mtog" data-mopen="${g.k}" aria-pressed="${rec.open[g.k] ? 'true' : 'false'}">→ Mandant</button></div>
    ${rows.map((row, i) => `<div class="grow"><div class="grow-h"><span>${esc(g.l.split(' / ')[0])} ${i + 1}</span><button type="button" class="btn sm danger" data-gdel="${g.k}.${i}">Entfernen</button></div>
      <div class="grid">${LG.rowFields(g, row, rec.data).map(f => fieldHtml(rec, f, `${g.k}.${i}.${f.k}`, { inGroup: true })).join('')}</div></div>`).join('')}
    <div><button type="button" class="btn sm" data-gadd="${g.k}">${I.plus}${esc(g.add || 'Hinzufügen')}</button></div></div>`;
}
function docsHtml(rec) {
  const docs = LG.docsFor(rec.typ), max = C.maxUploadMB || 25;
  return `<div class="docs">${docs.map(doc => {
    const st = (rec.docs[doc.k] || {}), files = st.files || [], state = st.state || 'offen', key = 'doc:' + doc.k;
    const isOpen = rec.open[key] && state === 'offen';
    return `<div class="doc ${isOpen ? 'isopen' : ''}">
      <div><div class="dl">${esc(doc.l)}${doc.req ? '<span class="req">*</span>' : ''}</div></div>
      <div class="dctl">
        <button type="button" tabindex="-1" class="mtog" data-mopen="${key}" aria-pressed="${rec.open[key] ? 'true' : 'false'}">→ Mandant</button>
        <label class="sr" for="ds-${doc.k}">Status ${esc(doc.l)}</label>
        <select id="ds-${doc.k}" data-docstate="${doc.k}">${SC.DOC_STATES.map(s => `<option value="${s.k}" ${s.k === state ? 'selected' : ''} ${s.k === 'datei' && !files.length ? 'disabled' : ''}>${esc(s.l)}</option>`).join('')}</select>
        <span class="btn sm up">${I.up}Hochladen<input type="file" multiple data-upload="${doc.k}" accept=".pdf,.jpg,.jpeg,.png,.heic,.tif,.tiff,.docx,.xlsx" aria-label="Datei hochladen: ${esc(doc.l)}"></span>
      </div>
      ${files.length ? `<div class="files">${files.map(f => { const h = S.store.href(f); return h ? `<a class="file" href="${esc(h)}" target="_blank" rel="noopener">${I.file}${esc(f.name)}</a>` : `<span class="file" title="In der Demo nach dem Neuladen nicht mehr verfügbar">${I.file}${esc(f.name)}</span>`; }).join('')}</div>` : ''}
    </div>`;
  }).join('')}</div><div class="fhint">PDF, Bilder, Word, Excel · max. ${max} MB je Datei · Ablage in SharePoint, Bibliothek „${LIB_NAME}“, Ordner „${esc(rec.folder || '')}“</div>`;
}
function renderSection() {
  const rec = cur(), el = $('#secbody'); if (!el) return;
  const secs = sectionList(rec), i = secs.findIndex(s => s.k === S.sec), sec = secs[i];
  let body;
  if (sec.docs) body = docsHtml(rec);
  else {
    const fs = LG.fieldsFor(sec, rec);
    body = `<div class="grid">${fs.map(f => f.t === 'group' ? groupHtml(rec, f) : fieldHtml(rec, f, f.k)).join('')}</div>`;
  }
  const prev = secs[i - 1], next = secs[i + 1];
  el.innerHTML = `<div><h2>${esc(sec.l)}</h2>${sec.hint ? `<div class="hint">${esc(sec.hint)}</div>` : ''}
    ${rec.typ === 'gmbh' && sec.k === 'gwg' ? '<div class="hint">Identifiziert wird die für die Gesellschaft auftretende Person (in der Regel der Geschäftsführer).</div>' : ''}</div>
    ${body}
    <div class="secfoot"><span>${prev ? `<button class="btn" data-sec="${prev.k}">${I.back}${esc(prev.l)}</button>` : ''}</span>
      <span>${next ? `<button class="btn soft" data-sec="${next.k}">Weiter: ${esc(next.l)}</button>` : `<button class="btn soft" data-act="print">${I.print}Stammblatt drucken</button>`}</span></div>`;
}
function renderSide() {
  const rec = cur(), el = $('#side'); if (!el) return;
  const ev = LG.evaluate(rec);
  const ck = (l, t) => `<div class="ck ${l}"><b aria-hidden="true">${{ err: '✕', warn: '!', ok: '✓' }[l]}</b><span>${esc(t)}</span></div>`;
  const miss = ev.missing;
  const canFlag = miss.filter(m => !m.mandant && !m.kanzlei && !m.key.includes('.')).length;
  el.innerHTML = `<section class="card card-b" aria-label="Vollständigkeit">
      <div class="pct"><span class="card-t">Vollständigkeit</span><b>${ev.pct} %</b></div>
      <div class="pbar"><span style="width:${ev.pct}%"></span></div>
      <div class="checks">${ev.errors.map(t => ck('err', t)).join('')}${ev.warns.map(t => ck('warn', t)).join('')}${ev.oks.map(t => ck('ok', t)).join('')}</div>
      ${miss.length ? `<details class="miss"><summary>${miss.length} Pflichtangabe(n) fehlen</summary><ul>${miss.map(m => `<li><button data-goto="${m.sec}|${m.key}"><span class="sec">${esc(m.section)}:</span> ${esc(m.label)}${m.mandant ? ' (angefragt)' : ''}</button></li>`).join('')}</ul></details>` : ''}
    </section>
    <section class="card card-b" aria-label="Offen für den Mandanten" style="display:flex;flex-direction:column;gap:10px">
      <span class="card-t">Beim Mandanten anfragen <span class="muted" style="font-weight:400">(${ev.openList.length})</span></span>
      ${ev.openList.length ? `<ul class="openlist">${ev.openList.map(o => `<li>${esc(o.label)}</li>`).join('')}</ul>` : '<div class="sm muted">Noch nichts markiert. Über „→ Mandant“ an einem Feld oder einer Unterlage festlegen, was der Mandant nachliefern soll.</div>'}
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        ${canFlag ? `<button class="btn sm" data-act="flag-all">Alle fehlenden übernehmen (${canFlag})</button>` : ''}
        <button class="btn sm primary" data-act="mail" ${ev.openList.length ? '' : 'disabled'}>${I.mail}E-Mail vorbereiten</button>
      </div>
    </section>
    <section class="card card-b" aria-label="Verlauf">
      <div class="card-t" style="margin-bottom:8px">Verlauf</div>
      <ul class="hist">${(rec.history || []).slice(-8).reverse().map(h => `<li><time>${fmtDT(h.at)} · ${esc(h.by)}</time>${esc(h.t)}</li>`).join('')}</ul>
      ${rec.status !== 'abgebrochen' && rec.status !== 'angelegt' ? '<div style="margin-top:12px"><button class="btn sm danger" data-act="cancel">Aufnahme abbrechen</button></div>' : ''}
    </section>`;
}
function refreshLive(path) {
  const rec = cur();
  const t = LG.computeTitle(rec); if (t !== rec.title) { rec.title = t; const h = $('#rtitle'); if (h) h.textContent = t; }
  if (path) {
    const f = fieldBy(rec, path), w = document.querySelector(`[data-fwrap="${CSS.escape(path)}"]`);
    if (f && w) {
      const err = LG.checkValue(f, getVal(rec, path), rec.data);
      w.classList.toggle('bad', !!err);
      w.classList.toggle('isopen', !!rec.open[path] && LG.isEmpty(getVal(rec, path)));
      const e = document.getElementById(fid(path) + '-e'); if (e) e.textContent = err;
    }
  }
  renderNav(); renderSide();
}
function rerenderKeepFocus() {
  const a = document.activeElement, id = a && a.id;
  renderNav(); renderSection(); renderSide();
  if (id) { const n = document.getElementById(id); if (n) n.focus(); }
}

/* ================= Speichern ================= */
function saveText() {
  if (S.saveState === 'saving') return 'Speichert …';
  if (S.saveState === 'pending') return 'Änderungen werden gespeichert …';
  if (S.saveState === 'error') return 'Nicht gespeichert!';
  const r = cur(); return r ? 'Gespeichert ' + fmtT(r.updatedAt) : '';
}
function showSave() { const el = $('#saveinfo'); if (el) { el.textContent = saveText(); el.className = 'saveinfo' + (S.saveState === 'error' ? ' err' : ''); } }
function touch(rec, histText) {
  rec.updatedAt = nowIso(); rec.updatedBy = S.user.name; rec.title = LG.computeTitle(rec); rec.pct = LG.evaluate(rec).pct;
  if (histText) { rec.history = rec.history || []; rec.history.push({ at: rec.updatedAt, by: S.user.name, t: histText }); }
  S.saveState = 'pending'; showSave();
  clearTimeout(S.saveTimer); S.saveTimer = setTimeout(() => saveNow(rec), histText ? 50 : 1200);
}
async function saveNow(rec, force) {
  S.saveTimer = null;
  if (S.saving) { S.saveAgain = true; return; }
  S.saving = true; S.saveState = 'saving'; showSave();
  try { await S.store.save(rec, force); S.saveState = ''; S.conflict = false; }
  catch (e) {
    S.saveState = 'error';
    if (e.conflict) { S.conflict = true; if (S.view === 'detail') { render(); } }
    else toast('Speichern fehlgeschlagen: ' + e.message, 'err');
  }
  S.saving = false; showSave();
  if (S.saveAgain && !S.conflict) { S.saveAgain = false; saveNow(rec); }
}

/* ================= Status ================= */
async function setStatus(k) {
  const rec = cur(); if (!rec || rec.status === k) return;
  const ev = LG.evaluate(rec), target = SC.STATUS.findIndex(s => s.k === k);
  if (target >= SC.STATUS.findIndex(s => s.k === 'geprueft')) {
    if (ev.errors.length || ev.pct < 100) { toast(`Noch nicht prüfbar: ${ev.missing.length} Pflichtangabe(n) fehlen, ${ev.errors.length} Fehler.`, 'err'); return; }
  }
  if (k === 'geprueft' || (k === 'angelegt' && !rec.geprueftAm)) {
    rec.geprueftVon = S.user.name; rec.geprueftAm = nowIso();
    if (rec.createdBy === S.user.name) toast('Hinweis: Prüfung durch eine zweite Person (Vier-Augen-Prinzip) ist empfehlenswert.');
  }
  if (k === 'angelegt') {
    const nr = await ask('In DATEV angelegt', `<div class="f"><label for="a-nr">DATEV-Mandantennummer<span class="req">*</span></label><input id="a-nr" inputmode="numeric" value="${esc(rec.datevNr || '')}"></div>
      <div class="hint">Die Nummer erscheint in der Liste und auf dem Stammblatt.</div>`, 'Speichern', () => { const v = $('#a-nr').value.trim(); return /^\d{1,7}$/.test(v) ? v : null; });
    if (!nr) return;
    rec.datevNr = nr;
  }
  if (target < SC.STATUS.findIndex(s => s.k === 'geprueft')) { delete rec.geprueftVon; delete rec.geprueftAm; }
  rec.status = k;
  touch(rec, 'Status: ' + statusLabel(k) + (k === 'angelegt' ? ` (DATEV-Nr. ${rec.datevNr})` : ''));
  render();
  if (k === 'mandant' && ev.openList.length) dlgMail();
}

/* ================= E-Mail an Mandanten ================= */
function dlgMail() {
  const rec = cur(), ev = LG.evaluate(rec), K = C.kanzlei || {};
  const sections = LG.checklist(rec), count = sections.reduce((n, x) => n + x.items.length, 0);
  const m = LG.mailShort(rec, S.user.name, K, C.mailHinweis);
  const pdfOpts = () => ({ rec, sections, kanzlei: K, logo: C.logoDunkel || 'icons/logo-dunkel.png', user: S.user.name });
  openDlg(`<form method="dialog"><div class="dlg-h">E-Mail an den Mandanten</div>
    <div class="dlg-b">
      ${count ? `<div class="banner ok" style="background:var(--ok-bg);color:var(--ok);border-color:var(--ok)">Checkliste mit ${count} Punkt(en) wird als PDF angehängt.
        <button type="button" class="btn sm" id="m-pdf" style="margin-left:auto">Checkliste ansehen</button></div>`
      : '<div class="banner warn">Es ist noch nichts für den Mandanten markiert. Über „→ Mandant“ an Feldern und Unterlagen festlegen, was angefragt werden soll.</div>'}
      <div class="f"><label for="m-to">An</label><input id="m-to" type="email" value="${esc(m.to)}"></div>
      <div class="f"><label for="m-sub">Betreff</label><input id="m-sub" value="${esc(m.subject)}"></div>
      <div class="f"><label for="m-body">Text</label><textarea id="m-body" class="mail" style="min-height:240px">${esc(m.text)}</textarea></div>
      ${rec.status === 'entwurf' ? '<label class="chk"><input type="checkbox" id="m-st" checked>Status auf „Beim Mandanten“ setzen</label>' : ''}
      <div class="hint">Es wird nichts automatisch versendet: In Outlook entsteht ein Entwurf mit der Checkliste im Anhang, den Sie prüfen und selbst senden. Er liegt auch im Ordner „Entwürfe“ Ihres Postfachs.</div>
    </div>
    <div class="dlg-f"><button type="button" class="btn" data-close>Schließen</button>
      <button type="button" class="btn primary" id="m-draft" ${count ? '' : 'disabled'}>${I.mail}Outlook-Entwurf erstellen</button></div></form>`);
  const pdfBtn = $('#m-pdf');
  if (pdfBtn) pdfBtn.addEventListener('click', async () => {
    try { const blob = await CHECKLIST.blob(pdfOpts()); const u = URL.createObjectURL(blob); window.open(u, '_blank'); setTimeout(() => URL.revokeObjectURL(u), 60000); }
    catch (e) { toast('Checkliste konnte nicht erstellt werden: ' + e.message, 'err'); }
  });
  $('#m-draft').addEventListener('click', async e => {
    const btn = e.currentTarget; btn.disabled = true; btn.textContent = 'Erstelle Entwurf …';
    try {
      const token = await S.store.mailToken();
      const opts = pdfOpts(), name = CHECKLIST.fileName(rec);
      const to = $('#m-to').value.trim(), subject = $('#m-sub').value, text = $('#m-body').value;
      let link = null;
      if (C.demo) {
        const blob = await CHECKLIST.blob(opts); const u = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = u; a.download = name; document.body.appendChild(a); a.click(); a.remove();
        toast('Demo: Checkliste heruntergeladen. In der echten App entsteht hier ein Outlook-Entwurf mit Anhang.');
      } else {
        const pdfB64 = await CHECKLIST.base64(opts);
        link = await S.store.createDraft(token, { to, subject, html: m.toHtml(text), fileName: name, pdfB64 });
      }
      const st = $('#m-st');
      const hist = `Outlook-Entwurf mit Checkliste erstellt (${count} Punkte)`;
      if (st && st.checked) rec.status = 'mandant';
      touch(rec, hist);
      closeDlg(); render();
      if (link) { window.open(link, '_blank', 'noopener'); toast('Entwurf in Outlook erstellt – bitte prüfen und senden', 'ok'); }
    } catch (err) {
      console.error(err); toast(err.message, 'err');
      btn.disabled = false; btn.innerHTML = `${I.mail}Outlook-Entwurf erstellen`;
    }
  });
}

/* ================= Stammblatt ================= */
function fmtVal(f, v) {
  if (LG.isEmpty(v)) return '';
  switch (f.t) {
    case 'date': return LG.deDate(v);
    case 'check': return v ? 'ja' : 'nein';
    case 'multi': return v.join(', ');
    case 'iban': return LG.fmtIban(v);
    case 'ustid': return LG.normUst(v);
    default: return String(v);
  }
}
function printStammblatt() {
  const rec = cur(), ev = LG.evaluate(rec), K = C.kanzlei || {};
  const cell = (f, v, key) => { const t = fmtVal(f, v); return t ? esc(t) : (rec.open[key] ? '<span class="pb-open">beim Mandanten angefragt</span>' : '—'); };
  let html = `<div class="pb-head"><div><img src="${esc(C.logoDunkel || 'icons/logo-dunkel.png')}" alt="${esc(K.name || '')}"></div>
    <div class="pb-k">${esc(K.name || '')}<br>${esc(K.adresse || '')}<br>${esc([K.telefon, K.email].filter(Boolean).join(' · '))}</div></div>
    <p class="pb-title">Mandantenstammblatt</p><p class="pb-sub">${esc(rec.title)} · ${esc(SC.TYPES[rec.typ].label)}</p>
    <div class="pb-meta"><div><b>Status</b>${esc(statusLabel(rec.status))}</div><div><b>DATEV-Mandantennr.</b>${esc(rec.datevNr || '—')}</div>
      <div><b>Vollständigkeit</b>${ev.pct} %</div><div><b>Stand</b>${fmtDT(nowIso())}</div></div>`;
  for (const sec of LG.sectionsFor(rec.typ)) {
    if (sec.docs) {
      html += `<h2>${esc(sec.l)}</h2><table><thead><tr><th>Unterlage</th><th>Status</th><th>Dateien</th></tr></thead><tbody>${LG.docsFor(rec.typ).map(doc => {
        const st = rec.docs[doc.k] || {}, s = SC.DOC_STATES.find(x => x.k === (st.state || 'offen'));
        return `<tr><td class="l">${esc(doc.l)}${doc.req ? ' *' : ''}</td><td>${st.state && st.state !== 'offen' ? esc(s.l) : (rec.open['doc:' + doc.k] ? '<span class="pb-open">beim Mandanten angefragt</span>' : 'offen')}</td><td>${esc((st.files || []).map(f => f.name).join(', '))}</td></tr>`;
      }).join('')}</tbody></table>`;
      continue;
    }
    const fs = LG.fieldsFor(sec, rec).filter(f => !f.internal);
    let rows = '', groups = '';
    for (const f of fs) {
      if (f.t === 'group') {
        const list = Array.isArray(rec.data[f.k]) ? rec.data[f.k] : [];
        groups += `<table style="margin-top:6px"><thead><tr><th colspan="${f.fields.length}">${esc(f.l)}${list.length ? '' : (rec.open[f.k] ? ' – <span class="pb-open">beim Mandanten angefragt</span>' : ' – keine')}</th></tr>
          ${list.length ? `<tr>${f.fields.map(x => `<th>${esc(x.l)}</th>`).join('')}</tr>` : ''}</thead>
          <tbody>${list.map(row => `<tr>${f.fields.map(x => `<td>${esc(fmtVal(x, row[x.k])) || '—'}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
      } else rows += `<tr><td class="l">${esc(f.l)}</td><td>${cell(f, rec.data[f.k], f.k)}</td></tr>`;
    }
    html += `<h2>${esc(sec.l)}</h2>${rows ? `<table><tbody>${rows}</tbody></table>` : ''}${groups}`;
  }
  html += `<h2>Prüfvermerk</h2><table><tbody>
    <tr><td class="l">Erfasst von / am</td><td>${esc(rec.createdBy || '')} · ${fmtDT(rec.createdAt)}</td></tr>
    <tr><td class="l">Geprüft von / am</td><td>${rec.geprueftVon ? esc(rec.geprueftVon) + ' · ' + fmtDT(rec.geprueftAm) : '—'}</td></tr>
    ${ev.errors.length || ev.missing.length ? `<tr><td class="l">Offene Punkte</td><td>${ev.missing.length} fehlende Pflichtangabe(n)${ev.errors.length ? ', ' + ev.errors.length + ' Fehler' : ''}</td></tr>` : ''}
  </tbody></table>
  <div class="pb-sign"><div>Ort, Datum, Unterschrift Mandant<br><span style="color:#555">Die Angaben sind vollständig und richtig.</span></div><div>Ort, Datum, Unterschrift ${esc(K.kurz || 'Kanzlei')}</div></div>
  <div class="pb-foot">${esc(VERSION)} · * Pflichtunterlage · Vertraulich – unterliegt der Verschwiegenheitspflicht (§ 57 StBerG)</div>`;
  if (PREVIEW) {
    openDlg(`<form method="dialog"><div class="dlg-h">Stammblatt (Vorschau)</div><div class="dlg-b"><div id="print" class="preview">${html}</div></div>
      <div class="dlg-f"><span class="hint" style="margin-right:auto">In der App öffnet sich hier der Druckdialog (Drucken oder als PDF speichern).</span><button type="button" class="btn" data-close>Schließen</button></div></form>`);
    $('#dlg').classList.add('wide'); $('#dlg').addEventListener('close', () => $('#dlg').classList.remove('wide'), { once: true });
    return;
  }
  $('#print').innerHTML = html;
  const old = document.title; document.title = 'Stammblatt ' + rec.title;
  setTimeout(() => { window.print(); document.title = old; }, 50);
}

/* ================= Dialog-Helfer ================= */
document.addEventListener('click', e => { if (e.target.closest('#dlg [data-close]')) closeDlg(); });
function openDlg(html) { const d = $('#dlg'); d.innerHTML = html; d.showModal(); const first = d.querySelector('input:not([type=radio]):not([hidden]),textarea,select'); if (first) first.focus(); }
function closeDlg() { const d = $('#dlg'); if (d.open) d.close(); }
function confirmDlg(title, text, okLabel = 'OK') {
  return ask(title, `<p style="margin:0">${esc(text)}</p>`, okLabel, () => true);
}
function ask(title, bodyHtml, okLabel, getter) {
  return new Promise(res => {
    openDlg(`<form method="dialog"><div class="dlg-h">${esc(title)}</div><div class="dlg-b">${bodyHtml}</div>
      <div class="dlg-f"><button type="button" class="btn" data-close>Abbrechen</button><button class="btn primary" id="ask-ok">${esc(okLabel)}</button></div></form>`);
    const d = $('#dlg');
    const onClose = () => { d.removeEventListener('close', onClose); res(null); };
    d.addEventListener('close', onClose);
    $('#ask-ok').addEventListener('click', e => {
      e.preventDefault(); const v = getter(); if (v == null || v === '') { const i = d.querySelector('input,textarea'); if (i) i.focus(); return; }
      d.removeEventListener('close', onClose); d.close(); res(v);
    });
  });
}

/* ================= Ereignisse ================= */
function bindEvents() {
  const app = $('#app');
  app.addEventListener('click', async e => {
    const t = e.target.closest('button, a[data-act]'); if (!t) return;
    const ds = t.dataset, rec = cur();
    if (t.tagName === 'A' && ds.act) e.preventDefault();
    if (ds.act) {
      switch (ds.act) {
        case 'home': return go('#/');
        case 'login': return Auth.login();
        case 'logout': return Auth.logout();
        case 'reload': return location.reload();
        case 'refresh': await loadAll(); toast('Daten neu geladen'); return route();
        case 'demo-reset': if (await confirmDlg('Demo zurücksetzen', 'Alle Demo-Daten löschen und neue Testdaten erzeugen?', 'Zurücksetzen')) { Demo.reset(); await Demo.init(); await loadAll(); go('#/'); } return;
        case 'setup': t.disabled = true; t.textContent = 'Lege an …';
          try { await SP.setup(); await loadAll(); toast('Listen angelegt', 'ok'); route(); } catch (err) { toast('Anlegen fehlgeschlagen: ' + err.message, 'err'); t.disabled = false; t.textContent = 'Listen jetzt anlegen'; } return;
        case 'new': return dlgNew();
        case 'mail': return dlgMail();
        case 'print': return printStammblatt();
        case 'flag-all': { const ev = LG.evaluate(rec); ev.missing.filter(m => !m.key.includes('.') && !m.kanzlei).forEach(m => { rec.open[m.key] = true; }); touch(rec); rerenderKeepFocus(); return; }
        case 'cancel': {
          const g = await ask('Aufnahme abbrechen', `<div class="f"><label for="c-g">Grund<span class="req">*</span></label><input id="c-g" placeholder="z. B. Mandant hat sich anders entschieden"></div><div class="hint">Die Aufnahme bleibt mit allen Unterlagen erhalten (Aufbewahrung nach GwG) und kann wieder aufgenommen werden.</div>`, 'Abbrechen bestätigen', () => $('#c-g').value.trim());
          if (!g) return; rec.status = 'abgebrochen'; rec.abbruchGrund = g; touch(rec, 'Abgebrochen: ' + g); return render();
        }
        case 'reactivate': rec.status = 'entwurf'; delete rec.abbruchGrund; touch(rec, 'Wieder aufgenommen'); return render();
        case 'conflict-reload': { const fresh = await S.store.reload(rec); const i = S.recs.findIndex(r => r.id === rec.id); S.recs[i] = fresh; S.conflict = false; S.saveState = ''; return render(); }
        case 'conflict-force': S.conflict = false; await saveNow(rec, true); return render();
      }
    }
    if (ds.filter) { S.filter = ds.filter; if (S.view !== 'liste') return go('#/'); return render(); }
    if (ds.openRec) return go('#/a/' + ds.openRec);
    if (ds.status) return setStatus(ds.status);
    if (ds.sec) { S.sec = ds.sec; if (!PREVIEW) history.replaceState(null, '', `#/a/${rec.id}/${ds.sec}`); renderNav(); renderSection(); $('#secbody').scrollIntoView({ block: 'nearest' }); return; }
    if (ds.goto) {
      const [sec, key] = ds.goto.split('|'); S.sec = sec; renderNav(); renderSection();
      const el = document.getElementById(fid(key.replace(/^doc:/, ''))) || document.getElementById('ds-' + key.replace(/^doc:/, ''));
      if (el) { el.scrollIntoView({ block: 'center' }); el.focus(); }
      return;
    }
    if (ds.mopen) { rec.open[ds.mopen] = !rec.open[ds.mopen]; if (!rec.open[ds.mopen]) delete rec.open[ds.mopen]; t.setAttribute('aria-pressed', rec.open[ds.mopen] ? 'true' : 'false'); touch(rec); refreshLive(ds.mopen.startsWith('doc:') ? null : ds.mopen); if (ds.mopen.startsWith('doc:')) renderSection(); return; }
    if (ds.gadd) { if (!Array.isArray(rec.data[ds.gadd])) rec.data[ds.gadd] = []; rec.data[ds.gadd].push({}); touch(rec); rerenderKeepFocus(); return; }
    if (ds.gdel) { const [g, i] = ds.gdel.split('.'); if (!await confirmDlg('Eintrag entfernen', 'Diesen Eintrag mit allen Angaben entfernen?', 'Entfernen')) return; rec.data[g].splice(+i, 1); touch(rec); rerenderKeepFocus(); return; }
  });

  app.addEventListener('input', e => {
    const el = e.target;
    if (el.id === 'q') { S.q = el.value; const pos = el.selectionStart; render(); const q = $('#q'); q.focus(); q.setSelectionRange(pos, pos); return; }
    if (!el.dataset.p || el.type === 'checkbox' || el.tagName === 'SELECT') return;
    const rec = cur(); setValAt(rec, el.dataset.p, el.value); touch(rec); refreshLive(el.dataset.p);
  });

  app.addEventListener('change', async e => {
    const el = e.target, rec = cur();
    if (el.dataset.actChange === 'typ') {
      const neu = el.value; el.value = rec.typ;
      if (!await confirmDlg('Mandantentyp ändern', 'Bereits erfasste Angaben bleiben gespeichert, angezeigt werden aber nur die zum Typ passenden Felder.', 'Ändern')) return;
      el.value = neu;
      rec.typ = el.value; touch(rec, 'Mandantentyp: ' + SC.TYPES[rec.typ].label); S.sec = null; return render();
    }
    if (el.dataset.p) {
      const p = el.dataset.p;
      if (el.dataset.multi) { const arr = Array.isArray(getVal(rec, p)) ? [...getVal(rec, p)] : []; const o = el.dataset.multi; const i = arr.indexOf(o); if (el.checked && i < 0) arr.push(o); if (!el.checked && i >= 0) arr.splice(i, 1); setValAt(rec, p, arr); }
      else if (el.type === 'checkbox') setValAt(rec, p, el.checked);
      else setValAt(rec, p, el.value);
      const before = visibleSig(rec);
      const filled = autoFill(rec, p, el);
      touch(rec);
      if (filled || visibleSig(rec) !== before) return rerenderKeepFocus();
      // Textfelder wurden schon beim Tippen übernommen – hier nichts neu zeichnen,
      // sonst geht ein Klick, der das Feld verlässt, ins Leere.
      if (el.tagName === 'SELECT' || el.type === 'checkbox') return refreshLive(p);
      return;
    }
    if (el.dataset.docstate) {
      const k = el.dataset.docstate; rec.docs[k] = { ...(rec.docs[k] || {}), state: el.value };
      touch(rec, `Unterlage „${(SC.DOCS.find(d => d.k === k) || {}).l}“: ${SC.DOC_STATES.find(s => s.k === el.value).l}`); return rerenderKeepFocus();
    }
    if (el.dataset.upload) {
      const k = el.dataset.upload, files = [...el.files], max = (C.maxUploadMB || 25) * 1024 * 1024; el.value = '';
      for (const f of files) {
        if (f.size > max) { toast(`${f.name}: zu groß (max. ${C.maxUploadMB || 25} MB)`, 'err'); continue; }
        toast(`Lade „${f.name}“ hoch …`);
        try {
          const meta = await S.store.upload(rec, k, f);
          const st = rec.docs[k] = { ...(rec.docs[k] || {}), files: [...((rec.docs[k] || {}).files || []), { ...meta, at: nowIso(), by: S.user.name }] };
          st.state = 'datei';
          touch(rec, `Datei hochgeladen: ${meta.name}`); toast('Hochgeladen', 'ok');
        } catch (err) { toast(`Upload fehlgeschlagen (${f.name}): ${err.message}`, 'err'); }
      }
      return rerenderKeepFocus();
    }
  });
}
function visibleSig(rec) {
  return LG.sectionsFor(rec.typ).map(sec => LG.fieldsFor(sec, rec).map(f => f.t === 'group'
    ? f.k + '[' + (rec.data[f.k] || []).map(r => LG.rowFields(f, r, rec.data).map(x => x.k).join(',')).join('|') + ']' : f.k).join(';')).join('/');
}
/** ergänzt abhängige Werte; true = andere Felder geändert (neu zeichnen) */
function autoFill(rec, p, el) {
  const d = rec.data; let changed = false;
  if (p === 'identArt' && d.identArt) {
    changed = !d.identVon || !d.identAm || !d.identPerson;
    if (!d.identVon) d.identVon = S.user.name;
    if (!d.identAm) d.identAm = LG.todayIso();
    if (!d.identPerson) {
      if (rec.typ === 'gmbh' && Array.isArray(d.gf) && d.gf[0]) d.identPerson = [d.gf[0].vorname, d.gf[0].nachname].filter(Boolean).join(' ');
      else if (rec.typ !== 'gmbh') d.identPerson = [d.titel, d.vorname, d.nachname].filter(Boolean).join(' ');
    }
  }
  if (p === 'iban' && d.iban && LG.validIban(d.iban)) { d.iban = LG.fmtIban(d.iban); if (el) el.value = d.iban; }
  if (p === 'sepa' && d.sepa === 'erteilt' && !d.kontoinhaber) { d.kontoinhaber = rec.typ === 'gmbh' ? (d.firma || '') : [d.vorname, d.nachname].filter(Boolean).join(' '); changed = true; }
  return changed;
}

/* ================= Demo-Daten ================= */
function demoSeed() {
  const d = n => { const x = new Date(); x.setDate(x.getDate() - n); return x.toISOString(); };
  const day = n => d(n).slice(0, 10);
  const base = (id, typ, status, data, extra = {}) => {
    const rec = { id, typ, status, data, open: {}, docs: {}, history: [{ at: d(extra.age || 1), by: 'Xenia Peter', t: 'Aufnahme angelegt' }], createdAt: d(extra.age || 1), createdBy: 'Xenia Peter', updatedAt: d(extra.upd || 0), updatedBy: 'Xenia Peter', rev: 1, ...extra };
    delete rec.age; delete rec.upd;
    rec.title = LG.computeTitle(rec); rec.folder = `${rec.createdAt.slice(0, 10)} ${safeName(rec.title)} ${id.slice(-4)}`; rec.pct = LG.evaluate(rec).pct; return rec;
  };
  const a = base('demo1hofm', 'gmbh', 'entwurf', {
    firma: 'Hofmann Haustechnik GmbH', rechtsform: 'GmbH', strasse: 'Industriestraße 12', plz: '94315', ort: 'Straubing', registergericht: 'Amtsgericht Straubing',
    wj: 'Kalenderjahr', email: 't.hofmann@beispiel.de', branche: 'Sanitär, Heizung, Klima',
    gf: [{ anrede: 'Herr', vorname: 'Thomas', nachname: 'Hofmann', wohnort: 'Straubing', staatsang: 'deutsch', email: 't.hofmann@beispiel.de' }],
    gs: [{ name: 'Thomas Hofmann', anteil: '60', wb: 'ja', wohnort: 'Straubing', staatsang: 'deutsch' }, { name: 'Maria Hofmann', anteil: '40', wb: 'ja' }],
    finanzamt: 'Straubing', leistungen: ['Finanzbuchhaltung', 'Lohnabrechnung', 'Jahresabschluss', 'Betriebliche Steuererklärungen'], verguetung: 'Pauschalvereinbarung'
  }, { age: 0, upd: 0 });
  a.open = { hrnr: true, gruendung: true, 'doc:hr': true, 'doc:gv': true, 'doc:gsl': true, 'doc:ausweis': true };
  const b = base('demo2bran', 'freiberufler', 'mandant', {
    anrede: 'Frau', vorname: 'Laura', nachname: 'Brandl', geburtsdatum: '1991-04-17', geburtsort: 'Regensburg', staatsang: 'deutsch', beruf: 'Grafikdesignerin',
    strasse: 'Am Stadtpark 4', plz: '94315', ort: 'Straubing', email: 'laura.brandl@beispiel.de', telefon: '0171 0000000', familienstand: 'ledig', konfession: 'keine',
    taetigkeit: 'Grafikdesign und Illustration', art: 'freiberuflich (§ 18 EStG)', beginn: day(-30), umsatzLfd: '8.000', umsatzFolge: '42.000', kleinunternehmer: 'nein', versteuerung: 'Ist-Versteuerung',
    gewinnermittlung: 'Einnahmen-Überschuss-Rechnung', fragebogen: 'übernimmt ProTax', finanzamt: 'Straubing', einkuenfte: ['Selbständige Arbeit'],
    leistungen: ['Finanzbuchhaltung', 'Einkommensteuererklärung', 'Existenzgründungsberatung'], mandatAb: day(-30), verguetung: 'nach StBVV', auftrag: 'verschickt', vollmacht: 'offen'
  }, { age: 6, upd: 3 });
  b.open = { stid: true, iban: true, kontoinhaber: true, sepa: true, 'doc:ausweis': true, 'doc:nachweis': true, 'doc:bescheid': true };
  b.history.push({ at: d(3), by: 'Xenia Peter', t: 'E-Mail an Mandanten erstellt (7 offene Punkte)' }, { at: d(3), by: 'Xenia Peter', t: 'Status: Beim Mandanten' });
  const c = base('demo3wagn', 'privat', 'eingereicht', {
    anrede: 'Herr', titel: 'Dr.', vorname: 'Thomas', nachname: 'Wagner', geburtsdatum: '1968-09-02', geburtsort: 'München', staatsang: 'deutsch', beruf: 'Zahnarzt',
    strasse: 'Lindenweg 9', plz: '94377', ort: 'Steinach', email: 'dr.wagner@beispiel.de', familienstand: 'verheiratet', familienstandSeit: '1996-06-15', konfession: 'römisch-katholisch',
    p_vorname: 'Sabine', p_nachname: 'Wagner', p_geburtsdatum: '1970-03-11', p_konfession: 'römisch-katholisch', veranlagung: 'Zusammenveranlagung',
    kinder: [{ vorname: 'Lena', geburtsdatum: '2008-05-20', kindergeld: 'ja' }],
    stid: '86095742719', stnr: '162/123/45678', finanzamt: 'Straubing', einkuenfte: ['Nichtselbständige Arbeit', 'Kapitalvermögen', 'Vermietung und Verpachtung'],
    identPerson: 'Dr. Thomas Wagner', identArt: 'persönlich mit Originaldokument', ausweisArt: 'Personalausweis', ausweisNr: 'L01X00T47', behoerde: 'Stadt München', ausweisGueltig: '2031-03-14', pep: 'nein', risiko: 'normal', identAm: day(10), identVon: 'Xenia Peter',
    iban: 'DE89 3704 0044 0532 0130 00', kontoinhaber: 'Dr. Thomas Wagner', sepa: 'erteilt', sepaDatum: day(9),
    leistungen: ['Einkommensteuererklärung', 'Laufende Steuerberatung'], mandatAb: day(10), verguetung: 'nach StBVV', auftrag: 'unterschrieben', vollmacht: 'beantragt', datenschutz: true
  }, { age: 10, upd: 1 });
  c.docs = { ausweis: { state: 'papier' }, bescheid: { state: 'papier' }, auftragDoc: { state: 'papier' }, sepaDoc: { state: 'papier' } };
  c.history.push({ at: d(1), by: 'Xenia Peter', t: 'Status: Eingereicht' });
  [a, b, c].forEach(r => { r.pct = LG.evaluate(r).pct; });
  return [a, b, c];
}
