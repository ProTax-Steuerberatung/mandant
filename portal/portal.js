/* ProTax Mandantenportal – Mandant ergänzt angefragte Angaben und lädt Unterlagen hoch.
 * Alles wird hier im Browser verschlüsselt; der Server sieht nur unlesbare Daten. */
'use strict';
(function () {
  const C = window.CONFIG || {}, P = C.portal || {}, K = C.kanzlei || {};
  const SC = window.SCHEMA, LG = window.LOGIC, X = window.PTCRYPTO;
  const CHUNK = 2 * 1024 * 1024, MAX_MB = 40;
  const $ = s => document.querySelector(s);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const deDate = ts => new Date(ts * 1000).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const mb = n => n < 1048576 ? Math.max(1, Math.round(n / 1024)) + ' KB' : (n / 1048576).toLocaleString('de-DE', { maximumFractionDigits: 1 }) + ' MB';
  const I = {
    back: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><path d="M12.5 4.5 7 10l5.5 5.5"/></svg>',
    lock: '<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.4" aria-hidden="true"><rect x="3" y="6" width="8" height="6" rx="1"/><path d="M5 6V4.5a2 2 0 0 1 4 0V6"/></svg>',
    cam: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" aria-hidden="true"><path d="M3 8h4l2-3h6l2 3h4v11H3z"/><circle cx="12" cy="13" r="3.5"/></svg>',
    file: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" aria-hidden="true"><path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4"/></svg>',
    doc: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M5 2.5h7l3 3v12H5z"/><path d="M8 10h4M8 13h4"/></svg>',
    ok: '<svg viewBox="0 0 28 28" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M7 14.5 12 19.5 21 9"/></svg>',
    x: '<svg width="18" height="18" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><path d="M5 5l10 10M15 5 5 15"/></svg>'
  };
  // Hilfetexte für Mandanten (Kanzlei-Begriffe übersetzt)
  const HINT = {
    stid: '11 Ziffern – steht z. B. auf Ihrem Steuerbescheid oder Ihrer Lohnsteuerbescheinigung.',
    p_stid: '11 Ziffern – steht z. B. auf dem Steuerbescheid.',
    stnr: 'Steht auf Schreiben Ihres Finanzamts, z. B. 162/123/45678.',
    konfession: 'Wichtig für die Kirchensteuer.',
    iban: 'Für das SEPA-Lastschriftmandat unserer Honorarrechnungen.',
    ustid: 'Nur falls vorhanden, beginnt mit DE.',
    umsatzLfd: 'Eine grobe Schätzung genügt.', umsatzFolge: 'Eine grobe Schätzung genügt.'
  };
  // Felder, die erst durch eine Antwort des Mandanten nötig werden
  const DYN = {
    familienstand: ['familienstandSeit', 'p_vorname', 'p_nachname', 'p_geburtsdatum', 'p_stid', 'p_konfession', 'veranlagung'],
    wj: ['wjBeginn'], kleinunternehmer: ['versteuerung'], fragebogen: ['fragebogenDatum'],
    holding: ['beteiligungen'], rechtsform: ['registergericht', 'hrnr'], vorberater: []
  };
  const ROWNAME = { kinder: 'Kind', gf: 'Geschäftsführer/in', gs: 'Gesellschafter/in', beteiligungen: 'Beteiligung' };

  const S = { id: null, code: null, session: null, anfrage: null, stepKey: 'intro', answers: {}, groups: {}, files: [], names: {}, up: {}, pub: null, fp: null, note: '' };
  const KEY = () => 'ptportal:' + S.id;
  function save() {
    try { sessionStorage.setItem(KEY(), JSON.stringify({ session: S.session, anfrage: S.anfrage, stepKey: S.stepKey, answers: S.answers, groups: S.groups, files: S.files, names: S.names })); } catch (e) { }
  }
  function restore() { try { return JSON.parse(sessionStorage.getItem(KEY()) || 'null'); } catch (e) { return null; } }
  function forget() { try { sessionStorage.removeItem(KEY()); } catch (e) { } }

  /* ---------------- Server ---------------- */
  async function api(r, body, params = '') {
    let res;
    try {
      res = await fetch(`${P.api}?r=${r}${params}`, { method: 'POST', headers: { 'Content-Type': 'application/json', ...(S.session ? { 'X-Session': S.session } : {}) }, body: JSON.stringify(body || {}) });
    } catch (e) { const err = new Error('Keine Verbindung. Bitte prüfen Sie Ihre Internetverbindung und versuchen Sie es erneut.'); err.code = 'netz'; throw err; }
    let j = {}; try { j = await res.json(); } catch (e) { }
    if (!res.ok || !j.ok) { const err = new Error(j.error || 'Es ist ein Fehler aufgetreten.'); err.code = j.code || ''; err.status = res.status; throw err; }
    return j;
  }

  /* ---------------- Schritte ---------------- */
  const merged = () => ({ ...(S.anfrage.ctx || {}), ...S.answers });
  function eligible(f) {
    const A = S.anfrage, req = A.felder.includes(f.k) || A.gruppen.includes(f.k);
    const dyn = Object.entries(DYN).some(([dep, list]) => list.includes(f.k) && A.felder.includes(dep));
    if (f.kanzlei || f.internal || !LG.appliesTo(f, A.typ) || !(req || dyn)) return false;
    if (!f.when) return true;
    try { return !!f.when(merged()); } catch (e) { return false; }
  }
  function steps() {
    const A = S.anfrage, out = [{ k: 'intro' }];
    for (const sec of LG.sectionsFor(A.typ)) {
      if (sec.docs) continue;
      const fields = (sec.fields || []).filter(eligible);
      if (fields.length) out.push({ k: sec.k, title: sec.l, fields });
    }
    if (A.docs.length) out.push({ k: 'docs', title: 'Unterlagen' });
    out.push({ k: 'summary', title: 'Übersicht' });
    return out;
  }
  const docList = () => SC.DOCS.filter(d => S.anfrage.docs.includes(d.k));

  /* ---------------- Start ---------------- */
  async function start() {
    renderFoot();
    const m = location.hash.match(/a=([a-f0-9]{32})/);
    if (!m) return msg('Bitte öffnen Sie den Link aus unserer E-Mail', 'Diese Seite ist nur über den persönlichen Link erreichbar, den Sie von uns per E-Mail erhalten haben.');
    S.id = m[1];
    if (!window.isSecureContext || !X.available()) return msg('Ihr Browser wird nicht unterstützt', 'Bitte öffnen Sie den Link in einem aktuellen Browser (z. B. Safari, Chrome, Edge oder Firefox).');
    if (!P.api || !P.publicKey) return msg('Das Portal ist gerade nicht verfügbar', 'Bitte versuchen Sie es später erneut oder wenden Sie sich an die Kanzlei.');
    try { S.pub = await X.importPublic(P.publicKey); S.fp = await X.fingerprint(P.publicKey); }
    catch (e) { return msg('Das Portal ist gerade nicht verfügbar', 'Bitte wenden Sie sich an die Kanzlei.'); }
    const r = restore();
    if (r && r.anfrage) { Object.assign(S, { session: r.session, anfrage: r.anfrage, stepKey: r.stepKey || 'intro', answers: r.answers || {}, groups: r.groups || {}, files: r.files || [], names: r.names || {} }); }
    if (S.session && S.anfrage) return render();
    renderCode();
  }

  function msg(title, text, extra = '') {
    $('#app').innerHTML = `<div class="rule"></div><h1>${esc(title)}</h1><p class="lead">${esc(text)}</p>${extra}${contactHtml()}`;
  }
  function contactHtml() {
    return `<div class="card small">Fragen? ${esc(K.name || '')}<br>${K.email ? `E-Mail: <a href="mailto:${esc(K.email)}">${esc(K.email)}</a>` : ''}${K.telefon ? ` · Telefon: <a href="tel:${esc(K.telefon.replace(/\s/g, ''))}">${esc(K.telefon)}</a>` : ''}</div>`;
  }
  function renderFoot() {
    $('#foot').innerHTML = `<span>${esc(K.name || '')}${K.adresse ? ' · ' + esc(K.adresse) : ''}</span>${P.datenschutz ? `<a href="${esc(P.datenschutz)}" target="_blank" rel="noopener">Datenschutz</a>` : ''}${P.impressum ? `<a href="${esc(P.impressum)}" target="_blank" rel="noopener">Impressum</a>` : ''}`;
  }

  /* ---------------- Code eingeben ---------------- */
  function renderCode(note) {
    $('#app').innerHTML = `<div class="rule"></div><h1>Willkommen im Mandantenportal</h1>
      <p class="lead">Bitte geben Sie Ihren 6-stelligen Zugangscode ein. Sie haben ihn von uns im Gespräch, per SMS oder in einer separaten E-Mail erhalten.</p>
      ${note ? `<div class="note warn">${esc(note)}</div>` : ''}
      <form id="fcode" class="card stack" novalidate>
        <div class="f"><label for="code">Zugangscode</label>
          <input id="code" class="code" inputmode="numeric" autocomplete="one-time-code" maxlength="7" placeholder="000000" required></div>
        <div id="cerr" class="note err" hidden></div>
        <button class="btn" type="submit" id="cbtn">Weiter</button>
      </form>
      <p class="lock">${I.lock}Ihre Angaben werden auf Ihrem Gerät verschlüsselt und können nur von der Kanzlei gelesen werden.</p>`;
    const inp = $('#code'); inp.focus();
    $('#fcode').addEventListener('submit', async e => {
      e.preventDefault();
      const code = inp.value.replace(/\D/g, ''), err = $('#cerr'), btn = $('#cbtn');
      if (code.length !== 6) { err.textContent = 'Bitte geben Sie alle 6 Ziffern ein.'; err.hidden = false; return; }
      btn.disabled = true; btn.textContent = 'Prüfe …'; err.hidden = true;
      try {
        const r = await api('zugang', { id: S.id, code });
        const anfrage = await X.openWithCode(code, S.id, r.anfrage);
        if (anfrage.fp !== S.fp) throw Object.assign(new Error('Sicherheitsprüfung fehlgeschlagen. Bitte wenden Sie sich an die Kanzlei.'), { code: 'fp' });
        S.session = r.sitzung; S.anfrage = anfrage; S.code = code;
        S.files = (r.dateien || []).map(f => ({ fid: f.fid, doc: f.doc, size: f.size }));
        save(); render();
      } catch (ex) {
        if (['eingereicht', 'abgelaufen', 'gesperrt'].includes(ex.code) && ex.status !== 403) {
          forget();
          return msg(ex.code === 'eingereicht' ? 'Bereits eingereicht' : ex.code === 'abgelaufen' ? 'Link abgelaufen' : 'Zugang gesperrt', ex.message);
        }
        err.textContent = ex.message; err.hidden = false; btn.disabled = false; btn.textContent = 'Weiter'; inp.select();
      }
    });
  }

  /* ---------------- Rahmen eines Schritts ---------------- */
  function render() {
    const list = steps();
    let i = list.findIndex(s => s.k === S.stepKey); if (i < 0) { i = 0; S.stepKey = 'intro'; }
    const st = list[i], n = list.length - 1;
    const head = i === 0 ? '' : `<div class="progress"><div class="row"><button class="back" data-go="${list[i - 1].k}">${I.back}Zurück</button><span>Schritt ${i} von ${n}</span><span style="width:60px"></span></div>
      <div class="bars" style="grid-template-columns:repeat(${n},minmax(0,1fr))" aria-hidden="true">${list.slice(1).map((s, j) => `<span class="${j + 1 < i ? 'done' : j + 1 === i ? 'cur' : ''}"></span>`).join('')}</div></div>`;
    let body = '';
    if (st.k === 'intro') body = introHtml(list);
    else if (st.k === 'docs') body = docsHtml();
    else if (st.k === 'summary') body = summaryHtml(list);
    else body = sectionHtml(st);
    const next = list[i + 1];
    const foot = st.k === 'summary' ? '' : `<div class="actions"><button class="btn" data-next="${next ? next.k : ''}" id="next">${st.k === 'intro' ? 'Los geht’s' : 'Weiter'}</button>
      <p class="lock">${I.lock}Zwischenstand bleibt erhalten, solange dieses Fenster geöffnet ist.</p></div>`;
    $('#app').innerHTML = head + body + foot;
    bind(st);
    window.scrollTo(0, 0);
  }

  function introHtml(list) {
    const A = S.anfrage, docs = docList(), nf = list.reduce((a, s) => a + (s.fields ? s.fields.length : 0), 0);
    const min = Math.max(5, Math.min(30, Math.round(nf * 0.6 + docs.length * 1.5 + 2)));
    const prep = docs.map(d => d.l);
    if (A.felder.includes('stid')) prep.push('Ihre Steuer-Identifikationsnummer');
    if (A.felder.includes('iban')) prep.push('Ihre Bankverbindung (IBAN)');
    return `<div class="rule"></div><h1>${A.anrede ? 'Willkommen, ' + esc(A.anrede) : 'Willkommen'}</h1>
      <p class="lead">Für Ihre Mandatsaufnahme bei ${esc(K.name || 'uns')} benötigen wir noch einige Angaben${docs.length ? ' und Unterlagen' : ''}. Das dauert etwa ${min} Minuten.</p>
      ${prep.length ? `<div class="card stack"><h2>Bitte halten Sie bereit</h2><ul class="check-list">${prep.map(t => `<li>${I.doc}<span>${esc(t)}</span></li>`).join('')}</ul>
        <p class="small" style="margin:0">Fotos können Sie direkt mit dem Smartphone aufnehmen. Fehlt etwas, können Sie trotzdem einreichen – wir klären den Rest mit Ihnen.</p></div>` : ''}
      <p class="small">Der Link ist gültig bis ${esc(deDate(A.expires))}. ${A.kontakt ? 'Ihre Ansprechpartnerin bzw. Ihr Ansprechpartner: ' + esc(A.kontakt) + '.' : ''}</p>`;
  }

  /* ---------------- Felder ---------------- */
  function valOf(path) { const p = path.split('.'); if (p.length === 1) return S.answers[p[0]]; return ((S.groups[p[0]] || [])[+p[1]] || {})[p[2]]; }
  function setVal(path, v) {
    const p = path.split('.');
    if (p.length === 1) { if (v === '' || v == null || (Array.isArray(v) && !v.length) || v === false) delete S.answers[p[0]]; else S.answers[p[0]] = v; return; }
    const rows = S.groups[p[0]] = S.groups[p[0]] || []; rows[+p[1]] = rows[+p[1]] || {}; rows[+p[1]][p[2]] = v;
  }
  const fid = path => 'q-' + path.replace(/\./g, '-');
  function inputHtml(f, path) {
    const v = valOf(path) ?? '', id = fid(path);
    const ph = f.ph ? ` placeholder="${esc(f.ph)}"` : '';
    switch (f.t) {
      case 'select': {
        const opts = f.o.filter(o => o !== 'ohne Anrede' || f.k === 'anrede');
        if (opts.length <= 4) return `<div class="pills" role="radiogroup" aria-labelledby="${id}-l">${opts.map(o => `<label class="pill"><input type="radio" name="${id}" data-p="${path}" value="${esc(o)}" ${o === v ? 'checked' : ''}>${esc(o)}</label>`).join('')}</div>`;
        return `<select id="${id}" data-p="${path}"><option value="">Bitte wählen</option>${opts.map(o => `<option ${o === v ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
      }
      case 'multi': { const arr = Array.isArray(v) ? v : []; return `<div class="multi" role="group" aria-labelledby="${id}-l">${f.o.map(o => `<label class="chk"><input type="checkbox" data-p="${path}" data-multi="${esc(o)}" ${arr.includes(o) ? 'checked' : ''}>${esc(o)}</label>`).join('')}</div>`; }
      case 'textarea': return `<textarea id="${id}" data-p="${path}" rows="3"${ph}>${esc(v)}</textarea>`;
      case 'date': return `<input id="${id}" type="date" data-p="${path}" value="${esc(v)}" max="9999-12-31">`;
      case 'email': return `<input id="${id}" type="email" inputmode="email" autocomplete="email" data-p="${path}" value="${esc(v)}"${ph}>`;
      case 'tel': return `<input id="${id}" type="tel" autocomplete="tel" data-p="${path}" value="${esc(v)}"${ph}>`;
      case 'num': return `<input id="${id}" inputmode="decimal" data-p="${path}" value="${esc(v)}"${ph}>`;
      case 'stid': return `<input id="${id}" inputmode="numeric" data-p="${path}" value="${esc(v)}" placeholder="z. B. 12 345 678 901" autocomplete="off">`;
      case 'plz': return `<input id="${id}" inputmode="numeric" autocomplete="postal-code" data-p="${path}" value="${esc(v)}" maxlength="10">`;
      case 'iban': return `<input id="${id}" data-p="${path}" value="${esc(v)}" placeholder="DE00 0000 0000 0000 0000 00" autocomplete="off" autocapitalize="characters" spellcheck="false">`;
      default: return `<input id="${id}" data-p="${path}" value="${esc(v)}"${ph} autocomplete="off">`;
    }
  }
  function fieldHtml(f, path) {
    const id = fid(path), err = LG.checkValue(f, valOf(path), merged());
    if (f.t === 'check') return `<label class="chk"><input type="checkbox" id="${id}" data-p="${path}" ${valOf(path) ? 'checked' : ''}>${esc(f.l)}</label>`;
    const lab = f.t === 'select' && f.o.length <= 4 || f.t === 'multi' ? `<span class="lbl" id="${id}-l">${esc(f.l)}</span>` : `<label for="${id}">${esc(f.l)}</label>`;
    const hint = HINT[f.k] || f.hint;
    return `<div class="f ${err ? 'bad' : ''}" data-fw="${path}">${lab}${hint ? `<span class="hint">${esc(hint)}</span>` : ''}${inputHtml(f, path)}<div class="err" id="${id}-e">${esc(err)}</div></div>`;
  }
  function groupHtml(g) {
    let rows = S.groups[g.k]; if (!rows || !rows.length) { rows = S.groups[g.k] = [{}]; }
    const nm = ROWNAME[g.k] || g.l;
    return `<div class="stack"><span class="lbl" style="font:600 15px var(--sans);color:var(--ink)">${esc(g.l)}</span>
      ${rows.map((row, i) => `<div class="group"><div class="group-h"><span>${esc(nm)} ${i + 1}</span>${rows.length > 1 || g.k === 'kinder' ? `<button type="button" class="btn sm danger" data-gdel="${g.k}.${i}">Entfernen</button>` : ''}</div>
        ${LG.rowFields(g, row, merged()).map(f => fieldHtml(f, `${g.k}.${i}.${f.k}`)).join('')}</div>`).join('')}
      <button type="button" class="btn ghost" data-gadd="${g.k}">+ ${esc(nm)} hinzufügen</button></div>`;
  }
  function sectionHtml(st) {
    return `<h1>${esc(st.title)}</h1><div class="card stack">${st.fields.map(f => f.t === 'group' ? groupHtml(f) : fieldHtml(f, f.k)).join('')}</div>`;
  }
  function stepErrors(st) {
    const errs = [];
    for (const f of st.fields || []) {
      if (f.t === 'group') (S.groups[f.k] || []).forEach((row, i) => LG.rowFields(f, row, merged()).forEach(gf => { const e = LG.checkValue(gf, row[gf.k], merged()); if (e) errs.push(fid(`${f.k}.${i}.${gf.k}`)); }));
      else if (LG.checkValue(f, S.answers[f.k], merged())) errs.push(fid(f.k));
    }
    return errs;
  }

  /* ---------------- Unterlagen ---------------- */
  function docsHtml() {
    return `<h1>Unterlagen</h1><p class="lead">Bitte fotografieren Sie die Unterlagen oder wählen Sie eine Datei (PDF oder Foto).</p>
      <div class="stack">${docList().map(d => {
        const fs = S.files.filter(f => f.doc === d.k), up = Object.values(S.up).filter(u => u.doc === d.k);
        return `<div class="doc ${fs.length ? 'has' : ''}"><div class="doc-h"><span class="doc-t">${esc(d.l)}</span><span class="doc-s">${fs.length ? fs.length + ' Datei' + (fs.length > 1 ? 'en' : '') : 'fehlt'}</span></div>
          ${fs.length || up.length ? `<div class="files">${fs.map(f => `<div class="file"><span class="n">${esc(S.names[f.fid] || 'Datei')}</span><span class="sz">${mb(f.size)}</span><button type="button" data-del="${f.fid}" aria-label="Datei entfernen">${I.x}</button></div>`).join('')}
            ${up.map(u => `<div class="file" style="flex-direction:column;align-items:stretch;gap:6px;padding:10px 12px"><span class="n">${esc(u.name)} – ${u.err ? `<b style="color:var(--err)">${esc(u.err)}</b>` : 'wird verschlüsselt hochgeladen …'}</span>${u.err ? '' : `<div class="bar"><span style="width:${u.pct}%"></span></div>`}</div>`).join('')}</div>` : ''}
          <div class="doc-btns"><span class="btn ghost sm up">${I.cam}Foto<input type="file" accept="image/*" capture="environment" data-up="${d.k}" aria-label="Foto aufnehmen: ${esc(d.l)}"></span>
            <span class="btn ghost sm up">${I.file}Datei<input type="file" accept="application/pdf,image/*" multiple data-up="${d.k}" aria-label="Datei wählen: ${esc(d.l)}"></span></div></div>`;
      }).join('')}</div>`;
  }
  async function shrink(file) {
    if (!/^image\/(jpeg|png|webp)$/.test(file.type) || typeof createImageBitmap !== 'function') return file;
    try {
      const bmp = await createImageBitmap(file), max = 2400, sc = Math.min(1, max / Math.max(bmp.width, bmp.height));
      if (sc === 1 && file.size < 2.5 * 1048576) return file;
      const cv = document.createElement('canvas'); cv.width = Math.round(bmp.width * sc); cv.height = Math.round(bmp.height * sc);
      cv.getContext('2d').drawImage(bmp, 0, 0, cv.width, cv.height);
      const blob = await new Promise(r => cv.toBlob(r, 'image/jpeg', 0.85));
      if (!blob || blob.size >= file.size) return file;
      return new File([blob], file.name.replace(/\.\w+$/, '') + '.jpg', { type: 'image/jpeg' });
    } catch (e) { return file; }
  }
  async function upload(doc, original) {
    const key = X.randomFid(), u = S.up[key] = { doc, name: original.name || 'Foto', pct: 0 };
    rerender();
    try {
      const file = await shrink(original);
      if (file.size > MAX_MB * 1048576) throw new Error(`Datei zu groß (max. ${MAX_MB} MB)`);
      if (!/^(application\/pdf|image\/)/.test(file.type || '') && !/\.(pdf|jpe?g|png|heic|heif|webp)$/i.test(file.name)) throw new Error('Bitte nur PDF oder Fotos hochladen');
      const bytes = new Uint8Array(await file.arrayBuffer());
      const { key: k, fk } = await X.newDataKey(S.pub);
      const meta = await X.encText(k, JSON.stringify({ name: file.name, type: file.type, size: file.size }));
      const ct = await X.encBytes(k, bytes);
      const fidv = X.randomFid(), parts = Math.max(1, Math.ceil(ct.length / CHUNK));
      for (let i = 0; i < parts; i++) {
        const q = `&id=${S.id}&fid=${fidv}&doc=${encodeURIComponent(doc)}&part=${i}${i === parts - 1 ? '&final=1' : ''}${i === 0 ? `&meta=${encodeURIComponent(meta)}&fk=${encodeURIComponent(fk)}` : ''}`;
        let res;
        try { res = await fetch(`${P.api}?r=upload${q}`, { method: 'POST', headers: { 'Content-Type': 'application/octet-stream', 'X-Session': S.session }, body: ct.subarray(i * CHUNK, (i + 1) * CHUNK) }); }
        catch (e) { throw new Error('Verbindung unterbrochen – bitte erneut versuchen'); }
        let j = {}; try { j = await res.json(); } catch (e) { }
        if (!res.ok || !j.ok) { const err = new Error(j.error || 'Upload fehlgeschlagen'); err.code = j.code; throw err; }
        u.pct = Math.round((i + 1) / parts * 100); rerender();
      }
      S.files.push({ fid: fidv, doc, size: ct.length }); S.names[fidv] = file.name;
      delete S.up[key]; save(); rerender();
    } catch (e) {
      if (e.code === 'sitzung') return sessionLost();
      u.err = e.message; rerender();
      setTimeout(() => { delete S.up[key]; rerender(); }, 6000);
    }
  }
  function rerender() { if (S.stepKey === 'docs') { const y = window.scrollY; render(); window.scrollTo(0, y); } }

  /* ---------------- Übersicht & Einreichen ---------------- */
  function fmt(f, v) {
    if (v == null || v === '' || (Array.isArray(v) && !v.length)) return '';
    if (f.t === 'date') return LG.deDate(v); if (f.t === 'multi') return v.join(', '); if (f.t === 'check') return v ? 'ja' : '';
    if (f.t === 'iban') return LG.fmtIban(v); return String(v);
  }
  function summaryHtml(list) {
    let missing = 0;
    const secs = list.filter(s => s.fields).map(s => `<div class="card"><h2 style="font-size:19px;margin-bottom:10px">${esc(s.title)}</h2><dl class="sum">${s.fields.map(f => {
      if (f.t === 'group') {
        const rows = (S.groups[f.k] || []).filter(r => Object.values(r).some(v => v !== '' && v != null));
        if (!rows.length && f.k !== 'kinder') missing++;
        return `<dt>${esc(f.l)}</dt><dd class="${rows.length ? '' : 'empty'}">${rows.length ? rows.map((r, i) => esc((ROWNAME[f.k] || '') + ' ' + (i + 1) + ': ' + f.fields.map(x => fmt(x, r[x.k])).filter(Boolean).join(', '))).join('<br>') : (f.k === 'kinder' ? 'keine' : 'nicht angegeben')}</dd>`;
      }
      const t = fmt(f, S.answers[f.k]); if (!t) missing++;
      return `<dt>${esc(f.l)}</dt><dd class="${t ? '' : 'empty'}">${t ? esc(t) : 'nicht angegeben'}</dd>`;
    }).join('')}</dl><button class="btn ghost sm" data-go="${s.k}">Ändern</button></div>`).join('');
    const docs = docList().map(d => { const n = S.files.filter(f => f.doc === d.k).length; if (!n) missing++; return `<dt>${esc(d.l)}</dt><dd class="${n ? '' : 'empty'}">${n ? n + ' Datei' + (n > 1 ? 'en' : '') : 'nicht hochgeladen'}</dd>`; }).join('');
    return `<h1>Übersicht</h1><p class="lead">Bitte prüfen Sie Ihre Angaben und senden Sie sie dann verschlüsselt an uns.</p>
      ${missing ? `<div class="note warn">${missing} Punkt(e) sind noch offen. Sie können trotzdem einreichen – wir melden uns dann wegen der fehlenden Angaben bei Ihnen.</div>` : '<div class="note ok">Alle angefragten Punkte sind ausgefüllt.</div>'}
      ${secs}${docs ? `<div class="card"><h2 style="font-size:19px;margin-bottom:10px">Unterlagen</h2><dl class="sum">${docs}</dl><button class="btn ghost sm" data-go="docs">Ändern</button></div>` : ''}
      <div class="card stack"><label class="chk"><input type="checkbox" id="confirm">Ich bestätige, dass meine Angaben nach bestem Wissen richtig sind.</label>
        <p class="small" style="margin:0">Mit dem Absenden werden Ihre Angaben und Unterlagen verschlüsselt an ${esc(K.name || 'die Kanzlei')} übermittelt.${P.datenschutz ? ` Hinweise zum Datenschutz finden Sie <a href="${esc(P.datenschutz)}" target="_blank" rel="noopener">hier</a>.` : ''}</p>
        <div id="serr" class="note err" hidden></div>
        <button class="btn" id="submit" disabled>Verschlüsselt an die Kanzlei senden</button></div>`;
  }
  async function submit(btn) {
    const err = $('#serr'); btn.disabled = true; btn.textContent = 'Wird verschlüsselt und gesendet …'; err.hidden = true;
    try {
      const list = steps(), values = {}, groups = {};
      list.filter(s => s.fields).forEach(s => s.fields.forEach(f => {
        if (f.t === 'group') { const rows = (S.groups[f.k] || []).filter(r => Object.values(r).some(v => v !== '' && v != null)); if (rows.length) groups[f.k] = rows; }
        else if (S.answers[f.k] != null && S.answers[f.k] !== '') values[f.k] = S.answers[f.k];
      }));
      const payload = { v: 1, values, groups, files: S.files.map(f => ({ fid: f.fid, doc: f.doc })), at: new Date().toISOString() };
      const { key, fk } = await X.newDataKey(S.pub);
      const ct = await X.encText(key, JSON.stringify(payload));
      await api('einreichen', { id: S.id, antwort: { v: 1, fp: S.fp, wk: fk, ct } });
      const name = S.anfrage.anrede; forget();
      $('#app').innerHTML = `<div class="badge-ok">${I.ok}</div><h1>Vielen Dank${name ? ', ' + esc(name) : ''}</h1>
        <p class="lead">Ihre Angaben${S.files.length ? ' und Unterlagen' : ''} sind verschlüsselt bei uns eingegangen. Wir prüfen alles und melden uns bei Ihnen.</p>
        <p class="small">Sie können dieses Fenster jetzt schließen. Der Link ist damit nicht mehr gültig.</p>${contactHtml()}`;
      window.scrollTo(0, 0);
    } catch (e) {
      if (e.code === 'sitzung') return sessionLost();
      err.textContent = e.message; err.hidden = false; btn.disabled = false; btn.textContent = 'Verschlüsselt an die Kanzlei senden';
    }
  }
  function sessionLost() { S.session = null; save(); renderCode('Ihre Sitzung ist abgelaufen. Bitte geben Sie Ihren Code erneut ein – Ihre Eingaben bleiben erhalten.'); }

  /* ---------------- Ereignisse ---------------- */
  function bind(st) {
    const app = $('#app');
    app.onclick = async e => {
      const t = e.target.closest('button'); if (!t) return;
      if (t.dataset.go) { S.stepKey = t.dataset.go; save(); return render(); }
      if (t.id === 'next') {
        const errs = stepErrors(st);
        if (errs.length) { const el = document.getElementById(errs[0]); if (el) { el.scrollIntoView({ block: 'center' }); el.focus(); } render(); return; }
        S.stepKey = t.dataset.next || 'summary'; save(); return render();
      }
      if (t.dataset.gadd) { (S.groups[t.dataset.gadd] = S.groups[t.dataset.gadd] || []).push({}); save(); return render(); }
      if (t.dataset.gdel) { const [g, i] = t.dataset.gdel.split('.'); S.groups[g].splice(+i, 1); save(); return render(); }
      if (t.dataset.del) {
        t.disabled = true;
        try { await api('datei-weg', { id: S.id, fid: t.dataset.del }); S.files = S.files.filter(f => f.fid !== t.dataset.del); delete S.names[t.dataset.del]; save(); }
        catch (ex) { if (ex.code === 'sitzung') return sessionLost(); alertNote(ex.message); }
        return rerender();
      }
      if (t.id === 'submit') return submit(t);
    };
    app.oninput = e => {
      const el = e.target; if (!el.dataset.p || el.type === 'checkbox' || el.type === 'radio') return;
      setVal(el.dataset.p, el.value); save();
    };
    app.onchange = e => {
      const el = e.target;
      if (el.id === 'confirm') { $('#submit').disabled = !el.checked; return; }
      if (el.dataset.up) { const files = [...el.files]; el.value = ''; files.forEach(f => upload(el.dataset.up, f)); return; }
      if (!el.dataset.p) return;
      const p = el.dataset.p, before = JSON.stringify(steps().map(s => s.k + ':' + (s.fields || []).map(f => f.k).join(',')));
      if (el.dataset.multi) { const arr = Array.isArray(valOf(p)) ? [...valOf(p)] : []; const o = el.dataset.multi, i = arr.indexOf(o); if (el.checked && i < 0) arr.push(o); if (!el.checked && i >= 0) arr.splice(i, 1); setVal(p, arr); }
      else if (el.type === 'checkbox') setVal(p, el.checked);
      else { let v = el.value; if (el.dataset.p.endsWith('iban') && LG.validIban(v)) { v = LG.fmtIban(v); el.value = v; } setVal(p, v); }
      save();
      const after = JSON.stringify(steps().map(s => s.k + ':' + (s.fields || []).map(f => f.k).join(',')));
      if (after !== before || el.type === 'radio' || el.tagName === 'SELECT') { const y = window.scrollY; render(); window.scrollTo(0, y); return; }
      showErr(p);
    };
    app.onfocusout = e => { if (e.target.dataset && e.target.dataset.p) showErr(e.target.dataset.p); };
  }
  function fieldOf(path) {
    const p = path.split('.'), all = steps().flatMap(s => s.fields || []);
    const f = all.find(x => x.k === p[0]); if (!f) return null;
    return p.length === 1 ? f : (f.fields || []).find(x => x.k === p[2]);
  }
  function showErr(path) {
    const f = fieldOf(path), w = document.querySelector(`[data-fw="${CSS.escape(path)}"]`); if (!f || !w) return;
    const e = LG.checkValue(f, valOf(path), merged());
    w.classList.toggle('bad', !!e); const el = document.getElementById(fid(path) + '-e'); if (el) el.textContent = e;
  }
  function alertNote(t) { const n = document.createElement('div'); n.className = 'note err'; n.textContent = t; $('#app').prepend(n); setTimeout(() => n.remove(), 6000); }

  start();
})();
