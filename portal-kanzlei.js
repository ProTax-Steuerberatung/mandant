/* ProTax Mandant – Kanzlei-Seite des Mandantenportals
 * Link + Code erzeugen, Status abfragen, eingereichte Daten abholen, entschlüsseln,
 * nach SharePoint übernehmen und auf dem Rückkanal löschen.
 * Nutzt Funktionen aus app.js (S, cur, render, toast, openDlg …) erst zur Laufzeit. */
'use strict';
window.PORTAL = (function () {
  const P = () => (window.CONFIG || {}).portal || {};
  const X = () => window.PTCRYPTO;
  let keyCache = null, infoCache = null;

  /* ---------- Rückkanal ---------- */
  async function token() { if (window.__PT_TEST_TOKEN) return window.__PT_TEST_TOKEN; if (C.demo) throw new Error('Im Demo-Modus ist das Portal nicht verfügbar.'); return Auth.idToken(); }
  async function call(r, body, params = '', raw = false) {
    const t = await token();
    let res;
    try {
      res = await fetch(`${P().api}?r=${r}${params}`, body !== undefined
        ? { method: 'POST', headers: { Authorization: 'Bearer ' + t, 'Content-Type': 'application/json' }, body: JSON.stringify(body) }
        : { headers: { Authorization: 'Bearer ' + t } });
    } catch (e) { throw new Error('Rückkanal nicht erreichbar (' + P().api + '). Ist upload.protax.de eingerichtet?'); }
    if (raw && res.ok) return res.arrayBuffer();
    let j = {}; try { j = await res.json(); } catch (e) { }
    if (!res.ok || !j.ok) { const err = new Error(j.error || ('Rückkanal: HTTP ' + res.status)); err.status = res.status; throw err; }
    return j;
  }
  async function info() { if (!infoCache) infoCache = await call('info', {}); return infoCache; }

  /* ---------- Schlüssel ---------- */
  async function loadKey() {
    if (keyCache) return keyCache;
    const k = await S.store.getPortalKey();
    if (!k) return null;
    keyCache = { ...k, priv: await X().importPrivate(k.privateKey), fp: await X().fingerprint(k.publicKey) };
    return keyCache;
  }
  async function keyState() {
    const cfg = P().publicKey, k = await loadKey();
    if (!k) return { state: cfg ? 'fehlt' : 'neu' };
    if (!cfg) return { state: 'eintragen', key: k };
    if ((await X().fingerprint(cfg)) !== k.fp) return { state: 'abweichend', key: k };
    return { state: 'ok', key: k };
  }
  async function setupKey() {
    const kp = await X().generateKeyPair();
    const obj = { v: 1, created: new Date().toISOString(), createdBy: S.user.name, publicKey: kp.publicKey, privateKey: kp.privateKey };
    await S.store.putPortalKey(obj);
    keyCache = null; return loadKey();
  }

  /* ---------- Hilfen ---------- */
  const unix = () => Math.floor(Date.now() / 1000);
  const deD = ts => new Date(ts * 1000).toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const link = id => (P().url || '') + '#a=' + id;
  const fmtCode = c => String(c).replace(/(\d{3})(\d{3})/, '$1 $2');
  function anrede(rec) {
    const d = rec.data || {};
    let a = d.anrede, t = d.titel, n = d.nachname;
    if (rec.typ === 'gmbh' && Array.isArray(d.gf) && d.gf[0]) { a = d.gf[0].anrede; t = ''; n = d.gf[0].nachname; }
    if (!n) return '';
    return (a === 'Herr' || a === 'Frau' ? a + ' ' : '') + [t, n].filter(Boolean).join(' ');
  }
  function request(rec) {
    const ev = LG.evaluate(rec), felder = [], gruppen = [], docs = [];
    for (const o of ev.openList) {
      if (o.kind === 'doc') docs.push(o.key.slice(4));
      else { const f = fieldBy(rec, o.key); if (f && f.t === 'group') gruppen.push(o.key); else felder.push(o.key); }
    }
    const d = rec.data || {};
    const ctx = {}; ['familienstand', 'rechtsform', 'wj', 'holding', 'kleinunternehmer', 'fragebogen', 'pep'].forEach(k => { if (d[k] != null && d[k] !== '') ctx[k] = d[k]; });
    return { felder, gruppen, docs, ctx };
  }
  const html = t => '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">' + String(t).split('\n').map(l => esc(l).replace(/(https:\/\/[^\s<]+)/g, '<a href="$1">$1</a>')).join('<br>') + '</div>';
  function mailLink(rec, p, weg) {
    const K = C.kanzlei || {}, c = LG.contact(rec);
    const codeTxt = { gespraech: 'den wir Ihnen im Gespräch mitgeteilt haben', sms: 'den Sie per SMS erhalten', mail: 'den Sie in einer separaten E-Mail erhalten' }[weg];
    return {
      to: p.an, subject: `Ihre Mandatsaufnahme bei ${K.kurz || K.name} – sicheres Mandantenportal`,
      text: [c.gruss, '',
        'vielen Dank für das angenehme Gespräch und Ihr Vertrauen in unsere Kanzlei.', '',
        'Damit wir Ihr Mandat vollständig anlegen können, benötigen wir noch einige Angaben und Unterlagen. Bitte nutzen Sie dafür unser sicheres Mandantenportal:', '',
        link(p.id), '',
        `Zur Anmeldung benötigen Sie Ihren persönlichen 6-stelligen Zugangscode, ${codeTxt}. Der Link ist bis zum ${deD(p.expires)} gültig.`, '',
        'Unterlagen können Sie direkt mit dem Smartphone fotografieren oder als PDF hochladen. Ihre Angaben werden bereits auf Ihrem Gerät verschlüsselt und können nur von uns gelesen werden.', '',
        'Bei Fragen bin ich gerne für Sie da.', '', 'Mit freundlichen Grüßen', '', S.user.name, K.name || '',
        [K.adresse, K.telefon, K.email].filter(Boolean).join(' · ')].join('\n')
    };
  }
  function mailCode(rec, p, code) {
    const K = C.kanzlei || {}, c = LG.contact(rec);
    return { to: p.an, subject: 'Ihr Zugangscode', text: [c.gruss, '', `Ihr persönlicher Zugangscode lautet: ${fmtCode(code)}`, '',
      'Bitte geben Sie den Code nur auf der Seite ein, deren Link wir Ihnen in einer separaten E-Mail gesendet haben. Geben Sie den Code nicht weiter.', '',
      'Mit freundlichen Grüßen', '', S.user.name, K.name || ''].join('\n') };
  }
  async function draft(m) {
    const t = await S.store.mailToken();
    return S.store.createDraft(t, { to: m.to, subject: m.subject, html: html(m.text) });
  }

  /* ---------- Karte in der Seitenspalte ---------- */
  function cardHtml(rec) {
    const p = rec.portal;
    if (!p) return '';
    const st = p.status === 'offen' && p.expires < unix() ? 'abgelaufen' : p.status;
    const weg = { gespraech: 'im Gespräch', sms: 'per SMS', mail: 'per separater E-Mail' }[p.weg] || '';
    let body = '';
    if (st === 'offen') body = `<div class="sm">Link aktiv bis <b>${deD(p.expires)}</b> · Code ${esc(weg)}${p.lastSeen ? `<br><span class="muted">Zuletzt geöffnet ${fmtDT(new Date(p.lastSeen * 1000).toISOString())}</span>` : '<br><span class="muted">Noch nicht geöffnet</span>'}${p.locked ? '<br><b style="color:var(--err)">Gesperrt nach falschen Codes</b>' : ''}</div>
      <div class="row"><button class="btn sm" data-act="portal-code">Code anzeigen</button><button class="btn sm" data-act="portal-copy">Link kopieren</button><button class="btn sm danger" data-act="portal-revoke">Widerrufen</button></div>`;
    else if (st === 'eingereicht') body = `<div class="banner" style="background:var(--ok-bg);color:var(--ok);border-color:var(--ok)">Der Mandant hat am ${fmtDT(new Date((p.submitted || unix()) * 1000).toISOString())} eingereicht.</div>
      <button class="btn primary" data-act="portal-fetch">Abholen und übernehmen</button>`;
    else body = `<div class="sm muted">${st === 'abgeholt' ? `Abgeholt am ${fmtDT(p.fetchedAt)}` : st === 'widerrufen' ? 'Link widerrufen' : st === 'abgelaufen' ? `Link abgelaufen am ${deD(p.expires)}` : 'Link nicht mehr vorhanden'}</div>`;
    return `<section class="card card-b" aria-label="Mandantenportal" style="display:flex;flex-direction:column;gap:10px"><span class="card-t">Mandantenportal</span>${body}</section>`;
  }

  /* ---------- Dialog: Link senden ---------- */
  async function dlgSend(rec) {
    const ev = LG.evaluate(rec);
    if (!P().api) return toast('Mandantenportal ist in config.js nicht eingerichtet (portal.api).', 'err');
    if (!X().available()) return toast('Verschlüsselung im Browser nicht verfügbar.', 'err');
    let ks;
    try { ks = await keyState(); } catch (e) { return toast('Portal-Schlüssel konnte nicht geladen werden: ' + e.message, 'err'); }
    if (ks.state !== 'ok') return dlgKey(ks);
    if (!ev.openList.length) return toast('Bitte zuerst mit „→ Mandant“ festlegen, was angefragt werden soll.', 'err');
    let sms = false; try { sms = !!(await info()).sms; } catch (e) { return toast(e.message, 'err'); }
    const c = LG.contact(rec), d = rec.data || {}, tel = d.telefon || ((d.gf || [])[0] || {}).telefon || '';
    const days = +(P().gueltigTage || 14);
    openDlg(`<form method="dialog" id="pform"><div class="dlg-h">Portal-Link an den Mandanten</div>
      <div class="dlg-b">
        <div class="banner" style="background:var(--bg2);border-color:var(--border2);color:var(--text)">Der Mandant ergänzt ${ev.openList.length} angefragte Punkt(e) im Portal und lädt Unterlagen hoch. Alles wird auf seinem Gerät verschlüsselt.</div>
        <div class="f"><label for="p-an">E-Mail des Mandanten</label><input id="p-an" type="email" value="${esc(c.email)}" required></div>
        <div class="f"><label for="p-tage">Link gültig</label><select id="p-tage">${[7, 14, 30].map(n => `<option value="${n}" ${n === days ? 'selected' : ''}>${n} Tage</option>`).join('')}</select></div>
        <div class="f"><span class="lbl">Zugangscode übermitteln</span><div class="typcards">
          <label class="typcard"><input type="radio" name="weg" value="gespraech" checked><span>Im Gespräch<br><span class="sm muted">mündlich / Zettel</span></span></label>
          <label class="typcard" ${sms ? '' : 'title="SMS ist auf dem Rückkanal nicht eingerichtet"'}><input type="radio" name="weg" value="sms" ${sms ? '' : 'disabled'}><span>Per SMS<br><span class="sm muted">${sms ? 'an Handynummer' : 'nicht eingerichtet'}</span></span></label>
          <label class="typcard"><input type="radio" name="weg" value="mail"><span>Separate E-Mail<br><span class="sm muted">zweiter Entwurf</span></span></label></div></div>
        <div class="f" id="p-telw" hidden><label for="p-tel">Handynummer</label><input id="p-tel" type="tel" value="${esc(tel)}" placeholder="0171 …"></div>
        <div class="hint">Es entsteht ein Outlook-Entwurf mit dem Link, den Sie prüfen und selbst senden.</div>
        <div id="p-err" class="banner err" hidden></div>
      </div>
      <div class="dlg-f"><button type="button" class="btn" data-close>Abbrechen</button><button type="button" class="btn primary" id="p-go">Link erstellen</button></div></form>`);
    const wegSel = () => document.querySelector('#pform input[name=weg]:checked').value;
    document.querySelectorAll('#pform input[name=weg]').forEach(r => r.addEventListener('change', () => { $('#p-telw').hidden = wegSel() !== 'sms'; }));
    $('#p-go').addEventListener('click', async e => {
      const btn = e.currentTarget, err = $('#p-err'), an = $('#p-an').value.trim(), weg = wegSel(), tage = +$('#p-tage').value, tel = $('#p-tel').value.trim();
      err.hidden = true;
      if (!LG.validEmail(an)) { err.textContent = 'Bitte eine gültige E-Mail-Adresse eintragen.'; err.hidden = false; return; }
      if (weg === 'sms' && tel.replace(/\D/g, '').length < 8) { err.textContent = 'Bitte eine Handynummer eintragen.'; err.hidden = false; return; }
      btn.disabled = true; btn.textContent = 'Erstelle …';
      try {
        const id = X().randomId(), code = X().randomCode(), expires = unix() + tage * 86400, rq = request(rec);
        const anfrage = { v: 1, fp: ks.key.fp, typ: rec.typ, anrede: anrede(rec), kontakt: S.user.name, expires, ...rq };
        const box = await X().sealWithCode(code, id, anfrage);
        const res = await call('auftrag', { id, code, tage, anfrage: box });
        if (rec.portal && rec.portal.status === 'offen') { try { await call('loeschen', { id: rec.portal.id }); } catch (x) { } }
        rec.portal = { id, code, weg, an, tel: weg === 'sms' ? tel : '', created: unix(), expires: res.expires || expires, status: 'offen', by: S.user.name, punkte: rq.felder.length + rq.gruppen.length + rq.docs.length };
        if (rec.status === 'entwurf') rec.status = 'mandant';
        touch(rec, `Portal-Link erstellt (gültig bis ${deD(rec.portal.expires)}, Code ${({ gespraech: 'im Gespräch', sms: 'per SMS', mail: 'per E-Mail' })[weg]})`);
        const p = rec.portal, notes = [];
        let web = null;
        try { web = await draft(mailLink(rec, p, weg)); } catch (x) { notes.push('Outlook-Entwurf konnte nicht erstellt werden: ' + x.message); }
        if (weg === 'mail') { try { await draft(mailCode(rec, p, code)); notes.push('Zweiter Entwurf mit dem Code liegt in Outlook – bitte erst nach der Link-E-Mail senden.'); } catch (x) { notes.push('Code-E-Mail konnte nicht erstellt werden: ' + x.message); } }
        if (weg === 'sms') {
          try { await call('sms', { an: tel, text: `${(C.kanzlei || {}).kurz || 'ProTax'}: Ihr Zugangscode für das Mandantenportal lautet ${code}. Gültig bis ${deD(p.expires)}.` }); notes.push('SMS mit dem Code wurde gesendet.'); }
          catch (x) { notes.push('SMS fehlgeschlagen: ' + x.message + ' – Code bitte telefonisch mitteilen.'); }
        }
        if (web) window.open(web, '_blank', 'noopener');
        render();
        dlgDone(p, code, weg, notes, !!web);
      } catch (x) {
        console.error(x); err.textContent = x.message; err.hidden = false; btn.disabled = false; btn.textContent = 'Link erstellen';
      }
    });
  }
  function dlgDone(p, code, weg, notes, drafted) {
    openDlg(`<form method="dialog"><div class="dlg-h">Portal-Link erstellt</div><div class="dlg-b">
      ${drafted ? '<div class="banner" style="background:var(--ok-bg);color:var(--ok);border-color:var(--ok)">Der Outlook-Entwurf mit dem Link ist geöffnet – bitte prüfen und senden.</div>' : ''}
      ${weg === 'gespraech' ? '<div class="sm">Bitte teilen Sie dem Mandanten diesen Code mit:</div>' : '<div class="sm">Zugangscode (zur Kontrolle):</div>'}
      <div style="font:500 40px/1.1 var(--font);letter-spacing:.18em;text-align:center;padding:10px 0;font-variant-numeric:tabular-nums">${fmtCode(code)}</div>
      ${notes.map(n => `<div class="sm">${esc(n)}</div>`).join('')}
      <div class="f"><label for="p-link">Link</label><input id="p-link" readonly value="${esc(link(p.id))}"></div>
      <div class="hint">Code und Link finden Sie später in der Aufnahme unter „Mandantenportal“.</div></div>
      <div class="dlg-f"><button type="button" class="btn" id="p-cp">Link kopieren</button><button type="button" class="btn primary" data-close>Fertig</button></div></form>`);
    $('#p-cp').addEventListener('click', () => copy(link(p.id)));
  }
  async function copy(t) { try { await navigator.clipboard.writeText(t); toast('Kopiert', 'ok'); } catch (e) { const i = $('#p-link'); if (i) { i.select(); document.execCommand('copy'); toast('Kopiert', 'ok'); } } }

  /* ---------- Dialog: Schlüssel einrichten ---------- */
  function dlgKey(ks) {
    const pub = ks.key ? JSON.stringify({ kty: ks.key.publicKey.kty, n: ks.key.publicKey.n, e: ks.key.publicKey.e }) : '';
    const text = {
      neu: 'Für das Portal braucht die Kanzlei einmalig ein Schlüsselpaar. Der private Schlüssel wird in SharePoint („ProTax Mandant“ → MA_Unterlagen → _System) abgelegt und verlässt Microsoft 365 nie.',
      fehlt: 'In config.js ist ein öffentlicher Schlüssel eingetragen, aber der private Schlüssel fehlt in SharePoint (MA_Unterlagen/_System/portal-schluessel.json). Ohne ihn können eingereichte Daten nicht geöffnet werden.',
      eintragen: 'Der Schlüssel ist angelegt. Jetzt muss der öffentliche Teil einmalig in config.js unter portal.publicKey eingetragen werden – erst dann kennt das Portal den Schlüssel.',
      abweichend: 'Der Schlüssel in SharePoint passt nicht zum öffentlichen Schlüssel in config.js. Bitte config.js prüfen.'
    }[ks.state];
    openDlg(`<form method="dialog"><div class="dlg-h">Mandantenportal einrichten</div><div class="dlg-b">
      <div class="${ks.state === 'neu' || ks.state === 'eintragen' ? 'hint' : 'banner err'}">${esc(text)}</div>
      ${ks.key ? `<div class="f"><label for="k-pub">Öffentlicher Schlüssel (ungefährlich, darf weitergegeben werden)</label><textarea id="k-pub" readonly style="font-family:ui-monospace,Consolas,monospace;font-size:11px;min-height:120px">publicKey: ${esc(pub)},</textarea></div>
        <div class="sm">Fingerabdruck: <b>${esc(ks.key.fp)}</b></div>` : ''}
      </div><div class="dlg-f"><button type="button" class="btn" data-close>Schließen</button>
      ${ks.state === 'neu' ? '<button type="button" class="btn primary" id="k-go">Schlüssel erzeugen</button>' : ks.key ? '<button type="button" class="btn primary" id="k-cp">Kopieren</button>' : ''}</div></form>`);
    const go = $('#k-go');
    if (go) go.addEventListener('click', async () => {
      go.disabled = true; go.textContent = 'Erzeuge …';
      try { const k = await setupKey(); dlgKey({ state: P().publicKey ? 'abweichend' : 'eintragen', key: k }); }
      catch (e) { toast('Schlüssel konnte nicht gespeichert werden: ' + e.message, 'err'); go.disabled = false; go.textContent = 'Schlüssel erzeugen'; }
    });
    const cp = $('#k-cp');
    if (cp) cp.addEventListener('click', async () => { const t = $('#k-pub'); try { await navigator.clipboard.writeText(t.value); } catch (e) { t.select(); document.execCommand('copy'); } toast('Kopiert', 'ok'); });
  }

  /* ---------- Status abgleichen ---------- */
  async function sync() {
    if (!P().api || (C.demo && !window.__PT_TEST_TOKEN)) return false;
    const open = S.recs.filter(r => r.portal && r.portal.status === 'offen');
    if (!open.length) return false;
    let res; try { res = await call('status', { ids: open.map(r => r.portal.id) }); } catch (e) { console.warn('Portal-Status:', e.message); return false; }
    let changed = false;
    for (const rec of open) {
      const s = res.auftraege[rec.portal.id]; if (!s) continue;
      const before = JSON.stringify(rec.portal);
      rec.portal.lastSeen = s.lastSeen || rec.portal.lastSeen; rec.portal.locked = !!s.locked;
      let hist = '';
      if (s.status === 'eingereicht') {
        rec.portal.status = 'eingereicht'; rec.portal.submitted = s.submitted;
        if (['entwurf', 'mandant'].includes(rec.status)) rec.status = 'eingereicht';
        hist = `Mandant hat über das Portal eingereicht (${s.files} Datei(en))`;
      } else if (s.status === 'geloescht') { rec.portal.status = 'geloescht'; }
      if (JSON.stringify(rec.portal) === before) continue;
      changed = true;
      rec.updatedAt = new Date().toISOString();
      if (hist) { rec.history = rec.history || []; rec.history.push({ at: rec.updatedAt, by: 'Mandantenportal', t: hist }); }
      try { await S.store.save(rec); } catch (e) { if (!e.conflict) console.warn(e); }
    }
    return changed;
  }

  /* ---------- Abholen ---------- */
  async function fetchIn(rec, btn) {
    const p = rec.portal; if (!p) return;
    if (btn) { btn.disabled = true; btn.textContent = 'Hole ab …'; }
    try {
      const k = await loadKey(); if (!k) throw new Error('Privater Portal-Schlüssel fehlt in SharePoint.');
      const r = await call('abholen', undefined, `&id=${p.id}`);
      if (r.antwort.fp && r.antwort.fp !== k.fp) throw new Error('Die Daten wurden mit einem anderen Schlüssel verschlüsselt.');
      const dk = await X().openDataKey(k.priv, r.antwort.wk);
      const payload = JSON.parse(await X().decText(dk, r.antwort.ct));
      let n = 0;
      for (const f of r.dateien) {
        if (btn) btn.textContent = `Datei ${++n} von ${r.dateien.length} …`;
        const fk = await X().openDataKey(k.priv, f.fk);
        const meta = f.meta ? JSON.parse(await X().decText(fk, f.meta)) : { name: 'Datei', type: 'application/octet-stream' };
        const bytes = await X().decBytes(fk, await call('datei', undefined, `&id=${p.id}&fid=${f.fid}`, true));
        const file = new File([bytes], meta.name || 'Datei', { type: meta.type || 'application/octet-stream' });
        const up = await S.store.upload(rec, f.doc, file);
        const st = rec.docs[f.doc] = { ...(rec.docs[f.doc] || {}), files: [...((rec.docs[f.doc] || {}).files || []), { ...up, at: new Date().toISOString(), by: 'Mandantenportal' }] };
        st.state = 'datei'; delete rec.open['doc:' + f.doc];
      }
      rec.vorschlag = { values: payload.values || {}, groups: payload.groups || {}, at: payload.at };
      p.status = 'abgeholt'; p.fetchedAt = new Date().toISOString(); delete p.code;
      rec.updatedAt = p.fetchedAt; rec.updatedBy = S.user.name;
      rec.history.push({ at: p.fetchedAt, by: S.user.name, t: `Vom Portal abgeholt: ${Object.keys(rec.vorschlag.values).length + Object.keys(rec.vorschlag.groups).length} Angabe(n), ${r.dateien.length} Datei(en)` });
      await saveNow(rec);
      if (S.saveState === 'error') throw new Error('Speichern in SharePoint fehlgeschlagen – Daten bleiben auf dem Rückkanal, bitte erneut versuchen.');
      try { await call('loeschen', { id: p.id }); } catch (e) { toast('Hinweis: Löschen auf dem Rückkanal fehlgeschlagen – wird automatisch nachgeholt.', 'err'); }
      render();
      if (Object.keys(rec.vorschlag.values).length || Object.keys(rec.vorschlag.groups).length) dlgReview(rec);
      else { delete rec.vorschlag; touch(rec); toast('Unterlagen übernommen', 'ok'); }
    } catch (e) {
      console.error(e); toast('Abholen fehlgeschlagen: ' + e.message, 'err');
      if (btn) { btn.disabled = false; btn.textContent = 'Abholen und übernehmen'; }
    }
  }

  /* ---------- Vorschläge prüfen und übernehmen ---------- */
  function dlgReview(rec) {
    const v = rec.vorschlag; if (!v) return;
    const rows = Object.entries(v.values).map(([k, val]) => {
      const f = fieldBy(rec, k) || { l: k, t: 'text' }, old = rec.data[k];
      const same = JSON.stringify(old) === JSON.stringify(val), empty = LG.isEmpty(old);
      return `<tr><td><label class="chk"><input type="checkbox" data-take="${esc(k)}" ${empty || same ? 'checked' : ''}>${esc(f.l)}</label></td>
        <td class="sm ${empty ? 'muted' : ''}">${empty ? '–' : esc(fmtVal(f, old))}</td><td class="sm"><b>${esc(fmtVal(f, val))}</b>${!empty && !same ? ' <span class="badge b-warn">abweichend</span>' : ''}</td></tr>`;
    }).join('') + Object.entries(v.groups).map(([k, list]) => {
      const f = fieldBy(rec, k) || { l: k };
      return `<tr><td><label class="chk"><input type="checkbox" data-takeg="${esc(k)}" checked>${esc(f.l)}</label></td><td class="sm muted">${(rec.data[k] || []).length} vorhanden</td>
        <td class="sm"><b>${list.length} neu</b><br>${list.map(r => esc(Object.values(r).filter(Boolean).join(', '))).join('<br>')}</td></tr>`;
    }).join('');
    openDlg(`<form method="dialog"><div class="dlg-h">Angaben des Mandanten übernehmen</div><div class="dlg-b">
      <div class="hint">Leere Felder sind vorausgewählt. Bei abweichenden Werten entscheiden Sie selbst.</div>
      <div class="tblwrap"><table class="tbl"><thead><tr><th>Feld</th><th>Bisher</th><th>Vom Mandanten</th></tr></thead><tbody>${rows}</tbody></table></div></div>
      <div class="dlg-f"><button type="button" class="btn" data-close>Später</button><button type="button" class="btn primary" id="rv-ok">Ausgewählte übernehmen</button></div></form>`);
    $('#dlg').classList.add('wide'); $('#dlg').addEventListener('close', () => $('#dlg').classList.remove('wide'), { once: true });
    $('#rv-ok').addEventListener('click', () => {
      let n = 0;
      document.querySelectorAll('#dlg [data-take]').forEach(c => { if (!c.checked) return; const k = c.dataset.take; rec.data[k] = v.values[k]; delete rec.open[k]; n++; });
      document.querySelectorAll('#dlg [data-takeg]').forEach(c => { if (!c.checked) return; const k = c.dataset.takeg; rec.data[k] = [...(Array.isArray(rec.data[k]) ? rec.data[k] : []), ...v.groups[k]]; delete rec.open[k]; n++; });
      delete rec.vorschlag;
      touch(rec, `Angaben des Mandanten übernommen (${n})`);
      closeDlg(); render(); toast(`${n} Angabe(n) übernommen`, 'ok');
    });
  }

  async function act(name, rec, btn) {
    const p = rec.portal;
    switch (name) {
      case 'portal-send': return dlgSend(rec);
      case 'portal-code': return openDlg(`<form method="dialog"><div class="dlg-h">Zugangscode</div><div class="dlg-b"><div style="font:500 40px/1.1 var(--font);letter-spacing:.18em;text-align:center;padding:10px 0">${p.code ? fmtCode(p.code) : '–'}</div><div class="hint">Nur telefonisch oder persönlich an den Mandanten weitergeben.</div></div><div class="dlg-f"><button type="button" class="btn primary" data-close>Schließen</button></div></form>`);
      case 'portal-copy': return copy(link(p.id));
      case 'portal-revoke':
        if (!await confirmDlg('Link widerrufen', 'Der Mandant kann den Link danach nicht mehr verwenden. Bereits hochgeladene Dateien werden gelöscht.', 'Widerrufen')) return;
        try { await call('loeschen', { id: p.id }); } catch (e) { return toast(e.message, 'err'); }
        p.status = 'widerrufen'; delete p.code; touch(rec, 'Portal-Link widerrufen'); return render();
      case 'portal-fetch': return fetchIn(rec, btn);
      case 'portal-review': return dlgReview(rec);
    }
  }
  return { cardHtml, act, sync, request };
})();
