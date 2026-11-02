/* ProTax Mandant – Rechen- und Prüfkern (ohne Oberfläche, mit Node testbar) */
(function (root) {
  'use strict';
  const SCHEMA = (typeof module !== 'undefined' && module.exports) ? require('./schema.js') : root.SCHEMA;

  /* ---------- Validierung ---------- */
  function validSteuerId(v) {
    const s = String(v || '').replace(/\s/g, '');
    if (!/^\d{11}$/.test(s) || s[0] === '0') return false;
    const first = s.slice(0, 10), cnt = {};
    for (const c of first) cnt[c] = (cnt[c] || 0) + 1;
    const multi = Object.entries(cnt).filter(([, n]) => n > 1);
    if (multi.length !== 1 || multi[0][1] > 3) return false;
    if (multi[0][1] === 3 && first.includes(multi[0][0].repeat(3))) return false;
    let p = 10;
    for (let i = 0; i < 10; i++) {
      let sum = (+s[i] + p) % 10;
      if (sum === 0) sum = 10;
      p = (sum * 2) % 11;
    }
    let check = 11 - p;
    if (check === 10) check = 0;
    return check === +s[10];
  }
  function normIban(v) { return String(v || '').replace(/\s/g, '').toUpperCase(); }
  function validIban(v) {
    const s = normIban(v);
    if (!/^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/.test(s)) return false;
    if (s.startsWith('DE') && s.length !== 22) return false;
    const r = s.slice(4) + s.slice(0, 4);
    let mod = 0;
    for (const ch of r) {
      const n = /[A-Z]/.test(ch) ? String(ch.charCodeAt(0) - 55) : ch;
      for (const d of n) mod = (mod * 10 + +d) % 97;
    }
    return mod === 1;
  }
  function fmtIban(v) { return normIban(v).replace(/(.{4})/g, '$1 ').trim(); }
  const normUst = v => String(v || '').replace(/[\s.\-]/g, '').toUpperCase();
  const validEmail = v => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(v || '').trim());
  function num(v) {
    if (typeof v === 'number') return isNaN(v) ? null : v;
    let s = String(v ?? '').trim().replace(/[€\s%]/g, '');
    if (!s) return null;
    if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? null : n;
  }
  const todayIso = () => new Date().toISOString().slice(0, 10);
  const deDate = iso => { if (!iso) return ''; const [y, m, d] = String(iso).split('-'); return d && m && y ? `${d}.${m}.${y}` : String(iso); };

  /** Liefert eine Fehlermeldung oder '' */
  function checkValue(f, v, d) {
    if (isEmpty(v)) return '';
    switch (f.t) {
      case 'stid': return validSteuerId(v) ? '' : 'Steuer-ID ungültig (Prüfziffer)';
      case 'iban': return validIban(v) ? '' : 'IBAN ungültig';
      case 'email': return validEmail(v) ? '' : 'E-Mail-Adresse ungültig';
      case 'ustid': { const s = normUst(v); if (/^DE\d{9}$/.test(s)) return ''; return /^[A-Z]{2}[A-Z0-9]{2,12}$/.test(s) && !s.startsWith('DE') ? '' : 'USt-IdNr. ungültig (DE + 9 Ziffern)'; }
      case 'plz': return /^\d{5}$/.test(String(v).trim()) ? '' : '';
      case 'num': return num(v) == null ? 'Keine gültige Zahl' : '';
      case 'date': {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) return 'Datum ungültig';
        if (f.past && v > todayIso()) return 'Datum liegt in der Zukunft';
        if (f.future && v < todayIso()) return 'Dokument ist abgelaufen';
        return '';
      }
      default: return '';
    }
  }
  function isEmpty(v) {
    if (v == null) return true;
    if (Array.isArray(v)) return v.length === 0;
    if (typeof v === 'boolean') return !v;
    return String(v).trim() === '';
  }

  /* ---------- Schema-Zugriff ---------- */
  const appliesTo = (x, typ) => !x.types || x.types.includes(typ);
  const isReq = (f, typ) => f.req === true || (Array.isArray(f.req) && f.req.includes(typ));
  function sectionsFor(typ) { return SCHEMA.SECTIONS.filter(s => appliesTo(s, typ)); }
  function fieldsFor(sec, rec) {
    const d = rec.data || {};
    return (sec.fields || []).filter(f => appliesTo(f, rec.typ) && (!f.when || f.when(d)));
  }
  function rowFields(g, row, d) { return g.fields.filter(f => !f.when || f.when(row || {}, d)); }
  function docsFor(typ) { return SCHEMA.DOCS.filter(x => appliesTo(x, typ)); }
  function docState(rec, k) { return (rec.docs && rec.docs[k] && rec.docs[k].state) || 'offen'; }

  /* ---------- Bewertung ---------- */
  function evaluate(rec) {
    const d = rec.data || {}, open = rec.open || {};
    const out = { sections: {}, total: 0, done: 0, errors: [], warns: [], oks: [], missing: [], openList: [] };
    const addMissing = (sec, key, label, s) => {
      s.missing.push(label);
      out.missing.push({ sec: sec.k, key, label, section: sec.l, mandant: !!open[key] });
    };
    for (const sec of sectionsFor(rec.typ)) {
      const s = { total: 0, done: 0, missing: [], errors: [] };
      if (sec.docs) {
        for (const doc of docsFor(rec.typ)) {
          const st = docState(rec, doc.k), key = 'doc:' + doc.k;
          if (doc.req) { s.total++; if (st !== 'offen') s.done++; else addMissing(sec, key, doc.l, s); }
          if (open[key] && st === 'offen') out.openList.push({ key, label: doc.l, kind: 'doc' });
        }
      } else {
        for (const f of fieldsFor(sec, rec)) {
          if (f.t === 'group') {
            const rows = Array.isArray(d[f.k]) ? d[f.k] : [];
            if (isReq(f, rec.typ)) { s.total++; if (rows.length) s.done++; else addMissing(sec, f.k, f.l, s); }
            rows.forEach((row, i) => {
              for (const gf of rowFields(f, row, d)) {
                const v = row[gf.k], lab = `${f.l} ${i + 1}: ${gf.l}`;
                if (gf.req) { s.total++; if (!isEmpty(v)) s.done++; else addMissing(sec, `${f.k}.${i}.${gf.k}`, lab, s); }
                const e = checkValue(gf, v, d);
                if (e) { s.errors.push(`${lab}: ${e}`); }
              }
            });
            if (open[f.k] && !rows.length) out.openList.push({ key: f.k, label: f.l, kind: 'field' });
          } else {
            const v = d[f.k];
            if (isReq(f, rec.typ)) { s.total++; if (!isEmpty(v)) s.done++; else addMissing(sec, f.k, f.l, s); }
            const e = checkValue(f, v, d);
            if (e) s.errors.push(`${f.l}: ${e}`);
            if (open[f.k] && isEmpty(v)) out.openList.push({ key: f.k, label: f.l, kind: 'field' });
          }
        }
      }
      s.state = s.errors.length ? 'err' : s.missing.length ? 'warn' : 'ok';
      out.sections[sec.k] = s;
      out.total += s.total; out.done += s.done;
      s.errors.forEach(t => out.errors.push(`${sec.l} – ${t}`));
    }
    out.pct = out.total ? Math.round(out.done / out.total * 100) : 0;
    crossChecks(rec, out);
    return out;
  }

  function crossChecks(rec, out) {
    const d = rec.data || {}, W = t => out.warns.push(t), E = t => out.errors.push(t), O = t => out.oks.push(t);
    if (rec.typ === 'gmbh') {
      const gs = Array.isArray(d.gs) ? d.gs : [];
      if (gs.length) {
        const sum = gs.reduce((a, r) => a + (num(r.anteil) || 0), 0);
        if (Math.abs(sum - 100) > 0.01) W(`Gesellschafteranteile ergeben ${String(Math.round(sum * 100) / 100).replace('.', ',')} % statt 100 %`);
        else O('Gesellschafteranteile ergeben 100 %');
        if (!gs.some(r => r.wb === 'ja') && gs.every(r => r.wb)) W('Kein wirtschaftlich Berechtigter > 25 % – fiktiven wirtschaftlich Berechtigten (gesetzlicher Vertreter) dokumentieren');
        gs.forEach((r, i) => { const a = num(r.anteil); if (a != null && a > 25 && r.wb === 'nein') W(`Gesellschafter ${i + 1} hält über 25 % – ist er/sie wirtschaftlich berechtigt? Bei Gesellschaften als Gesellschafter: dahinterstehende Personen prüfen`); });
      }
      if (d.transparenzStimmig && d.transparenzStimmig !== 'ja') W('Transparenzregister weicht ab – Unstimmigkeitsmeldung nach § 23a GwG prüfen');
    }
    if (d.pep === 'ja') W('PEP – verstärkte Sorgfaltspflichten (§ 15 GwG): Herkunft der Vermögenswerte klären, Zustimmung der Kanzleileitung dokumentieren');
    if (d.risiko === 'erhöht') W('Erhöhtes Risiko – verstärkte Sorgfaltspflichten dokumentieren');
    if (d.identArt === 'persönlich mit Originaldokument' && docState(rec, 'ausweis') === 'offen') W('Persönlich identifiziert – Ausweiskopie trotzdem zur Akte nehmen');
    if (d.auftrag && d.auftrag !== 'unterschrieben') W('Auftrag noch nicht unterschrieben');
    if (d.vollmacht && d.vollmacht !== 'erteilt') W('Vollmacht Finanzverwaltung noch nicht erteilt');
    if (d.sepa === 'noch offen') W('SEPA-Lastschriftmandat noch offen');
    if (rec.typ !== 'gmbh' && d.stid && d.p_stid && String(d.stid).replace(/\s/g, '') === String(d.p_stid).replace(/\s/g, '')) E('Steuer-ID von Mandant und Partner sind identisch');
    if (rec.typ === 'freiberufler' && d.kleinunternehmer === 'ja (§ 19 UStG)') {
      const u1 = num(d.umsatzLfd), u2 = num(d.umsatzFolge);
      if ((u1 != null && u1 > 25000) || (u2 != null && u2 > 100000)) W('Umsatzschätzung über den Grenzen der Kleinunternehmerregelung – bitte prüfen');
    }
    const mand = out.missing.filter(m => m.mandant).length;
    if (mand) W(`${mand} Angabe(n) beim Mandanten angefragt`);
    if (!out.errors.length && !out.missing.length) O('Alle Pflichtangaben vollständig');
  }

  /* ---------- Hilfen für Oberfläche ---------- */
  function computeTitle(rec) {
    const d = rec.data || {};
    if (rec.typ === 'gmbh') return (d.firma || '').trim() || rec.title || 'Neue Aufnahme';
    const n = [d.nachname, d.vorname].filter(x => x && String(x).trim()).join(', ');
    return n || rec.title || 'Neue Aufnahme';
  }
  function contact(rec) {
    const d = rec.data || {};
    let anrede = d.anrede, titel = d.titel, nachname = d.nachname, email = d.email;
    if (rec.typ === 'gmbh' && Array.isArray(d.gf) && d.gf[0]) {
      const g = d.gf[0]; anrede = g.anrede; titel = ''; nachname = g.nachname; email = g.email || d.email;
    }
    let gruss = 'Sehr geehrte Damen und Herren,';
    const nm = [titel, nachname].filter(Boolean).join(' ');
    if (nachname && anrede === 'Herr') gruss = `Sehr geehrter Herr ${nm},`;
    else if (nachname && anrede === 'Frau') gruss = `Sehr geehrte Frau ${nm},`;
    return { email: email || '', gruss };
  }
  function mailText(rec, ev, user, kanzlei, hinweis) {
    const c = contact(rec);
    const fields = ev.openList.filter(x => x.kind === 'field').map(x => '– ' + x.label);
    const docs = ev.openList.filter(x => x.kind === 'doc').map(x => '– ' + x.label);
    const parts = [c.gruss, '',
      'vielen Dank für das angenehme Gespräch und Ihr Vertrauen in unsere Kanzlei. Damit wir Ihr Mandat vollständig anlegen können, benötigen wir von Ihnen noch Folgendes:', ''];
    if (fields.length) parts.push('Angaben:', ...fields, '');
    if (docs.length) parts.push('Unterlagen:', ...docs, '');
    if (!fields.length && !docs.length) parts.push('[Bitte offene Punkte ergänzen]', '');
    parts.push(hinweis || '', '', 'Bei Fragen bin ich gerne für Sie da.', '', 'Mit freundlichen Grüßen', '', user || '', kanzlei.name || '',
      [kanzlei.adresse, kanzlei.telefon, kanzlei.email].filter(Boolean).join(' · '));
    return { to: c.email, subject: `Ihre Mandatsaufnahme bei ${kanzlei.kurz || kanzlei.name || 'uns'} – noch benötigte Angaben`, body: parts.join('\n') };
  }

  const L = { validSteuerId, validIban, normIban, fmtIban, normUst, validEmail, num, deDate, todayIso, checkValue, isEmpty, appliesTo, isReq, sectionsFor, fieldsFor, rowFields, docsFor, docState, evaluate, computeTitle, contact, mailText };
  if (typeof module !== 'undefined' && module.exports) module.exports = L;
  else root.LOGIC = L;
})(typeof window !== 'undefined' ? window : globalThis);
