// Tests des Prüfkerns: node test/logic.test.js
const assert = require('assert');
const L = require('../logic.js');
let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok  ', name); };

t('Steuer-ID gültig', () => assert.ok(L.validSteuerId('86095742719')));
t('Steuer-ID Prüfziffer falsch', () => assert.ok(!L.validSteuerId('86095742718')));
t('Steuer-ID führende 0', () => assert.ok(!L.validSteuerId('06095742719')));
t('Steuer-ID Länge', () => assert.ok(!L.validSteuerId('8609574271')));
t('IBAN gültig', () => assert.ok(L.validIban('DE89 3704 0044 0532 0130 00')));
t('IBAN ungültig', () => assert.ok(!L.validIban('DE89370400440532013001')));
t('IBAN DE falsche Länge', () => assert.ok(!L.validIban('DE8937040044053201300')));
t('Zahl deutsch', () => assert.strictEqual(L.num('45.000,50'), 45000.5));
t('Zahl Prozent', () => assert.strictEqual(L.num('50 %'), 50));
t('Datum Zukunft bei past', () => assert.ok(L.checkValue({ t: 'date', past: true }, '2999-01-01')));
t('Ausweis abgelaufen', () => assert.ok(L.checkValue({ t: 'date', future: true }, '2001-01-01')));
t('USt-IdNr', () => { assert.strictEqual(L.checkValue({ t: 'ustid' }, 'DE 123 456 789'), ''); assert.ok(L.checkValue({ t: 'ustid' }, 'DE12345')); });

t('Leere GmbH: 0 %, Pflichtfelder fehlen', () => {
  const ev = L.evaluate({ typ: 'gmbh', data: {}, docs: {}, open: {} });
  assert.strictEqual(ev.pct, 0);
  assert.ok(ev.total > 20);
  assert.strictEqual(ev.sections.unternehmen.state, 'warn');
  assert.ok(!ev.sections.person);
});
t('Bedingte Felder: Partner nur bei verheiratet', () => {
  const a = L.evaluate({ typ: 'privat', data: { familienstand: 'ledig' } });
  const b = L.evaluate({ typ: 'privat', data: { familienstand: 'verheiratet' } });
  assert.ok(b.total > a.total);
});
t('Anteile ungleich 100 → Warnung', () => {
  const ev = L.evaluate({ typ: 'gmbh', data: { gs: [{ name: 'A', anteil: '60', wb: 'ja' }, { name: 'B', anteil: '30', wb: 'ja' }] } });
  assert.ok(ev.warns.some(w => w.includes('90')));
});
t('Ungültige IBAN → Fehler in Abschnitt Bank', () => {
  const ev = L.evaluate({ typ: 'privat', data: { iban: 'DE00123' } });
  assert.strictEqual(ev.sections.bank.state, 'err');
});
t('Offen für Mandanten: Feld und Unterlage', () => {
  const ev = L.evaluate({ typ: 'gmbh', data: { firma: 'X GmbH' }, open: { hrnr: true, firma: true, 'doc:hr': true } });
  const keys = ev.openList.map(x => x.key);
  assert.deepStrictEqual(keys.sort(), ['doc:hr', 'hrnr']);
});
t('Unterlage "entfällt" zählt als erledigt', () => {
  const a = L.evaluate({ typ: 'privat', data: {}, docs: {} });
  const b = L.evaluate({ typ: 'privat', data: {}, docs: { bescheid: { state: 'entfaellt' } } });
  assert.strictEqual(b.done, a.done + 1);
});
t('Titel und Anrede', () => {
  const rec = { typ: 'gmbh', data: { firma: 'Hofmann GmbH', gf: [{ anrede: 'Herr', nachname: 'Hofmann', email: 'h@x.de' }] } };
  assert.strictEqual(L.computeTitle(rec), 'Hofmann GmbH');
  assert.strictEqual(L.contact(rec).gruss, 'Sehr geehrter Herr Hofmann,');
  const p = { typ: 'privat', data: { anrede: 'Frau', titel: 'Dr.', vorname: 'Eva', nachname: 'Brandl' } };
  assert.strictEqual(L.computeTitle(p), 'Brandl, Eva');
  assert.strictEqual(L.contact(p).gruss, 'Sehr geehrte Frau Dr. Brandl,');
});
t('E-Mail enthält offene Punkte und Sie-Ansprache', () => {
  const rec = { typ: 'privat', data: { anrede: 'Herr', nachname: 'Maier', email: 'm@x.de' }, open: { stid: true, 'doc:bescheid': true } };
  const m = L.mailText(rec, L.evaluate(rec), 'Xenia Peter', { name: 'ProTax Steuerberatung' }, 'Hinweis');
  assert.ok(m.body.includes('– Steuer-ID') && m.body.includes('– Letzter Steuerbescheid') && m.body.includes('Ihr'));
  assert.strictEqual(m.to, 'm@x.de');
});
t('Vollständige Privat-Aufnahme = 100 %', () => {
  const docs = {}; L.docsFor('privat').forEach(x => docs[x.k] = { state: 'papier' });
  const data = { anrede: 'Herr', vorname: 'Max', nachname: 'Muster', geburtsdatum: '1980-01-01', geburtsort: 'Straubing', staatsang: 'deutsch', strasse: 'A 1', plz: '94315', ort: 'Straubing', email: 'max@muster.de', familienstand: 'ledig', konfession: 'keine',
    stid: '86095742719', stnr: '123/456/78901', finanzamt: 'Straubing', einkuenfte: ['Nichtselbständige Arbeit'],
    identPerson: 'Max Muster', identArt: 'Video-Ident', ausweisArt: 'Personalausweis', ausweisNr: 'L01X00T47', behoerde: 'Stadt Straubing', ausweisGueltig: '2035-01-01', pep: 'nein', risiko: 'gering', identAm: '2026-11-01', identVon: 'XP',
    iban: 'DE89370400440532013000', kontoinhaber: 'Max Muster', sepa: 'erteilt', sepaDatum: '2026-11-01',
    leistungen: ['Einkommensteuererklärung'], mandatAb: '2026-11-01', verguetung: 'nach StBVV', auftrag: 'unterschrieben', vollmacht: 'erteilt', datenschutz: true };
  const ev = L.evaluate({ typ: 'privat', data, docs });
  assert.deepStrictEqual(ev.errors, []);
  assert.strictEqual(ev.pct, 100, JSON.stringify(ev.missing));
});
console.log(`\n${n} Tests bestanden`);
t('Kanzlei-Felder landen nie beim Mandanten', () => {
  const rec = { typ: 'privat', data: {}, open: { identArt: true, risiko: true, stid: true, 'doc:ausweis': true } };
  const ev = L.evaluate(rec);
  assert.deepStrictEqual(ev.openList.map(x => x.key).sort(), ['doc:ausweis', 'stid']);
  assert.ok(ev.missing.find(m => m.key === 'risiko').kanzlei);
  const cl = L.checklist(rec);
  assert.deepStrictEqual(cl.map(s => s.k), ['steuer', 'unterlagen']);
});
t('Kurze E-Mail ohne Einzelpunkte', () => {
  const m = L.mailShort({ typ: 'privat', data: { anrede: 'Frau', nachname: 'Kern', email: 'k@x.de' } }, 'Xenia Peter', { name: 'ProTax Steuerberatung', kurz: 'ProTax' }, 'Hinweis');
  assert.ok(m.text.startsWith('Sehr geehrte Frau Kern,') && m.text.includes('beigefügten Checkliste'));
  assert.ok(m.toHtml(m.text).includes('<br>'));
});
console.log('Zusatztests ok');
