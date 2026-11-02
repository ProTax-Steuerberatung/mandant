/* ProTax Mandant – Datenmodell der Mandantenaufnahme
 * Feldtypen: text, email, tel, date, select, check, multi, textarea, num, plz, stid, ustid, iban, group
 * types:  für welche Mandantentypen ein Feld/Abschnitt gilt (fehlt = alle)
 * req:    Pflichtfeld (true oder Liste von Mandantentypen)
 * when:   Funktion (daten) => sichtbar ja/nein; in Gruppen (zeile, daten)
 * w:      'full' = volle Breite
 * kanzlei: nur von der Kanzlei auszufüllen – wird nie beim Mandanten angefragt
 */
(function (root) {
  'use strict';
  const VERH = ['verheiratet', 'eingetragene Lebenspartnerschaft'];
  const ANREDE = ['Herr', 'Frau', 'ohne Anrede'];
  const KONF = ['keine', 'römisch-katholisch', 'evangelisch', 'sonstige'];

  const TYPES = {
    privat: { label: 'Privatperson / Investor', short: 'Privat' },
    freiberufler: { label: 'Freiberufler / Gründer', short: 'Freiberufler' },
    gmbh: { label: 'GmbH / Holding', short: 'GmbH' }
  };

  const STATUS = [
    { k: 'entwurf', l: 'Entwurf / Erstgespräch' },
    { k: 'mandant', l: 'Beim Mandanten' },
    { k: 'eingereicht', l: 'Eingereicht' },
    { k: 'geprueft', l: 'Geprüft' },
    { k: 'angelegt', l: 'In DATEV angelegt' }
  ];
  const STATUS_ABGEBROCHEN = { k: 'abgebrochen', l: 'Abgebrochen' };

  const SECTIONS = [
    {
      k: 'person', l: 'Person', types: ['privat', 'freiberufler'], fields: [
        { k: 'anrede', l: 'Anrede', t: 'select', o: ANREDE, req: true },
        { k: 'titel', l: 'Titel', t: 'text', ph: 'z. B. Dr.' },
        { k: 'vorname', l: 'Vorname', t: 'text', req: true },
        { k: 'nachname', l: 'Nachname', t: 'text', req: true },
        { k: 'geburtsdatum', l: 'Geburtsdatum', t: 'date', req: true, past: true },
        { k: 'geburtsort', l: 'Geburtsort', t: 'text', req: true },
        { k: 'staatsang', l: 'Staatsangehörigkeit', t: 'text', req: true, ph: 'deutsch' },
        { k: 'beruf', l: 'Beruf', t: 'text' },
        { k: 'strasse', l: 'Straße, Hausnummer', t: 'text', req: true },
        { k: 'plz', l: 'PLZ', t: 'plz', req: true },
        { k: 'ort', l: 'Ort', t: 'text', req: true },
        { k: 'email', l: 'E-Mail', t: 'email', req: true },
        { k: 'telefon', l: 'Telefon', t: 'tel' },
        { k: 'familienstand', l: 'Familienstand', t: 'select', o: ['ledig', 'verheiratet', 'eingetragene Lebenspartnerschaft', 'geschieden', 'dauernd getrennt lebend', 'verwitwet'], req: true },
        { k: 'familienstandSeit', l: 'Familienstand seit', t: 'date', when: d => d.familienstand && d.familienstand !== 'ledig', past: true },
        { k: 'konfession', l: 'Konfession (Kirchensteuer)', t: 'select', o: KONF, req: true }
      ]
    },
    {
      k: 'unternehmen', l: 'Unternehmen', types: ['gmbh'], fields: [
        { k: 'firma', l: 'Firma laut Handelsregister', t: 'text', req: true, w: 'full' },
        { k: 'rechtsform', l: 'Rechtsform', t: 'select', o: ['GmbH', 'UG (haftungsbeschränkt)', 'GmbH & Co. KG', 'AG', 'KG', 'OHG', 'GbR', 'sonstige'], req: true },
        { k: 'gruendung', l: 'Gründungsdatum', t: 'date', req: true, past: true },
        { k: 'strasse', l: 'Straße, Hausnummer', t: 'text', req: true },
        { k: 'plz', l: 'PLZ', t: 'plz', req: true },
        { k: 'ort', l: 'Ort', t: 'text', req: true },
        { k: 'registergericht', l: 'Registergericht', t: 'text', req: true, ph: 'Amtsgericht …', when: d => d.rechtsform !== 'GbR' },
        { k: 'hrnr', l: 'Handelsregister-Nummer', t: 'text', req: true, ph: 'HRB 12345', when: d => d.rechtsform !== 'GbR' },
        { k: 'gegenstand', l: 'Unternehmensgegenstand', t: 'textarea', w: 'full' },
        { k: 'branche', l: 'Branche', t: 'text' },
        { k: 'wj', l: 'Wirtschaftsjahr', t: 'select', o: ['Kalenderjahr', 'abweichend'], req: true },
        { k: 'wjBeginn', l: 'Beginn Wirtschaftsjahr', t: 'select', o: ['Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'], req: true, when: d => d.wj === 'abweichend' },
        { k: 'email', l: 'E-Mail Unternehmen', t: 'email', req: true },
        { k: 'telefon', l: 'Telefon', t: 'tel' },
        { k: 'holding', l: 'Holding – hält Beteiligungen an anderen Gesellschaften', t: 'check', w: 'full' },
        {
          k: 'beteiligungen', l: 'Beteiligungen', t: 'group', add: 'Beteiligung hinzufügen', req: true, when: d => !!d.holding, fields: [
            { k: 'firma', l: 'Gesellschaft', t: 'text', req: true },
            { k: 'anteil', l: 'Anteil %', t: 'num', req: true },
            { k: 'sitz', l: 'Sitz', t: 'text' }
          ]
        }
      ]
    },
    {
      k: 'familie', l: 'Partner & Kinder', types: ['privat', 'freiberufler'],
      hint: 'Keine Kinder? Abschnitt einfach leer lassen.', fields: [
        { k: 'p_vorname', l: 'Vorname Ehe-/Lebenspartner', t: 'text', req: true, when: d => VERH.includes(d.familienstand) },
        { k: 'p_nachname', l: 'Nachname', t: 'text', req: true, when: d => VERH.includes(d.familienstand) },
        { k: 'p_geburtsdatum', l: 'Geburtsdatum', t: 'date', req: true, past: true, when: d => VERH.includes(d.familienstand) },
        { k: 'p_stid', l: 'Steuer-ID Partner', t: 'stid', req: true, when: d => VERH.includes(d.familienstand) },
        { k: 'p_konfession', l: 'Konfession Partner', t: 'select', o: KONF, req: true, when: d => VERH.includes(d.familienstand) },
        { k: 'veranlagung', l: 'Veranlagung', t: 'select', o: ['Zusammenveranlagung', 'Einzelveranlagung', 'noch offen'], req: true, when: d => VERH.includes(d.familienstand) },
        {
          k: 'kinder', l: 'Kinder', t: 'group', add: 'Kind hinzufügen', fields: [
            { k: 'vorname', l: 'Vorname', t: 'text', req: true },
            { k: 'geburtsdatum', l: 'Geburtsdatum', t: 'date', req: true, past: true },
            { k: 'stid', l: 'Steuer-ID', t: 'stid' },
            { k: 'kindergeld', l: 'Kindergeld', t: 'select', o: ['ja', 'nein'] }
          ]
        }
      ]
    },
    {
      k: 'taetigkeit', l: 'Tätigkeit', types: ['freiberufler'], fields: [
        { k: 'taetigkeit', l: 'Tätigkeit / Berufsbezeichnung', t: 'text', req: true, w: 'full' },
        { k: 'art', l: 'Einordnung', t: 'select', o: ['freiberuflich (§ 18 EStG)', 'gewerblich (§ 15 EStG)', 'noch zu klären'], req: true },
        { k: 'beginn', l: 'Beginn der Tätigkeit', t: 'date', req: true },
        { k: 'betriebsstaette', l: 'Betriebsstätte (falls abweichend)', t: 'text', w: 'full' },
        { k: 'umsatzLfd', l: 'Umsatz lfd. Jahr (Schätzung, €)', t: 'num', req: true },
        { k: 'umsatzFolge', l: 'Umsatz Folgejahr (Schätzung, €)', t: 'num', req: true },
        { k: 'gewinnLfd', l: 'Gewinn lfd. Jahr (Schätzung, €)', t: 'num' },
        { k: 'kleinunternehmer', l: 'Kleinunternehmerregelung', t: 'select', o: ['nein', 'ja (§ 19 UStG)'], req: true },
        { k: 'versteuerung', l: 'Umsatzbesteuerung', t: 'select', o: ['Soll-Versteuerung', 'Ist-Versteuerung'], req: true, when: d => d.kleinunternehmer === 'nein' },
        { k: 'gewinnermittlung', l: 'Gewinnermittlung', t: 'select', o: ['Einnahmen-Überschuss-Rechnung', 'Bilanz'], req: true },
        { k: 'fragebogen', l: 'Fragebogen zur steuerlichen Erfassung', t: 'select', o: ['noch nicht abgegeben', 'übernimmt ProTax', 'bereits abgegeben'], req: true },
        { k: 'fragebogenDatum', l: 'abgegeben am', t: 'date', req: true, past: true, when: d => d.fragebogen === 'bereits abgegeben' }
      ]
    },
    {
      k: 'beteiligte', l: 'Geschäftsführer & Gesellschafter', types: ['gmbh'], fields: [
        {
          k: 'gf', l: 'Geschäftsführer / gesetzliche Vertreter', t: 'group', add: 'Geschäftsführer hinzufügen', req: true, fields: [
            { k: 'anrede', l: 'Anrede', t: 'select', o: ANREDE, req: true },
            { k: 'vorname', l: 'Vorname', t: 'text', req: true },
            { k: 'nachname', l: 'Nachname', t: 'text', req: true },
            { k: 'geburtsdatum', l: 'Geburtsdatum', t: 'date', req: true, past: true },
            { k: 'wohnort', l: 'Wohnort', t: 'text', req: true },
            { k: 'staatsang', l: 'Staatsangehörigkeit', t: 'text', req: true },
            { k: 'email', l: 'E-Mail', t: 'email' }
          ]
        },
        {
          k: 'gs', l: 'Gesellschafter', t: 'group', add: 'Gesellschafter hinzufügen', req: true, fields: [
            { k: 'name', l: 'Name / Firma', t: 'text', req: true },
            { k: 'anteil', l: 'Anteil %', t: 'num', req: true },
            { k: 'wb', l: 'Wirtschaftlich Berechtigter (> 25 %)', t: 'select', o: ['ja', 'nein'], req: true },
            { k: 'geburtsdatum', l: 'Geburtsdatum', t: 'date', req: true, past: true, when: r => r.wb === 'ja' },
            { k: 'wohnort', l: 'Wohnort', t: 'text', req: true, when: r => r.wb === 'ja' },
            { k: 'staatsang', l: 'Staatsangehörigkeit', t: 'text', req: true, when: r => r.wb === 'ja' }
          ]
        },
        { k: 'transparenz', kanzlei: true, l: 'Transparenzregister eingesehen am', t: 'date', req: true, past: true },
        { k: 'transparenzStimmig', kanzlei: true, l: 'Eintrag stimmt mit Angaben überein', t: 'select', o: ['ja', 'nein – Unstimmigkeit prüfen'], req: true }
      ]
    },
    {
      k: 'steuer', l: 'Steuer', fields: [
        { k: 'stid', l: 'Steuer-ID', t: 'stid', req: true, types: ['privat', 'freiberufler'] },
        { k: 'stnr', l: 'Steuernummer', t: 'text', req: ['privat', 'gmbh'], ph: '123/456/78901', hint: 'Bei Gründung ggf. noch nicht vorhanden' },
        { k: 'finanzamt', l: 'Finanzamt', t: 'text', req: true },
        { k: 'ustid', l: 'USt-IdNr.', t: 'ustid', types: ['freiberufler', 'gmbh'], ph: 'DE123456789' },
        { k: 'einkuenfte', l: 'Einkunftsarten', t: 'multi', types: ['privat', 'freiberufler'], req: true, w: 'full', o: ['Nichtselbständige Arbeit', 'Selbständige Arbeit', 'Gewerbebetrieb', 'Kapitalvermögen', 'Vermietung und Verpachtung', 'Renten / Pensionen', 'Sonstige'] },
        { k: 'vorberater', l: 'Bisheriger Steuerberater', t: 'text' },
        { k: 'vorberaterStatus', kanzlei: true, l: 'Unterlagen Vorberater', t: 'select', o: ['nicht zutreffend', 'angefordert', 'liegen vor'], when: d => !!d.vorberater }
      ]
    },
    {
      k: 'gwg', l: 'GwG-Identifizierung', fields: [
        { k: 'identPerson', kanzlei: true, l: 'Identifizierte Person', t: 'text', req: true, w: 'full', hint: 'Bei Gesellschaften: Geschäftsführer bzw. auftretende Person' },
        { k: 'identArt', kanzlei: true, l: 'Art der Identifizierung', t: 'select', o: ['persönlich mit Originaldokument', 'Kopie/Scan des Ausweises', 'Video-Ident', 'Online-Ausweis (eID)'], req: true },
        { k: 'ausweisArt', kanzlei: true, l: 'Dokument', t: 'select', o: ['Personalausweis', 'Reisepass', 'sonstiges amtliches Dokument'], req: true },
        { k: 'ausweisNr', kanzlei: true, l: 'Ausweisnummer', t: 'text', req: true },
        { k: 'behoerde', kanzlei: true, l: 'Ausstellende Behörde', t: 'text', req: true },
        { k: 'ausweisGueltig', kanzlei: true, l: 'Gültig bis', t: 'date', req: true, future: true },
        { k: 'pep', l: 'Politisch exponierte Person (PEP)', t: 'select', o: ['nein', 'ja'], req: true },
        { k: 'pepDetails', kanzlei: true, l: 'PEP – Funktion / Nähe', t: 'textarea', req: true, w: 'full', when: d => d.pep === 'ja' },
        { k: 'risiko', kanzlei: true, l: 'Risikoeinstufung', t: 'select', o: ['gering', 'normal', 'erhöht'], req: true },
        { k: 'identAm', kanzlei: true, l: 'Identifiziert am', t: 'date', req: true, past: true },
        { k: 'identVon', kanzlei: true, l: 'Identifiziert von', t: 'text', req: true }
      ]
    },
    {
      k: 'bank', l: 'Bank & Lastschrift', fields: [
        { k: 'iban', l: 'IBAN', t: 'iban', req: true, w: 'full' },
        { k: 'bic', l: 'BIC', t: 'text' },
        { k: 'kontoinhaber', l: 'Kontoinhaber', t: 'text', req: true },
        { k: 'sepa', kanzlei: true, l: 'SEPA-Lastschriftmandat', t: 'select', o: ['noch offen', 'erteilt'], req: true },
        { k: 'sepaDatum', kanzlei: true, l: 'Mandat erteilt am', t: 'date', req: true, past: true, when: d => d.sepa === 'erteilt' }
      ]
    },
    { k: 'unterlagen', l: 'Unterlagen', docs: true },
    {
      k: 'mandat', l: 'Auftrag & Mandat', fields: [
        { k: 'leistungen', kanzlei: true, l: 'Leistungsumfang', t: 'multi', req: true, w: 'full', o: ['Finanzbuchhaltung', 'Lohnabrechnung', 'Jahresabschluss', 'Einkommensteuererklärung', 'Betriebliche Steuererklärungen', 'Laufende Steuerberatung', 'Existenzgründungsberatung'] },
        { k: 'mandatAb', kanzlei: true, l: 'Mandat ab', t: 'date', req: true },
        { k: 'verguetung', kanzlei: true, l: 'Vergütung', t: 'select', o: ['nach StBVV', 'Pauschalvereinbarung', 'noch offen'], req: true },
        { k: 'auftrag', kanzlei: true, l: 'Auftrag', t: 'select', o: ['offen', 'verschickt', 'unterschrieben'], req: true },
        { k: 'vollmacht', kanzlei: true, l: 'Vollmacht Finanzverwaltung', t: 'select', o: ['offen', 'beantragt', 'erteilt'], req: true },
        { k: 'datenschutz', kanzlei: true, l: 'Datenschutzhinweise übergeben', t: 'check', req: true, w: 'full' },
        { k: 'notizen', l: 'Interne Notizen (nicht für den Mandanten)', t: 'textarea', w: 'full', internal: true }
      ]
    }
  ];

  const DOCS = [
    { k: 'ausweis', l: 'Ausweiskopie (Vorder- und Rückseite)', req: true },
    { k: 'bescheid', l: 'Letzter Steuerbescheid', types: ['privat', 'freiberufler'], req: true },
    { k: 'erklaerung', l: 'Letzte Steuererklärung', types: ['privat', 'freiberufler'] },
    { k: 'nachweis', l: 'Gewerbeanmeldung bzw. Nachweis der freiberuflichen Tätigkeit', types: ['freiberufler'], req: true },
    { k: 'fragebogenKopie', l: 'Fragebogen zur steuerlichen Erfassung (Kopie)', types: ['freiberufler'] },
    { k: 'hr', l: 'Handelsregisterauszug (aktuell)', types: ['gmbh'], req: true },
    { k: 'gv', l: 'Gesellschaftsvertrag / Satzung', types: ['gmbh'], req: true },
    { k: 'gsl', l: 'Gesellschafterliste', types: ['gmbh'], req: true },
    { k: 'tr', l: 'Transparenzregisterauszug', types: ['gmbh'], req: true },
    { k: 'jab', l: 'Letzter Jahresabschluss', types: ['gmbh'] },
    { k: 'bescheideFirma', l: 'Letzte Steuerbescheide (KSt, GewSt, USt)', types: ['gmbh'] },
    { k: 'auftragDoc', l: 'Unterschriebener Auftrag und Vollmacht', req: true },
    { k: 'sepaDoc', l: 'Unterschriebenes SEPA-Lastschriftmandat', req: true }
  ];

  const DOC_STATES = [
    { k: 'offen', l: 'offen' },
    { k: 'datei', l: 'Datei hochgeladen' },
    { k: 'papier', l: 'liegt in Papierform vor' },
    { k: 'entfaellt', l: 'entfällt' }
  ];

  const S = { TYPES, STATUS, STATUS_ABGEBROCHEN, SECTIONS, DOCS, DOC_STATES, VERH };
  if (typeof module !== 'undefined' && module.exports) module.exports = S;
  else root.SCHEMA = S;
})(typeof window !== 'undefined' ? window : globalThis);
