/* ProTax Mandant – Konfiguration
 * Einzige Datei, die bei der Einrichtung angepasst wird (siehe ANLEITUNG.md).
 * Hier stehen keine Geheimnisse: Mandanten-ID und App-ID sind öffentlich unbedenklich.
 */
window.CONFIG = {
  // true = Testdaten nur in diesem Browser, keine Anmeldung, nichts geht an Microsoft
  demo: false,

  // aus Microsoft Entra ID (App-Registrierung „Mandant ProTax“)
  tenantId: 'c630873a-2003-4e72-975c-9ba6c08254f5',
  clientId: '8aa10b71-d684-4ca9-8cbe-12ceb772b976',
  redirectUri: '',               // leer = aktuelle Adresse der App

  sharepoint: {
    hostname: 'protaxsteuerberatung.sharepoint.com',
    sitePath: '/sites/Mandant'
  },

  kanzlei: {
    name: 'ProTax Steuerberatung',
    kurz: 'ProTax',
    adresse: 'Wittelsbacherstr. 1a, 94377 Steinach',
    telefon: '+49 151 65157779',
    email: 'info@protax.de',
    web: 'www.protax.de'
  },

  // Logo-Datei im App-Ordner, z. B. 'icons/logo-weiss.svg' (weiß, für die Kopfleiste)
  // und 'icons/logo.svg' (dunkel, für das Stammblatt). Leer = Schriftzug.
  logoHell: '',
  logoDunkel: '',

  maxUploadMB: 25,

  // Mandantenportal: Mandant ergänzt fehlende Angaben und lädt Unterlagen hoch
  portal: {
    api: 'https://upload.protax.de/api.php',          // Rückkanal auf Alfahosting
    url: 'https://mandant.protax.de/portal/',          // Seite für den Mandanten
    publicKey: null,                                   // öffentlicher Kanzlei-Schlüssel (wird bei der Einrichtung eingetragen)
    gueltigTage: 14,                                   // Vorgabe für die Gültigkeit des Links
    datenschutz: 'https://www.protax.de/datenschutz',
    impressum: 'https://www.protax.de/impressum'
  },

  // Satz in der E-Mail an den Mandanten, wie Unterlagen zu uns kommen
  mailHinweis: 'Gerne können Sie uns die Unterlagen als Antwort auf diese E-Mail zusenden oder zu unserem nächsten Termin mitbringen.'
};
