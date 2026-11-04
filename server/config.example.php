<?php
/*
 * Konfiguration Rückkanal Mandantenportal
 * Diese Datei als config.php neben api.php ablegen und ausfüllen.
 * config.php gehört NICHT ins GitHub-Repository.
 */
return [
  // aus Entra (App-Registrierung „ProTax Mandant“) – dieselben Werte wie in config.js
  'tenant_id' => 'c630873a-2003-4e72-975c-9ba6c08254f5',
  'client_id' => '8aa10b71-d684-4ca9-8cbe-12ceb772b976',

  // Seiten, die den Rückkanal aufrufen dürfen
  'allowed_origins' => ['https://mandant.protax.de'],

  // Ablageordner – möglichst AUSSERHALB des Web-Ordners der Subdomain
  'data_dir' => __DIR__ . '/../portal-daten',

  // Zufallswert (beliebig lang) – einmal erzeugen, nicht ändern
  'secret' => 'HIER-EINEN-LANGEN-ZUFALLSWERT-EINTRAGEN',

  // SMS über seven.io (optional). Leer = SMS-Versand ausgeschaltet.
  'sms_api_key' => '',
  'sms_sender'  => 'ProTax',

  // Grenzen
  'max_file_mb'  => 40,   // je Datei
  'max_total_mb' => 250,  // je Mandant
  'max_files'    => 60,
  'max_tries'    => 5,    // falsche Codes bis zur Sperre
  'keep_days_after_expiry' => 30, // eingereichte, nicht abgeholte Daten werden danach gelöscht
];
