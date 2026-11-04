<?php
/*
 * ProTax Mandantenportal – Rückkanal (läuft auf Alfahosting, upload.protax.de)
 *
 * Speichert ausschließlich verschlüsselte Daten zwischen. Entschlüsseln kann nur die
 * Kanzlei-App (privater Schlüssel liegt in SharePoint, nie hier).
 *
 *  Kanzlei (Microsoft-Anmeldung, ID-Token wird geprüft):
 *    POST ?r=info | auftrag | sms | status | loeschen     GET ?r=abholen&id= | datei&id=&fid=
 *  Mandant (Link-ID + Zugangscode → Sitzung):
 *    POST ?r=zugang | upload&fid=&doc=&part=&final= | datei-weg | einreichen
 */
declare(strict_types=1);

const VERSION = '1.0';
$cfgFile = __DIR__ . '/config.php';
if (!is_file($cfgFile)) { http_response_code(500); header('Content-Type: application/json'); echo '{"error":"config.php fehlt"}'; exit; }
$C = require $cfgFile;
$C += [
  'allowed_origins' => [], 'data_dir' => __DIR__ . '/../portal-daten', 'jwks_url' => '',
  'max_file_mb' => 40, 'max_total_mb' => 250, 'max_files' => 60, 'max_tries' => 5,
  'ip_fail_limit' => 20, 'session_minutes' => 120, 'keep_days_after_expiry' => 30,
  'sms_api_key' => '', 'sms_sender' => 'ProTax', 'secret' => '',
];
if (!$C['jwks_url']) $C['jwks_url'] = 'https://login.microsoftonline.com/' . $C['tenant_id'] . '/discovery/v2.0/keys';

/* ---------------- Grundgerüst ---------------- */
header('X-Content-Type-Options: nosniff');
header('Cache-Control: no-store');
header('Referrer-Policy: no-referrer');
$origin = $_SERVER['HTTP_ORIGIN'] ?? '';
if ($origin && in_array($origin, $C['allowed_origins'], true)) {
  header('Access-Control-Allow-Origin: ' . $origin);
  header('Vary: Origin');
  header('Access-Control-Allow-Headers: Authorization, Content-Type, X-Session');
  header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
  header('Access-Control-Max-Age: 600');
}
if (($_SERVER['REQUEST_METHOD'] ?? '') === 'OPTIONS') { http_response_code(204); exit; }

function out($data, int $code = 200): void {
  http_response_code($code);
  header('Content-Type: application/json; charset=utf-8');
  echo json_encode($data, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
  exit;
}
function fail(string $msg, int $code = 400, array $extra = []): void { out(['error' => $msg] + $extra, $code); }
function body(): array {
  $raw = file_get_contents('php://input', false, null, 0, 2_000_000);
  $j = json_decode($raw ?: '[]', true);
  return is_array($j) ? $j : [];
}
function now(): int { return time(); }
function isId(?string $s): bool { return is_string($s) && preg_match('/^[a-f0-9]{32}$/', $s) === 1; }
function isFid(?string $s): bool { return is_string($s) && preg_match('/^[a-f0-9]{16,32}$/', $s) === 1; }

/* ---------------- Ablage ---------------- */
$D = rtrim($C['data_dir'], '/');
foreach (["$D", "$D/auftraege", "$D/limits"] as $dir) {
  if (!is_dir($dir) && !@mkdir($dir, 0700, true)) fail('Datenordner nicht beschreibbar', 500);
}
if (!is_file("$D/.htaccess")) @file_put_contents("$D/.htaccess", "Require all denied\nDeny from all\n");
if (!is_file("$D/index.html")) @file_put_contents("$D/index.html", '');

function apath(string $id): string { global $D; return "$D/auftraege/$id.json"; }
function fdir(string $id): string { global $D; return "$D/auftraege/$id"; }
function load(string $id): ?array {
  $p = apath($id); if (!is_file($p)) return null;
  $j = json_decode((string)file_get_contents($p), true); return is_array($j) ? $j : null;
}
/** Auftrag gesperrt lesen, ändern, zurückschreiben */
function withAuftrag(string $id, callable $fn) {
  $p = apath($id); if (!is_file($p)) fail('Link ungültig', 404, ['code' => 'unbekannt']);
  $h = fopen($p, 'c+'); if (!$h) fail('Speicherfehler', 500);
  flock($h, LOCK_EX);
  $a = json_decode(stream_get_contents($h) ?: '', true);
  if (!is_array($a)) { flock($h, LOCK_UN); fclose($h); fail('Speicherfehler', 500); }
  $res = $fn($a);
  ftruncate($h, 0); rewind($h); fwrite($h, json_encode($a, JSON_UNESCAPED_SLASHES)); fflush($h);
  flock($h, LOCK_UN); fclose($h);
  return $res;
}
function rrm(string $dir): void {
  if (!is_dir($dir)) return;
  foreach (scandir($dir) ?: [] as $f) { if ($f === '.' || $f === '..') continue; $p = "$dir/$f"; is_dir($p) ? rrm($p) : @unlink($p); }
  @rmdir($dir);
}
function deleteAuftrag(string $id): void { @unlink(apath($id)); rrm(fdir($id)); }

/* Abgelaufenes entfernen (höchstens alle 10 Minuten) */
(function () use ($D, $C) {
  $stamp = "$D/limits/.cleanup";
  if (is_file($stamp) && filemtime($stamp) > now() - 600) return;
  @touch($stamp);
  foreach (glob("$D/auftraege/*.json") ?: [] as $p) {
    $a = json_decode((string)@file_get_contents($p), true); if (!is_array($a)) continue;
    $exp = (int)($a['expires'] ?? 0);
    $limit = ($a['status'] ?? '') === 'eingereicht' ? $exp + 86400 * (int)$C['keep_days_after_expiry'] : $exp + 86400;
    if (now() > $limit) deleteAuftrag(basename($p, '.json'));
  }
  foreach (glob("$D/limits/*.json") ?: [] as $p) if (filemtime($p) < now() - 7200) @unlink($p);
})();

/* ---------------- Kanzlei-Anmeldung: Microsoft-ID-Token prüfen ---------------- */
function b64u(string $s): string { return (string)base64_decode(strtr($s, '-_', '+/') . str_repeat('=', (4 - strlen($s) % 4) % 4)); }
function fetchUrl(string $url): ?string {
  if (preg_match('#^https?://#', $url) !== 1) { $s = @file_get_contents($url); return $s === false ? null : $s; }
  if (function_exists('curl_init')) {
    $ch = curl_init($url);
    curl_setopt_array($ch, [CURLOPT_RETURNTRANSFER => true, CURLOPT_TIMEOUT => 10, CURLOPT_FOLLOWLOCATION => true]);
    $r = curl_exec($ch); $code = curl_getinfo($ch, CURLINFO_HTTP_CODE); curl_close($ch);
    return ($r !== false && $code === 200) ? (string)$r : null;
  }
  $r = @file_get_contents($url); return $r === false ? null : $r;
}
function jwks(bool $refresh = false): array {
  global $D, $C;
  $cache = "$D/limits/jwks.cache";
  if (!$refresh && is_file($cache) && filemtime($cache) > now() - 86400) {
    $j = json_decode((string)file_get_contents($cache), true); if (is_array($j)) return $j;
  }
  $raw = fetchUrl($C['jwks_url']); $j = $raw ? json_decode($raw, true) : null;
  if (!is_array($j) || empty($j['keys'])) fail('Anmeldeschlüssel von Microsoft nicht abrufbar', 502);
  @file_put_contents($cache, $raw);
  return $j;
}
function kanzlei(): array {
  global $C;
  $h = $_SERVER['HTTP_AUTHORIZATION'] ?? $_SERVER['REDIRECT_HTTP_AUTHORIZATION'] ?? '';
  if (!$h && function_exists('apache_request_headers')) { $hh = apache_request_headers(); $h = $hh['Authorization'] ?? $hh['authorization'] ?? ''; }
  if (!preg_match('/^Bearer\s+([A-Za-z0-9_\-]+)\.([A-Za-z0-9_\-]+)\.([A-Za-z0-9_\-]+)$/', $h, $m)) fail('Anmeldung fehlt', 401);
  $head = json_decode(b64u($m[1]), true); $claims = json_decode(b64u($m[2]), true); $sig = b64u($m[3]);
  if (!is_array($head) || !is_array($claims) || ($head['alg'] ?? '') !== 'RS256' || empty($head['kid'])) fail('Anmeldung ungültig', 401);
  $find = function (array $set) use ($head) { foreach ($set['keys'] as $k) if (($k['kid'] ?? '') === $head['kid']) return $k; return null; };
  $key = $find(jwks()); if (!$key) $key = $find(jwks(true));
  if (!$key) fail('Anmeldung ungültig (Schlüssel)', 401);
  if (!empty($key['x5c'][0])) {
    $pem = "-----BEGIN CERTIFICATE-----\n" . chunk_split($key['x5c'][0], 64, "\n") . "-----END CERTIFICATE-----\n";
  } else fail('Anmeldung ungültig (Zertifikat)', 401);
  $pub = openssl_pkey_get_public($pem);
  if (!$pub || openssl_verify($m[1] . '.' . $m[2], $sig, $pub, OPENSSL_ALGO_SHA256) !== 1) fail('Anmeldung ungültig (Signatur)', 401);
  $t = now();
  if (($claims['aud'] ?? '') !== $C['client_id']) fail('Anmeldung ungültig (App)', 401);
  if (($claims['tid'] ?? '') !== $C['tenant_id']) fail('Anmeldung ungültig (Organisation)', 401);
  if (($claims['iss'] ?? '') !== 'https://login.microsoftonline.com/' . $C['tenant_id'] . '/v2.0') fail('Anmeldung ungültig (Aussteller)', 401);
  if ((int)($claims['exp'] ?? 0) < $t - 60 || (int)($claims['nbf'] ?? 0) > $t + 60) fail('Anmeldung abgelaufen – bitte Seite neu laden', 401, ['code' => 'abgelaufen']);
  return ['name' => (string)($claims['name'] ?? ''), 'upn' => (string)($claims['preferred_username'] ?? '')];
}

/* ---------------- Mandanten-Sitzung ---------------- */
function ipKey(): string { global $D, $C; return "$D/limits/" . hash('sha256', ($_SERVER['REMOTE_ADDR'] ?? '') . $C['secret']) . '.json'; }
function ipFails(): int {
  $p = ipKey(); if (!is_file($p)) return 0;
  $j = json_decode((string)file_get_contents($p), true) ?: [];
  return count(array_filter($j, fn($t) => $t > now() - 3600));
}
function ipFail(): void {
  $p = ipKey(); $j = is_file($p) ? (json_decode((string)file_get_contents($p), true) ?: []) : [];
  $j = array_values(array_filter($j, fn($t) => $t > now() - 3600)); $j[] = now();
  file_put_contents($p, json_encode($j));
}
function mandant(string $id): array {
  global $C;
  $tok = $_SERVER['HTTP_X_SESSION'] ?? '';
  if (!isId($id) || !preg_match('/^[a-f0-9]{48}$/', $tok)) fail('Sitzung ungültig', 401, ['code' => 'sitzung']);
  $a = load($id); if (!$a) fail('Link ungültig', 404, ['code' => 'unbekannt']);
  $h = hash('sha256', $tok);
  if (empty($a['sessions'][$h]) || $a['sessions'][$h] < now()) fail('Sitzung abgelaufen – bitte Code erneut eingeben', 401, ['code' => 'sitzung']);
  if (($a['status'] ?? '') !== 'offen') fail('Bereits eingereicht', 409, ['code' => 'eingereicht']);
  if (now() > (int)$a['expires']) fail('Link abgelaufen', 410, ['code' => 'abgelaufen']);
  return $a;
}
function filesPublic(array $a): array {
  $o = [];
  foreach (($a['files'] ?? []) as $fid => $f) if (!empty($f['done'])) $o[] = ['fid' => $fid, 'doc' => $f['doc'], 'size' => $f['size'], 'meta' => $f['meta'] ?? null, 'fk' => $f['fk'] ?? null];
  return $o;
}

/* ---------------- Routen ---------------- */
$r = $_GET['r'] ?? '';
$method = $_SERVER['REQUEST_METHOD'] ?? 'GET';

switch ($r) {

  case 'info': {
    $u = kanzlei();
    out(['ok' => true, 'version' => VERSION, 'sms' => $C['sms_api_key'] !== '', 'user' => $u['name']]);
  }

  case 'auftrag': {
    if ($method !== 'POST') fail('Methode', 405);
    $u = kanzlei(); $b = body();
    $id = $b['id'] ?? ''; $code = (string)($b['code'] ?? ''); $days = (int)($b['tage'] ?? 14);
    if (!isId($id) || !preg_match('/^\d{6}$/', $code) || $days < 1 || $days > 60) fail('Ungültige Angaben');
    if (!is_array($b['anfrage'] ?? null) || strlen(json_encode($b['anfrage'])) > 200000) fail('Anfrage fehlt');
    if (is_file(apath($id))) fail('ID bereits vergeben', 409);
    $a = ['id' => $id, 'created' => now(), 'createdBy' => $u['upn'], 'expires' => now() + 86400 * $days,
      'codeHash' => password_hash($code, PASSWORD_DEFAULT), 'tries' => 0, 'locked' => false,
      'anfrage' => $b['anfrage'], 'status' => 'offen', 'files' => [], 'sessions' => [], 'total' => 0];
    file_put_contents(apath($id), json_encode($a, JSON_UNESCAPED_SLASHES), LOCK_EX);
    @chmod(apath($id), 0600);
    out(['ok' => true, 'id' => $id, 'expires' => $a['expires']]);
  }

  case 'sms': {
    if ($method !== 'POST') fail('Methode', 405);
    kanzlei(); $b = body();
    if ($C['sms_api_key'] === '') fail('SMS ist nicht eingerichtet', 400);
    $to = preg_replace('/[^\d+]/', '', (string)($b['an'] ?? '')); $text = mb_substr((string)($b['text'] ?? ''), 0, 300);
    if (strlen($to) < 8 || $text === '') fail('Handynummer oder Text fehlt');
    if (str_starts_with($to, '00')) $to = '+' . substr($to, 2);
    elseif (str_starts_with($to, '0')) $to = '+49' . substr($to, 1);
    $ch = curl_init('https://gateway.seven.io/api/sms');
    curl_setopt_array($ch, [CURLOPT_RETURNTRANSFER => true, CURLOPT_TIMEOUT => 15, CURLOPT_POST => true,
      CURLOPT_HTTPHEADER => ['X-Api-Key: ' . $C['sms_api_key'], 'Accept: application/json'],
      CURLOPT_POSTFIELDS => http_build_query(['to' => $to, 'text' => $text, 'from' => $C['sms_sender']])]);
    $res = curl_exec($ch); $code = curl_getinfo($ch, CURLINFO_HTTP_CODE); curl_close($ch);
    $j = is_string($res) ? json_decode($res, true) : null;
    $ok = $code === 200 && (is_array($j) ? (string)($j['success'] ?? '') === '100' : trim((string)$res) === '100');
    if (!$ok) fail('SMS konnte nicht gesendet werden', 502, ['detail' => is_string($res) ? mb_substr($res, 0, 200) : '']);
    out(['ok' => true]);
  }

  case 'status': {
    if ($method !== 'POST') fail('Methode', 405);
    kanzlei(); $b = body(); $o = [];
    foreach (array_slice((array)($b['ids'] ?? []), 0, 500) as $id) {
      if (!isId($id)) continue; $a = load($id);
      $o[$id] = $a ? ['status' => $a['status'], 'expires' => $a['expires'], 'submitted' => $a['submitted'] ?? null,
        'locked' => $a['locked'], 'files' => count(filesPublic($a)), 'lastSeen' => $a['lastSeen'] ?? null] : ['status' => 'geloescht'];
    }
    out(['ok' => true, 'auftraege' => $o]);
  }

  case 'abholen': {
    kanzlei(); $id = $_GET['id'] ?? '';
    if (!isId($id)) fail('ID ungültig'); $a = load($id); if (!$a) fail('Nicht gefunden', 404);
    if ($a['status'] !== 'eingereicht') fail('Noch nicht eingereicht', 409);
    out(['ok' => true, 'antwort' => $a['antwort'], 'dateien' => filesPublic($a), 'submitted' => $a['submitted']]);
  }

  case 'datei': {
    kanzlei(); $id = $_GET['id'] ?? ''; $fid = $_GET['fid'] ?? '';
    if (!isId($id) || !isFid($fid)) fail('ID ungültig');
    $p = fdir($id) . "/$fid.bin"; if (!is_file($p)) fail('Nicht gefunden', 404);
    header('Content-Type: application/octet-stream'); header('Content-Length: ' . filesize($p));
    readfile($p); exit;
  }

  case 'loeschen': {
    if ($method !== 'POST') fail('Methode', 405);
    kanzlei(); $b = body(); $id = $b['id'] ?? '';
    if (!isId($id)) fail('ID ungültig');
    deleteAuftrag($id); out(['ok' => true]);
  }

  /* ---- Mandant ---- */
  case 'zugang': {
    if ($method !== 'POST') fail('Methode', 405);
    if (ipFails() >= (int)$C['ip_fail_limit']) fail('Zu viele Versuche. Bitte in einer Stunde erneut versuchen.', 429, ['code' => 'gesperrt']);
    $b = body(); $id = $b['id'] ?? ''; $code = preg_replace('/\D/', '', (string)($b['code'] ?? ''));
    if (!isId($id)) { ipFail(); fail('Link ungültig', 404, ['code' => 'unbekannt']); }
    if (!is_file(apath($id))) { ipFail(); fail('Link ungültig oder bereits gelöscht', 404, ['code' => 'unbekannt']); }
    $res = withAuftrag($id, function (array &$a) use ($code, $C) {
      if ($a['status'] !== 'offen') return ['err' => ['Ihre Angaben wurden bereits eingereicht. Vielen Dank!', 409, 'eingereicht']];
      if (now() > (int)$a['expires']) return ['err' => ['Dieser Link ist abgelaufen. Bitte wenden Sie sich an die Kanzlei.', 410, 'abgelaufen']];
      if ($a['locked']) return ['err' => ['Der Zugang wurde nach mehreren falschen Codes gesperrt. Bitte wenden Sie sich an die Kanzlei.', 423, 'gesperrt']];
      if (!preg_match('/^\d{6}$/', $code) || !password_verify($code, $a['codeHash'])) {
        $a['tries']++; if ($a['tries'] >= (int)$C['max_tries']) $a['locked'] = true;
        $left = max(0, (int)$C['max_tries'] - $a['tries']);
        return ['fail' => true, 'err' => [$left ? "Der Code ist nicht korrekt. Noch $left Versuch(e)." : 'Der Code ist nicht korrekt. Der Zugang ist jetzt gesperrt – bitte wenden Sie sich an die Kanzlei.', $left ? 403 : 423, $left ? 'code' : 'gesperrt']];
      }
      $a['tries'] = 0;
      $tok = bin2hex(random_bytes(24));
      foreach ($a['sessions'] as $k => $exp) if ($exp < now()) unset($a['sessions'][$k]);
      $a['sessions'][hash('sha256', $tok)] = now() + 60 * (int)$C['session_minutes'];
      $a['lastSeen'] = now();
      return ['ok' => ['sitzung' => $tok, 'anfrage' => $a['anfrage'], 'expires' => $a['expires'], 'dateien' => filesPublic($a)]];
    });
    if (!empty($res['fail'])) ipFail();
    if (isset($res['err'])) fail($res['err'][0], $res['err'][1], ['code' => $res['err'][2]]);
    out(['ok' => true] + $res['ok']);
  }

  case 'upload': {
    if ($method !== 'POST') fail('Methode', 405);
    $id = $_GET['id'] ?? ''; $a = mandant($id);
    $fid = $_GET['fid'] ?? ''; $doc = $_GET['doc'] ?? ''; $part = (int)($_GET['part'] ?? -1); $final = ($_GET['final'] ?? '') === '1';
    if (!isFid($fid) || !preg_match('/^[A-Za-z]{1,40}$/', $doc) || $part < 0 || $part > 100) fail('Ungültige Angaben');
    $chunk = file_get_contents('php://input', false, null, 0, 9 * 1024 * 1024);
    if ($chunk === false || $chunk === '') fail('Leerer Upload');
    $maxFile = (int)$C['max_file_mb'] * 1048576; $maxTotal = (int)$C['max_total_mb'] * 1048576;
    $res = withAuftrag($id, function (array &$a) use ($id, $fid, $doc, $part, $final, $chunk, $maxFile, $maxTotal, $C) {
      $f = $a['files'][$fid] ?? null;
      if ($part === 0) {
        if ($f) return ['err' => ['Datei existiert bereits', 409]];
        if (count($a['files']) >= (int)$C['max_files']) return ['err' => ['Zu viele Dateien', 413]];
        $f = ['doc' => $doc, 'size' => 0, 'parts' => 0, 'done' => false, 'at' => now()];
        $meta = $_GET['meta'] ?? ''; $fk = $_GET['fk'] ?? '';
        if (strlen($meta) > 4000 || strlen($fk) > 2000 || !preg_match('/^[A-Za-z0-9+\/=._-]+$/', $meta . $fk)) return ['err' => ['Schlüsseldaten fehlen', 400]];
        $f['meta'] = $meta; $f['fk'] = $fk;
      } else {
        if (!$f || $f['done'] || $f['parts'] !== $part) return ['err' => ['Upload-Reihenfolge gestört', 409]];
      }
      $len = strlen($chunk);
      if ($f['size'] + $len > $maxFile) return ['err' => ['Datei zu groß (max. ' . $C['max_file_mb'] . ' MB)', 413]];
      if (($a['total'] ?? 0) + $len > $maxTotal) return ['err' => ['Gesamtgröße überschritten', 413]];
      $dir = fdir($id); if (!is_dir($dir)) mkdir($dir, 0700, true);
      if (file_put_contents("$dir/$fid.bin", $chunk, $part === 0 ? 0 : FILE_APPEND) !== $len) return ['err' => ['Speicherfehler', 500]];
      $f['size'] += $len; $f['parts']++; $a['total'] = ($a['total'] ?? 0) + $len;
      if ($final) $f['done'] = true;
      $a['files'][$fid] = $f; $a['lastSeen'] = now();
      return ['ok' => ['size' => $f['size'], 'done' => $f['done']]];
    });
    if (isset($res['err'])) fail($res['err'][0], $res['err'][1]);
    out(['ok' => true] + $res['ok']);
  }

  case 'datei-weg': {
    if ($method !== 'POST') fail('Methode', 405);
    $b = body(); $id = $b['id'] ?? ''; mandant($id); $fid = $b['fid'] ?? '';
    if (!isFid($fid)) fail('ID ungültig');
    withAuftrag($id, function (array &$a) use ($id, $fid) {
      if (!isset($a['files'][$fid])) return;
      $a['total'] = max(0, ($a['total'] ?? 0) - (int)$a['files'][$fid]['size']);
      unset($a['files'][$fid]); @unlink(fdir($id) . "/$fid.bin");
    });
    out(['ok' => true]);
  }

  case 'einreichen': {
    if ($method !== 'POST') fail('Methode', 405);
    $b = body(); $id = $b['id'] ?? ''; mandant($id);
    $ant = $b['antwort'] ?? null;
    if (!is_array($ant) || empty($ant['wk']) || empty($ant['ct']) || strlen(json_encode($ant)) > 1_500_000) fail('Antwort fehlt');
    withAuftrag($id, function (array &$a) use ($ant, $id) {
      foreach ($a['files'] as $fid => $f) if (empty($f['done'])) { unset($a['files'][$fid]); @unlink(fdir($id) . "/$fid.bin"); }
      $a['antwort'] = $ant; $a['status'] = 'eingereicht'; $a['submitted'] = now(); $a['sessions'] = [];
    });
    out(['ok' => true]);
  }

  default:
    fail('Unbekannte Anfrage', 404);
}
