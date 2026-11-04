/* ProTax – Verschlüsselung für das Mandantenportal (Web Crypto, läuft im Browser)
 *
 * Kanzlei-Schlüsselpaar: RSA-OAEP 3072 / SHA-256
 *   öffentlicher Schlüssel  → config.js (fest im Portal hinterlegt)
 *   privater Schlüssel      → nur in SharePoint, nie auf dem Server
 * Jede Datei und jede Antwort bekommt einen eigenen zufälligen AES-GCM-256-Schlüssel,
 * der mit dem öffentlichen Kanzlei-Schlüssel eingepackt wird.
 * Die Anfrage (welche Punkte, Anrede) ist mit einem aus dem Zugangscode abgeleiteten
 * Schlüssel verschlüsselt (PBKDF2), damit auf dem Server keine Namen im Klartext liegen.
 */
(function (root) {
  'use strict';
  const subtle = (root.crypto || {}).subtle;
  const enc = new TextEncoder(), dec = new TextDecoder();
  const RSA = { name: 'RSA-OAEP', hash: 'SHA-256' };

  function b64(buf) {
    const b = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    let s = ''; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    return btoa(s);
  }
  function unb64(s) { const bin = atob(s); const b = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) b[i] = bin.charCodeAt(i); return b; }
  const rand = n => root.crypto.getRandomValues(new Uint8Array(n));
  const hex = b => Array.from(b, x => x.toString(16).padStart(2, '0')).join('');

  const C = {
    available: () => !!subtle,
    randomId: () => hex(rand(16)),
    randomFid: () => hex(rand(12)),
    randomCode() { // 6 Ziffern, gleichverteilt
      let n; do { n = new Uint32Array(rand(4).buffer)[0]; } while (n >= 4294000000);
      return String(n % 1000000).padStart(6, '0');
    },

    /* ---- Kanzlei-Schlüsselpaar ---- */
    async generateKeyPair() {
      const kp = await subtle.generateKey({ ...RSA, modulusLength: 3072, publicExponent: new Uint8Array([1, 0, 1]) }, true, ['encrypt', 'decrypt']);
      return { publicKey: await subtle.exportKey('jwk', kp.publicKey), privateKey: await subtle.exportKey('jwk', kp.privateKey) };
    },
    async fingerprint(pubJwk) {
      const canon = JSON.stringify({ e: pubJwk.e, kty: pubJwk.kty, n: pubJwk.n });
      const h = new Uint8Array(await subtle.digest('SHA-256', enc.encode(canon)));
      return hex(h.slice(0, 8)).match(/.{4}/g).join('-').toUpperCase();
    },
    importPublic: jwk => subtle.importKey('jwk', { kty: jwk.kty, n: jwk.n, e: jwk.e, alg: 'RSA-OAEP-256', ext: true }, RSA, false, ['encrypt']),
    importPrivate: jwk => subtle.importKey('jwk', { ...jwk, alg: 'RSA-OAEP-256', key_ops: ['decrypt'] }, RSA, false, ['decrypt']),

    /* ---- Daten für die Kanzlei verschlüsseln (Mandant) ---- */
    /** liefert { fk: eingepackter Schlüssel, key: AES-Schlüssel } */
    async newDataKey(pubKey) {
      const key = await subtle.generateKey({ name: 'AES-GCM', length: 256 }, true, ['encrypt', 'decrypt']);
      const raw = await subtle.exportKey('raw', key);
      const fk = b64(await subtle.encrypt(RSA, pubKey, raw));
      return { key, fk };
    },
    /** Bytes → Uint8Array (12 Byte IV + Geheimtext) */
    async encBytes(key, data) {
      const iv = rand(12);
      const ct = new Uint8Array(await subtle.encrypt({ name: 'AES-GCM', iv }, key, data));
      const out = new Uint8Array(12 + ct.length); out.set(iv); out.set(ct, 12); return out;
    },
    async encText(key, text) { return b64(await C.encBytes(key, enc.encode(text))); },

    /* ---- Entschlüsseln (Kanzlei) ---- */
    async openDataKey(privKey, fk) {
      const raw = await subtle.decrypt(RSA, privKey, unb64(fk));
      return subtle.importKey('raw', raw, 'AES-GCM', false, ['decrypt']);
    },
    async decBytes(key, data) {
      const b = data instanceof Uint8Array ? data : new Uint8Array(data);
      return new Uint8Array(await subtle.decrypt({ name: 'AES-GCM', iv: b.subarray(0, 12) }, key, b.subarray(12)));
    },
    async decText(key, s) { return dec.decode(await C.decBytes(key, unb64(s))); },

    /* ---- Anfrage mit Zugangscode schützen ---- */
    async codeKey(code, id) {
      const base = await subtle.importKey('raw', enc.encode(String(code)), 'PBKDF2', false, ['deriveKey']);
      return subtle.deriveKey({ name: 'PBKDF2', salt: enc.encode('protax-portal:' + id), iterations: 310000, hash: 'SHA-256' },
        base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
    },
    async sealWithCode(code, id, obj) { const k = await C.codeKey(code, id); return { v: 1, ct: await C.encText(k, JSON.stringify(obj)) }; },
    async openWithCode(code, id, box) { const k = await C.codeKey(code, id); return JSON.parse(await C.decText(k, box.ct)); },

    b64, unb64
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = C;
  else root.PTCRYPTO = C;
})(typeof window !== 'undefined' ? window : globalThis);
