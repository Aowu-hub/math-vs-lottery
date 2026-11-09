// Shared helpers for the repo tools (Node 18+; no extra packages needed).
import { readFileSync, existsSync } from 'node:fs';
import { createInterface } from 'node:readline';
import { webcrypto as crypto } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

export const ROOT = join(fileURLToPath(new URL('..', import.meta.url)), '/');
export const LINK_FILE = ROOT + 'secrets/master-link.json';
export const LOCAL_FILE = ROOT + '.local/secrets.json';

export const norm = k => String(k || '').trim().toLowerCase().replace(/\s+/g, ' ');
const unb64 = s => Uint8Array.from(Buffer.from(s, 'base64'));

// Same format as master.html: PBKDF2-SHA256 (150k) → AES-GCM 256.
export async function openLink(key, L) {
  const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(norm(key)), 'PBKDF2', false, ['deriveKey']);
  const k = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt: unb64(L.salt), iterations: L.iter || 150000, hash: 'SHA-256' },
    base, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
  try {
    return JSON.parse(new TextDecoder().decode(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(L.iv) }, k, unb64(L.ct))));
  } catch (e) { throw new Error('🔑 That Master Key is not right.'); }
}

export async function lockLink(key, data) { // used by the tests
  const salt = crypto.getRandomValues(new Uint8Array(16)), iv = crypto.getRandomValues(new Uint8Array(12)), iter = 150000;
  const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(norm(key)), 'PBKDF2', false, ['deriveKey']);
  const k = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt']);
  const ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, k, new TextEncoder().encode(JSON.stringify(data)));
  const b = x => Buffer.from(x).toString('base64');
  return { v: 1, salt: b(salt), iv: b(iv), ct: b(ct), iter };
}

export function readLink() {
  if (!existsSync(LINK_FILE)) throw new Error('secrets/master-link.json is missing — set the Master Key on …/exec?page=master first, then copy "Master Link.json" from Drive into secrets/master-link.json.');
  return JSON.parse(readFileSync(LINK_FILE, 'utf8'));
}

// Key from the MVL_MASTER_KEY environment variable, or typed (hidden).
export function askKey(prompt = '🔑 Master Key: ') {
  if (process.env.MVL_MASTER_KEY) return Promise.resolve(process.env.MVL_MASTER_KEY);
  return new Promise(resolve => {
    const rl = createInterface({ input: process.stdin, output: process.stdout, terminal: true });
    rl._writeToOutput = s => { if (s.startsWith(prompt)) rl.output.write(prompt); };
    rl.question(prompt, a => { rl.close(); process.stdout.write('\n'); resolve(a); });
  });
}
