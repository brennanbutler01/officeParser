/**
 * Decrypts password-protected OpenDocument files (.odt/.ods/.odp/.odg) per the ODF 1.2 packaging
 * spec (§3.4 "Encryption").
 *
 * Unlike OOXML, an encrypted ODF is still an ordinary zip: the entry *names* are visible, but each
 * listed entry's bytes are deflate-compressed and then AES-encrypted, with the per-entry parameters
 * recorded in `META-INF/manifest.xml`. We read the manifest, decrypt and inflate every encrypted
 * entry, then rebuild a plain zip the normal ODF parser can read. Modern AES-256/192/128-CBC is
 * supported; the legacy Blowfish scheme (OpenOffice.org < 3.4) is not, and raises a clear error.
 *
 * Because the input is untrusted, both the zip inflation (via the parser's streaming `extractFiles`
 * guard) and the second, per-entry deflate stream (`inflateRawCapped`) are bounded by
 * `decompressionLimits`, and the attacker-controlled PBKDF2 `iteration-count` is validated first.
 *
 * @module crypto/odfCrypto
 */

import { createDecipheriv, createHash, pbkdf2Sync } from 'crypto';
import { Inflate, zipSync, Zippable } from 'fflate';
import { DecompressionLimits, OfficeParserConfig } from '../types.js';
import { extractFiles } from '../utils/zipUtils.js';
import { WRONG_PASSWORD, DecryptionError } from './wrongPassword.js';

/** Real ODF files use 1024–100000 PBKDF2 iterations; reject anything that would be a CPU DoS. */
const MAX_ITERATIONS = 10_000_000;
const DEFAULT_MAX_BYTES = 512 * 1024 * 1024;

interface EncryptionData {
    fullPath: string;
    checksumType: string;
    checksum: Buffer;
    algoName: string;
    iv: Buffer;
    keyDerivName: string;
    keySize: number;
    iterationCount: number;
    salt: Buffer;
    startKeyName: string;
}

/** True when the buffer even looks like a zip (PK\x03\x04); avoids streaming a PDF/RTF through Unzip. */
function isZip(buf: Uint8Array): boolean {
    return buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b && buf[2] === 0x03 && buf[3] === 0x04;
}

function maxBytesOf(limits?: DecompressionLimits): number {
    const v = limits?.maxUncompressedBytes;
    return v !== undefined && Number.isFinite(v) && v >= 0 ? v : DEFAULT_MAX_BYTES;
}

/** Streaming, limit-enforcing unzip (reuses the parser's zip-bomb guard); returns a name -> bytes map. */
async function unzipCapped(buf: Uint8Array, filter: (name: string) => boolean, limits?: DecompressionLimits, config?: OfficeParserConfig): Promise<Record<string, Uint8Array>> {
    const src = Buffer.isBuffer(buf) ? buf : Buffer.from(buf.buffer, buf.byteOffset, buf.byteLength);
    const files = await extractFiles(src, filter, limits ?? {}, config);
    const map: Record<string, Uint8Array> = {};
    for (const f of files) map[f.path] = f.content;
    return map;
}

async function readManifest(buf: Uint8Array, limits?: DecompressionLimits, config?: OfficeParserConfig): Promise<string | null> {
    if (!isZip(buf)) return null;
    const map = await unzipCapped(buf, name => name === 'META-INF/manifest.xml', limits, config);
    const bytes = map['META-INF/manifest.xml'];
    return bytes ? Buffer.from(bytes).toString('utf8') : null;
}

/**
 * True when the buffer is a zip whose `META-INF/manifest.xml` marks at least one entry encrypted.
 * A plain ODF (or any other zip, e.g. a docx) has no `<manifest:encryption-data>` and returns false.
 */
export async function isEncryptedOdf(buf: Uint8Array, limits?: DecompressionLimits, config?: OfficeParserConfig): Promise<boolean> {
    try {
        const manifest = await readManifest(buf, limits, config);
        return !!manifest && manifest.includes('encryption-data');
    } catch {
        return false;
    }
}

function attr(scope: string, name: string): string {
    const m = new RegExp(`\\b${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}="([^"]*)"`).exec(scope);
    return m ? m[1] : '';
}

/** Parses every `<manifest:file-entry>` that carries encryption metadata. */
function parseEncryptionData(manifest: string): Map<string, EncryptionData> {
    const map = new Map<string, EncryptionData>();
    // Match a self-closing file-entry OR a child-bearing one. A naive `[\s\S]*?(?:\/>|</...>)` would
    // stop at the first inner self-closing child (e.g. `<manifest:algorithm/>`), truncating the entry
    // before `key-derivation`.
    const entryRe = /<manifest:file-entry\b[^>]*\/>|<manifest:file-entry\b[^>]*>[\s\S]*?<\/manifest:file-entry>/g;
    let m: RegExpExecArray | null;
    while ((m = entryRe.exec(manifest))) {
        const block = m[0];
        if (!block.includes('encryption-data')) continue;
        const algo = /<manifest:algorithm\b[^>]*>/.exec(block)?.[0] || '';
        const kd = /<manifest:key-derivation\b[^>]*>/.exec(block)?.[0] || '';
        const skg = /<manifest:start-key-generation\b[^>]*>/.exec(block)?.[0] || '';
        const enc = /<manifest:encryption-data\b[^>]*>/.exec(block)?.[0] || '';
        map.set(attr(block, 'manifest:full-path'), {
            fullPath: attr(block, 'manifest:full-path'),
            checksumType: attr(enc, 'manifest:checksum-type'),
            checksum: Buffer.from(attr(enc, 'manifest:checksum'), 'base64'),
            algoName: attr(algo, 'manifest:algorithm-name'),
            iv: Buffer.from(attr(algo, 'manifest:initialisation-vector'), 'base64'),
            keyDerivName: attr(kd, 'manifest:key-derivation-name'),
            keySize: parseInt(attr(kd, 'manifest:key-size'), 10) || 16,
            iterationCount: parseInt(attr(kd, 'manifest:iteration-count'), 10) || 1024,
            salt: Buffer.from(attr(kd, 'manifest:salt'), 'base64'),
            startKeyName: skg ? attr(skg, 'manifest:start-key-generation-name') : '',
        });
    }
    return map;
}

/** AES key size in bits from the algorithm URI, defaulting to 256 (the ODF 1.2 default). */
function aesBitsFromAlgo(algoName: string): number {
    if (/aes256/i.test(algoName)) return 256;
    if (/aes192/i.test(algoName)) return 192;
    if (/aes128/i.test(algoName)) return 128;
    if (/blowfish/i.test(algoName)) return -1; // signals unsupported
    return 256;
}

function checksumHash(type: string, data: Uint8Array): Buffer {
    const algo = /sha256/i.test(type) ? 'sha256' : 'sha1';
    // ODF checksums cover the first 1024 bytes of the compressed (pre-encryption) data.
    return createHash(algo).update(data.subarray(0, 1024)).digest();
}

/**
 * Raw-inflates `data`, aborting once the output would exceed `maxBytes`. The input is pushed in
 * bounded chunks so a deflate bomb inflates at most ~one chunk past the cap before we stop, rather
 * than allocating gigabytes. Malformed input (a wrong AES key produces garbage) makes fflate throw,
 * which the caller treats as a wrong password.
 */
function inflateRawCapped(data: Uint8Array, maxBytes: number): Uint8Array {
    const parts: Uint8Array[] = [];
    let total = 0;
    let over = false;
    const inf = new Inflate(chunk => {
        if (over) return;
        total += chunk.length;
        if (total > maxBytes) { over = true; return; }
        parts.push(chunk);
    });
    const CH = 1 << 16;
    for (let o = 0; o < data.length && !over; o += CH) {
        inf.push(data.subarray(o, Math.min(o + CH, data.length)), o + CH >= data.length);
    }
    if (over) throw new DecryptionError('an encrypted ODF entry decompresses beyond the allowed size');
    const out = new Uint8Array(total);
    let off = 0;
    for (const p of parts) { out.set(p, off); off += p.length; }
    return out;
}

/**
 * Decrypts an encrypted ODF buffer with the given password, returning a plain (re-zipped) ODF.
 * Throws {@link WRONG_PASSWORD} on a bad password, or a {@link DecryptionError} for an unsupported
 * cipher, an implausible key-derivation parameter, a decompression-limit breach, or a broken package.
 */
export async function decryptOdf(buf: Uint8Array, password: string, limits?: DecompressionLimits, config?: OfficeParserConfig): Promise<Uint8Array> {
    const manifest = await readManifest(buf, limits, config);
    if (!manifest) throw new DecryptionError('encrypted ODF: missing META-INF/manifest.xml');
    const encMap = parseEncryptionData(manifest);
    const all = await unzipCapped(buf, () => true, limits, config); // encrypted entries are STORED -> raw bytes
    const maxBytes = maxBytesOf(limits);

    const out: Zippable = {};
    const pwBytes = Buffer.from(password, 'utf8');

    for (const [name, bytes] of Object.entries(all)) {
        const enc = encMap.get(name);
        if (!enc) { out[name] = bytes; continue; } // mimetype, manifest.xml, directories: copy as-is

        const bits = aesBitsFromAlgo(enc.algoName);
        if (bits < 0) throw new DecryptionError(`encrypted ODF: '${enc.algoName}' (Blowfish) is not supported; only AES-CBC is`);
        if (!/cbc/i.test(enc.algoName)) throw new DecryptionError(`encrypted ODF: unsupported cipher '${enc.algoName}' (only AES-CBC)`);
        // iteration-count is attacker-controlled; reject an absurd value rather than grinding PBKDF2.
        if (!Number.isFinite(enc.iterationCount) || enc.iterationCount < 1 || enc.iterationCount > MAX_ITERATIONS) {
            throw new DecryptionError(`encrypted ODF: implausible iteration-count ${enc.iterationCount}`);
        }

        // Start key: hash of the password; then PBKDF2 (HMAC-SHA1) stretches it to the AES key. The
        // key length comes from the cipher (bits/8), which is inherently one of 16/24/32.
        const startAlgo = /sha256/i.test(enc.startKeyName) ? 'sha256' : 'sha1';
        const startKey = createHash(startAlgo).update(pwBytes).digest();
        const key = pbkdf2Sync(startKey, enc.salt, enc.iterationCount, bits / 8, 'sha1');

        // The encrypted length is a whole number of AES blocks; trim any trailing partial block.
        const aligned = bytes.subarray(0, bytes.length - (bytes.length % 16));
        const decipher = createDecipheriv(`aes-${bits}-cbc`, key, enc.iv);
        decipher.setAutoPadding(false);
        const compressed = Buffer.concat([decipher.update(aligned), decipher.final()]);

        // The ODF checksum (SHA of the first 1 KiB of the compressed data) is authoritative, but only
        // when we can reproduce those 1 KiB exactly: the tail is zero-padded to the block size, so a
        // short entry's checksum would cover padding the encoder never hashed. Use it above ~1 KiB and,
        // always, fall back to raw inflate, which effectively never succeeds on a wrong AES key.
        if (enc.checksum.length && compressed.length > 1024 + 16) {
            const actual = checksumHash(enc.checksumType, compressed);
            if (!actual.subarray(0, enc.checksum.length).equals(enc.checksum)) throw WRONG_PASSWORD;
        }

        let plain: Uint8Array;
        try { plain = inflateRawCapped(compressed, maxBytes); }
        catch (e) {
            if (e instanceof DecryptionError) throw e; // a real limit breach, not a wrong password
            throw WRONG_PASSWORD;                      // an undamaged package only fails to inflate on a bad key
        }
        out[name] = plain;
    }

    // mimetype must be first and stored for the package to be recognized as ODF.
    if (out['mimetype']) {
        const mimetype = out['mimetype'];
        delete out['mimetype'];
        return zipSync({ mimetype: [mimetype as Uint8Array, { level: 0 }], ...out });
    }
    return zipSync(out);
}
