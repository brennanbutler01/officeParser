/**
 * Decrypts password-protected OOXML documents (.docx/.xlsx/.pptx) per [MS-OFFCRYPTO].
 *
 * Two schemes are handled: **agile encryption** (the modern default, an XML descriptor plus
 * AES-CBC) and **standard encryption** (older, a binary header plus AES-ECB). Both wrap the real
 * document zip in a CFB container's `EncryptedPackage` stream, described by `EncryptionInfo`.
 * Extensible encryption and RC4/CryptoAPI schemes are rare and not supported; they raise a clear
 * "unsupported" error rather than a wrong-password one.
 *
 * @module crypto/ooxmlCrypto
 */

import { createDecipheriv, createHash } from 'crypto';
import { CfbContainer, isCfb } from './cfb.js';
import { WRONG_PASSWORD, DecryptionError } from './wrongPassword.js';

/** Block keys from [MS-OFFCRYPTO] §2.3.4.10, one per derived sub-key in the agile scheme. */
const BLOCK_VERIFIER_INPUT = Uint8Array.from([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const BLOCK_VERIFIER_VALUE = Uint8Array.from([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const BLOCK_KEY_VALUE = Uint8Array.from([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);

/** True when the buffer is a CFB container that carries the OOXML encryption streams. */
export function isEncryptedOoxml(buf: Uint8Array): boolean {
    if (!isCfb(buf)) return false;
    try {
        const cfb = new CfbContainer(buf);
        return cfb.has('EncryptionInfo') && cfb.has('EncryptedPackage');
    } catch {
        return false;
    }
}

/**
 * Decrypts an encrypted OOXML buffer with the given password, returning the plaintext document zip.
 * Throws {@link WRONG_PASSWORD} when the password does not verify, or a descriptive Error for a
 * malformed or unsupported container.
 */
export function decryptOoxml(buf: Uint8Array, password: string): Uint8Array {
    const cfb = new CfbContainer(buf);
    const info = cfb.read('EncryptionInfo');
    const pkg = cfb.read('EncryptedPackage');
    if (!info || !pkg) throw new DecryptionError('encrypted OOXML: missing EncryptionInfo/EncryptedPackage');

    const view = new DataView(info.buffer, info.byteOffset, info.byteLength);
    const versionMajor = view.getUint16(0, true);
    const versionMinor = view.getUint16(2, true);

    if (versionMajor === 4 && versionMinor === 4) return decryptAgile(info, pkg, password);
    if (versionMinor === 2 && (versionMajor === 2 || versionMajor === 3 || versionMajor === 4)) return decryptStandard(info, pkg, password);
    throw new DecryptionError(`encrypted OOXML: unsupported encryption version ${versionMajor}.${versionMinor} (only agile and standard AES are supported)`);
}

// ── agile ──────────────────────────────────────────────────────────────────────

interface AgileKeyBlock {
    saltValue: Buffer;
    blockSize: number;
    keyBits: number;
    hashAlgorithm: string;
    cipherAlgorithm: string;
}

function attr(xml: string, tag: string, name: string): string {
    // Attributes never contain '>' unescaped, so a per-element slice keeps `keyData` and
    // `encryptedKey` attributes of the same name apart.
    const el = new RegExp(`<[^>]*\\b${tag}\\b[^>]*>`).exec(xml);
    const scope = el ? el[0] : xml;
    const m = new RegExp(`\\b${name}="([^"]*)"`).exec(scope);
    return m ? m[1] : '';
}

/** Maps an [MS-OFFCRYPTO] hash name to a Node hash id, rejecting anything unrecognized. */
function hashId(name: string): string {
    const n = name.toUpperCase().replace(/-/g, '');
    if (n === 'SHA512') return 'sha512';
    if (n === 'SHA384') return 'sha384';
    if (n === 'SHA256') return 'sha256';
    if (n === 'SHA1') return 'sha1';
    if (n === 'MD5') return 'md5';
    throw new DecryptionError(`encrypted OOXML: unsupported hash algorithm '${name}'`);
}

function hash(algo: string, ...parts: Uint8Array[]): Buffer {
    const h = createHash(algo);
    for (const p of parts) h.update(p);
    return h.digest();
}

function le32(n: number): Buffer {
    const b = Buffer.alloc(4);
    b.writeUInt32LE(n >>> 0, 0);
    return b;
}

function utf16le(s: string): Buffer {
    return Buffer.from(s, 'utf16le');
}

function aesDecryptNoPad(data: Uint8Array, key: Buffer, iv: Buffer | null, keyBits: number, mode: 'cbc' | 'ecb'): Buffer {
    // ECB takes no IV. Node tolerates `null`, but the browser bundle's crypto polyfill dereferences it
    // and throws; an empty buffer is accepted by both. CBC always receives a real IV from the caller.
    const decipher = createDecipheriv(`aes-${keyBits}-${mode}`, key, mode === 'ecb' ? Buffer.alloc(0) : iv);
    decipher.setAutoPadding(false);
    return Buffer.concat([decipher.update(data), decipher.final()]);
}

/** Derives the sub-key for a given block key: `hash(passwordHash || blockKey)`, sized to `keyBits`. */
function deriveBlockKey(passwordHash: Buffer, blockKey: Uint8Array, keyBits: number, algo: string): Buffer {
    let k = hash(algo, passwordHash, blockKey);
    const bytes = keyBits / 8;
    if (k.length < bytes) k = Buffer.concat([k, Buffer.alloc(bytes - k.length, 0x36)]);
    return k.subarray(0, bytes);
}

/** Iterated password hash: `H(salt||pw)` then `spinCount` rounds of `H(LE32(i)||H)`. */
function agilePasswordHash(salt: Buffer, password: string, spinCount: number, algo: string): Buffer {
    let h = hash(algo, salt, utf16le(password));
    for (let i = 0; i < spinCount; i++) h = hash(algo, le32(i), h);
    return h;
}

function decryptAgile(info: Uint8Array, pkg: Uint8Array, password: string): Uint8Array {
    const xml = Buffer.from(info.subarray(8)).toString('utf8'); // skip 4-byte version + 4-byte reserved
    const cipher = attr(xml, 'keyData', 'cipherAlgorithm');
    if (cipher && cipher.toUpperCase() !== 'AES') throw new DecryptionError(`encrypted OOXML: unsupported cipher '${cipher}' (only AES)`);
    // We decrypt every segment as CBC; a file declaring another chaining mode (e.g. ChainingModeCFB)
    // would silently produce garbage and be misreported as a wrong password. Reject it as unsupported.
    const chaining = attr(xml, 'keyData', 'cipherChaining');
    if (chaining && chaining.toUpperCase() !== 'CHAININGMODECBC') throw new DecryptionError(`encrypted OOXML: unsupported cipher chaining '${chaining}' (only ChainingModeCBC)`);

    const keyData: AgileKeyBlock = {
        saltValue: Buffer.from(attr(xml, 'keyData', 'saltValue'), 'base64'),
        blockSize: parseInt(attr(xml, 'keyData', 'blockSize'), 10) || 16,
        keyBits: parseInt(attr(xml, 'keyData', 'keyBits'), 10) || 256,
        hashAlgorithm: hashId(attr(xml, 'keyData', 'hashAlgorithm')),
        cipherAlgorithm: 'AES',
    };
    const encSalt = Buffer.from(attr(xml, 'encryptedKey', 'saltValue'), 'base64');
    const encKeyBits = parseInt(attr(xml, 'encryptedKey', 'keyBits'), 10) || 256;
    const encHash = hashId(attr(xml, 'encryptedKey', 'hashAlgorithm'));
    const spinCount = parseInt(attr(xml, 'encryptedKey', 'spinCount'), 10) || 100000;
    // `spinCount`, `keyBits` and `blockSize` come from the attacker-controlled descriptor. Real Office
    // files use spinCount=100000; reject anything absurd so a hostile file cannot make us grind through
    // billions of hash rounds (CPU DoS) before we can even check the password. Office/LibreOffice cap
    // the spin count at ~10 million.
    if (!Number.isFinite(spinCount) || spinCount < 0 || spinCount > 10_000_000) {
        throw new DecryptionError(`encrypted OOXML: implausible spinCount ${spinCount}`);
    }
    for (const bits of [keyData.keyBits, encKeyBits]) {
        if (bits !== 128 && bits !== 192 && bits !== 256) throw new DecryptionError(`encrypted OOXML: unsupported keyBits ${bits}`);
    }
    for (const bs of [keyData.blockSize, parseInt(attr(xml, 'encryptedKey', 'blockSize'), 10) || 16]) {
        if (!Number.isFinite(bs) || bs < 1 || bs > 64) throw new DecryptionError(`encrypted OOXML: implausible blockSize ${bs}`);
    }
    const encVerifierInput = Buffer.from(attr(xml, 'encryptedKey', 'encryptedVerifierHashInput'), 'base64');
    const encVerifierValue = Buffer.from(attr(xml, 'encryptedKey', 'encryptedVerifierHashValue'), 'base64');
    const encKeyValue = Buffer.from(attr(xml, 'encryptedKey', 'encryptedKeyValue'), 'base64');

    const pwHash = agilePasswordHash(encSalt, password, spinCount, encHash);

    // Verify the password before decrypting the (large) package.
    const verifierInputKey = deriveBlockKey(pwHash, BLOCK_VERIFIER_INPUT, encKeyBits, encHash);
    const verifierInput = aesDecryptNoPad(encVerifierInput, verifierInputKey, encSalt.subarray(0, keyData.blockSize), encKeyBits, 'cbc');
    const verifierHashKey = deriveBlockKey(pwHash, BLOCK_VERIFIER_VALUE, encKeyBits, encHash);
    const expectedHash = aesDecryptNoPad(encVerifierValue, verifierHashKey, encSalt.subarray(0, keyData.blockSize), encKeyBits, 'cbc');
    const actualHash = hash(encHash, verifierInput);
    if (!actualHash.subarray(0, actualHash.length).equals(expectedHash.subarray(0, actualHash.length))) throw WRONG_PASSWORD;

    // Recover the package key, then decrypt the package in fixed 4096-byte segments. The decrypted
    // key value is zero-padded to the block size, so trim it to the package cipher's key length
    // (24 bytes for AES-192, which otherwise arrives as 32 and makes createDecipheriv throw).
    const keyValueKey = deriveBlockKey(pwHash, BLOCK_KEY_VALUE, encKeyBits, encHash);
    const secretKey = aesDecryptNoPad(encKeyValue, keyValueKey, encSalt.subarray(0, keyData.blockSize), encKeyBits, 'cbc').subarray(0, keyData.keyBits / 8);

    const pkgView = new DataView(pkg.buffer, pkg.byteOffset, pkg.byteLength);
    const totalSize = Number(pkgView.getBigUint64(0, true));
    const body = pkg.subarray(8);
    const SEGMENT = 4096;
    const out: Buffer[] = [];
    for (let i = 0, off = 0; off < body.length; i++, off += SEGMENT) {
        const block = body.subarray(off, off + SEGMENT);
        const iv = hash(keyData.hashAlgorithm, keyData.saltValue, le32(i)).subarray(0, keyData.blockSize);
        out.push(aesDecryptNoPad(block, secretKey, iv, keyData.keyBits, 'cbc'));
    }
    return Buffer.concat(out).subarray(0, totalSize);
}

// ── standard ─────────────────────────────────────────────────────────────────────

function decryptStandard(info: Uint8Array, pkg: Uint8Array, password: string): Uint8Array {
    const view = new DataView(info.buffer, info.byteOffset, info.byteLength);
    const headerSize = view.getUint32(8, true);
    const headerStart = 12;
    const algId = view.getUint32(headerStart + 8, true);   // 0x660E/0x660F/0x6610 = AES 128/192/256
    let keyBits = view.getUint32(headerStart + 16, true);
    if (algId !== 0x660e && algId !== 0x660f && algId !== 0x6610) {
        throw new DecryptionError(`encrypted OOXML: standard encryption uses an unsupported cipher (algId 0x${algId.toString(16)}); only AES is supported`);
    }
    if (!keyBits) keyBits = algId === 0x660e ? 128 : algId === 0x660f ? 192 : 256;

    // EncryptionVerifier follows the header.
    let p = headerStart + headerSize;
    const saltSize = view.getUint32(p, true); p += 4;
    const salt = Buffer.from(info.subarray(p, p + saltSize)); p += saltSize;
    const encVerifier = Buffer.from(info.subarray(p, p + 16)); p += 16;
    const verifierHashSize = view.getUint32(p, true); p += 4;
    const encVerifierHash = Buffer.from(info.subarray(p, p + 32)); p += 32;

    const key = standardKey(salt, password, keyBits);
    const verifier = aesDecryptNoPad(encVerifier, key, null, keyBits, 'ecb');
    const verifierHash = aesDecryptNoPad(encVerifierHash, key, null, keyBits, 'ecb');
    const actual = hash('sha1', verifier);
    if (!actual.equals(verifierHash.subarray(0, verifierHashSize).subarray(0, actual.length))) throw WRONG_PASSWORD;

    const pkgView = new DataView(pkg.buffer, pkg.byteOffset, pkg.byteLength);
    const totalSize = Number(pkgView.getBigUint64(0, true));
    const body = pkg.subarray(8);
    // The stream is padded to a 16-byte boundary; trim to a whole number of blocks before ECB.
    const aligned = body.subarray(0, body.length - (body.length % 16));
    return aesDecryptNoPad(aligned, key, null, keyBits, 'ecb').subarray(0, totalSize);
}

/** Standard-encryption key derivation ([MS-OFFCRYPTO] §2.3.4.7): 50000 SHA-1 rounds, then X1||X2. */
function standardKey(salt: Buffer, password: string, keyBits: number): Buffer {
    let h = hash('sha1', salt, utf16le(password));
    for (let i = 0; i < 50000; i++) h = hash('sha1', le32(i), h);
    h = hash('sha1', h, le32(0)); // final block
    const x1 = hash('sha1', xorPad(h, 0x36));
    const x2 = hash('sha1', xorPad(h, 0x5c));
    return Buffer.concat([x1, x2]).subarray(0, keyBits / 8);
}

function xorPad(h: Buffer, pad: number): Buffer {
    const buf = Buffer.alloc(64, pad);
    for (let i = 0; i < h.length && i < 64; i++) buf[i] ^= h[i];
    return buf;
}
