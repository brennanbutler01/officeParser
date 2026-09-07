/**
 * Generates encrypted-document fixtures for the crypto tests. INDEPENDENT of src/crypto (written
 * straight from [MS-OFFCRYPTO] and the ODF 1.2 spec) so the parser's round-trip through it is a real
 * cross-check, not a mirror of the same code. Node's own `crypto` does the AES/hash primitives, and
 * for OOXML a tiny CFB writer wraps the two encryption streams.
 *
 * Run: `node scripts/generate-crypto-fixtures.mjs`  ->  test/files/encrypted/{agile.docx,standard.docx,encrypted.odt}
 * Password for all three: `test123`.
 */
import { createHash, createCipheriv, pbkdf2Sync, randomBytes } from 'crypto';
import { unzipSync, zipSync, deflateSync } from 'fflate';
import fs from 'fs';
import path from 'path';

const PASSWORD = 'test123';
const OUT_DIR = path.join('test', 'files', 'encrypted');
const u32 = n => { const b = Buffer.alloc(4); b.writeUInt32LE(n >>> 0); return b; };
const u64 = n => { const b = Buffer.alloc(8); b.writeBigUInt64LE(BigInt(n)); return b; };
const utf16 = s => Buffer.from(s, 'utf16le');
const H = (algo, ...parts) => { const h = createHash(algo); for (const p of parts) h.update(p); return h.digest(); };
const padTo = (buf, block) => { const r = buf.length % block; return r ? Buffer.concat([buf, Buffer.alloc(block - r)]) : buf; };
const aesEnc = (data, key, iv, bits, mode) => { const c = createCipheriv(`aes-${bits}-${mode}`, key, iv); c.setAutoPadding(false); return Buffer.concat([c.update(data), c.final()]); };

// ── minimal CFB (v3, 512-byte sectors) writer ──────────────────────────────────
function buildCfb(streams) {
    const SECTOR = 512, MINI = 64, CUTOFF = 4096, END = 0xfffffffe, FREE = 0xffffffff, FATSECT = 0xfffffffd;
    const mini = streams.filter(s => s.data.length < CUTOFF);
    const regular = streams.filter(s => s.data.length >= CUTOFF);

    // Mini stream: concat of the small streams, each padded to a mini-sector.
    const miniLoc = {};
    let miniParts = [], miniSectors = 0;
    for (const s of mini) {
        const n = Math.ceil(s.data.length / MINI) || 1;
        miniLoc[s.name] = { start: miniSectors, size: s.data.length };
        const padded = Buffer.alloc(n * MINI); s.data.copy(padded); miniParts.push(padded); miniSectors += n;
    }
    const miniStream = Buffer.concat(miniParts);
    const miniFat = [];
    { let idx = 0; for (const s of mini) { const n = Math.ceil(s.data.length / MINI) || 1; for (let k = 0; k < n; k++) { miniFat.push(k === n - 1 ? END : idx + 1); idx++; } } }

    const dirEntries = [{ name: 'Root Entry', type: 5 }, ...mini.map(s => ({ name: s.name, type: 2 })), ...regular.map(s => ({ name: s.name, type: 2 }))];
    const dirSectors = Math.ceil(dirEntries.length / 4) || 1;
    const miniFatSectors = Math.max(1, Math.ceil((miniFat.length * 4) / SECTOR));
    const miniStreamSectors = Math.ceil(miniStream.length / SECTOR);
    const regSizes = regular.map(s => Math.ceil(s.data.length / SECTOR));
    const nonFat = dirSectors + miniFatSectors + miniStreamSectors + regSizes.reduce((a, b) => a + b, 0);
    let fatSectors = 1; while (Math.ceil((nonFat + fatSectors) / (SECTOR / 4)) > fatSectors) fatSectors++;

    let cur = 0;
    const fatStart = cur; cur += fatSectors;
    const dirStart = cur; cur += dirSectors;
    const miniFatStart = cur; cur += miniFatSectors;
    const miniStreamStart = cur; cur += miniStreamSectors;
    const regStarts = regSizes.map(n => { const s = cur; cur += n; return s; });

    const fat = new Array(fatSectors * (SECTOR / 4)).fill(FREE);
    const chain = (start, count) => { for (let i = 0; i < count; i++) fat[start + i] = i === count - 1 ? END : start + i + 1; };
    for (let i = 0; i < fatSectors; i++) fat[fatStart + i] = FATSECT;
    chain(dirStart, dirSectors); chain(miniFatStart, miniFatSectors);
    if (miniStreamSectors) chain(miniStreamStart, miniStreamSectors);
    regSizes.forEach((n, i) => { if (n) chain(regStarts[i], n); });

    const header = Buffer.alloc(SECTOR);
    Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(header, 0);
    header.writeUInt16LE(0x003e, 24); header.writeUInt16LE(0x0003, 26); header.writeUInt16LE(0xfffe, 28);
    header.writeUInt16LE(9, 30); header.writeUInt16LE(6, 32);
    header.writeUInt32LE(fatSectors, 44); header.writeUInt32LE(dirStart, 48);
    header.writeUInt32LE(CUTOFF, 56);
    header.writeUInt32LE(miniStreamSectors ? miniFatStart : END, 60); header.writeUInt32LE(miniFatSectors, 64);
    header.writeUInt32LE(END, 68); header.writeUInt32LE(0, 72);
    for (let i = 0; i < 109; i++) header.writeUInt32LE(i < fatSectors ? fatStart + i : FREE, 76 + i * 4);

    const dir = Buffer.alloc(dirSectors * SECTOR);
    dirEntries.forEach((e, i) => {
        const off = i * 128;
        const nb = Buffer.from(e.name + '\0', 'utf16le'); nb.copy(dir, off);
        dir.writeUInt16LE(nb.length, off + 64); dir.writeUInt8(e.type, off + 66);
        dir.writeUInt32LE(FREE, off + 68); dir.writeUInt32LE(FREE, off + 72); dir.writeUInt32LE(FREE, off + 76);
        if (e.type === 5) { dir.writeUInt32LE(miniStreamSectors ? miniStreamStart : END, off + 116); dir.writeUInt32LE(miniStream.length, off + 120); }
        else if (miniLoc[e.name]) { dir.writeUInt32LE(miniLoc[e.name].start, off + 116); dir.writeUInt32LE(miniLoc[e.name].size, off + 120); }
        else { const ri = regular.findIndex(s => s.name === e.name); dir.writeUInt32LE(regSizes[ri] ? regStarts[ri] : END, off + 116); dir.writeUInt32LE(regular[ri].data.length, off + 120); }
    });

    const fatBuf = Buffer.alloc(fatSectors * SECTOR, 0xff); fat.forEach((v, i) => fatBuf.writeUInt32LE(v >>> 0, i * 4));
    const miniFatBuf = Buffer.alloc(miniFatSectors * SECTOR, 0xff); miniFat.forEach((v, i) => miniFatBuf.writeUInt32LE(v >>> 0, i * 4));
    const miniStreamPadded = Buffer.alloc(miniStreamSectors * SECTOR); miniStream.copy(miniStreamPadded);
    const regBufs = regular.map(s => { const b = Buffer.alloc(Math.ceil(s.data.length / SECTOR) * SECTOR); s.data.copy(b); return b; });
    return Buffer.concat([header, fatBuf, dir, miniFatBuf, miniStreamPadded, ...regBufs]);
}

// ── OOXML agile encryption ─────────────────────────────────────────────────────
const B_VER_IN = Buffer.from([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const B_VER_VAL = Buffer.from([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const B_KEY_VAL = Buffer.from([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);

function agileEncrypt(plain, password) {
    const algo = 'sha512', keyBits = 256, blockSize = 16, spinCount = 100000;
    const keySalt = randomBytes(16), encSalt = randomBytes(16);
    let h = H(algo, encSalt, utf16(password));
    for (let i = 0; i < spinCount; i++) h = H(algo, u32(i), h);
    const blockKey = (bk) => H(algo, h, bk).subarray(0, keyBits / 8);

    const secretKey = randomBytes(keyBits / 8);
    const encKeyValue = aesEnc(secretKey, blockKey(B_KEY_VAL), encSalt, keyBits, 'cbc');
    const verifierInput = randomBytes(16);
    const encVerIn = aesEnc(verifierInput, blockKey(B_VER_IN), encSalt, keyBits, 'cbc');
    const encVerVal = aesEnc(padTo(H(algo, verifierInput), blockSize), blockKey(B_VER_VAL), encSalt, keyBits, 'cbc');

    const SEG = 4096, segs = [];
    for (let off = 0, i = 0; off < plain.length; off += SEG, i++) {
        const iv = H(algo, keySalt, u32(i)).subarray(0, blockSize);
        segs.push(aesEnc(padTo(plain.subarray(off, off + SEG), blockSize), secretKey, iv, keyBits, 'cbc'));
    }
    const encryptedPackage = Buffer.concat([u64(plain.length), ...segs]);

    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
        + `<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">`
        + `<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${keySalt.toString('base64')}"/>`
        + `<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">`
        + `<p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" `
        + `saltValue="${encSalt.toString('base64')}" encryptedVerifierHashInput="${encVerIn.toString('base64')}" encryptedVerifierHashValue="${encVerVal.toString('base64')}" encryptedKeyValue="${encKeyValue.toString('base64')}"/>`
        + `</keyEncryptor></keyEncryptors></encryption>`;
    const encryptionInfo = Buffer.concat([Buffer.from([0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00]), Buffer.from(xml, 'utf8')]);
    return buildCfb([{ name: 'EncryptionInfo', data: encryptionInfo }, { name: 'EncryptedPackage', data: encryptedPackage }]);
}

// ── OOXML standard encryption (AES-128 ECB) ────────────────────────────────────
function standardEncrypt(plain, password) {
    const keyBits = 128, salt = randomBytes(16);
    let h = H('sha1', salt, utf16(password));
    for (let i = 0; i < 50000; i++) h = H('sha1', u32(i), h);
    h = H('sha1', h, u32(0));
    const xorPad = pad => { const b = Buffer.alloc(64, pad); for (let i = 0; i < h.length; i++) b[i] ^= h[i]; return b; };
    const key = Buffer.concat([H('sha1', xorPad(0x36)), H('sha1', xorPad(0x5c))]).subarray(0, keyBits / 8);

    const verifier = randomBytes(16);
    const encVerifier = aesEnc(verifier, key, null, keyBits, 'ecb');
    // SHA-1 is 20 bytes; the field is padded to a 32-byte (2-block) boundary before ECB.
    const verifierHash32 = padTo(H('sha1', verifier), 32);
    const encVerifierHash = aesEnc(verifierHash32, key, null, keyBits, 'ecb');
    const encryptedPackage = Buffer.concat([u64(plain.length), aesEnc(padTo(plain, 16), key, null, keyBits, 'ecb')]);

    const csp = Buffer.from('Microsoft Enhanced RSA and AES Cryptographic Provider\0', 'utf16le');
    const encHeader = Buffer.concat([u32(0x24), u32(0), u32(0x660e), u32(0x8004), u32(keyBits), u32(0x18), u32(0), u32(0), csp]);
    const verifierBlock = Buffer.concat([u32(16), salt, encVerifier, u32(20), encVerifierHash]);
    const info = Buffer.concat([u32(0x00020003), u32(0x24), u32(encHeader.length), encHeader, verifierBlock]);
    return buildCfb([{ name: 'EncryptionInfo', data: info }, { name: 'EncryptedPackage', data: encryptedPackage }]);
}

// ── ODF encryption (AES-256-CBC per entry) ─────────────────────────────────────
function odfEncrypt(odtBuffer, password) {
    const entries = unzipSync(odtBuffer);
    const startKey = H('sha256', Buffer.from(password, 'utf8'));
    const out = {};
    const manifestEntries = [`<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>`];

    for (const [name, data] of Object.entries(entries)) {
        if (name === 'mimetype' || name === 'META-INF/manifest.xml' || name.endsWith('/')) continue;
        const compressed = Buffer.from(deflateSync(data, { level: 9 })); // fflate's deflateSync is raw deflate
        const iv = randomBytes(16), salt = randomBytes(16), iter = 100000;
        const key = pbkdf2Sync(startKey, salt, iter, 32, 'sha1');
        const enc = aesEnc(padTo(compressed, 16), key, iv, 256, 'cbc');
        const checksum = H('sha256', compressed.subarray(0, 1024)).toString('base64');
        out[name] = [enc, { level: 0 }]; // STORED: the bytes are already encrypted
        manifestEntries.push(
            `<manifest:file-entry manifest:full-path="${name}" manifest:media-type="" manifest:size="${data.length}">`
            + `<manifest:encryption-data manifest:checksum-type="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k" manifest:checksum="${checksum}">`
            + `<manifest:algorithm manifest:algorithm-name="http://www.w3.org/2001/04/xmlenc#aes256-cbc" manifest:initialisation-vector="${iv.toString('base64')}"/>`
            + `<manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:key-size="32" manifest:iteration-count="${iter}" manifest:salt="${salt.toString('base64')}"/>`
            + `<manifest:start-key-generation manifest:start-key-generation-name="http://www.w3.org/2000/09/xmldsig#sha256" manifest:key-size="32"/>`
            + `</manifest:encryption-data></manifest:file-entry>`);
    }
    const manifest = `<?xml version="1.0" encoding="UTF-8"?>`
        + `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">${manifestEntries.join('')}</manifest:manifest>`;
    return Buffer.from(zipSync({
        mimetype: [entries['mimetype'], { level: 0 }],
        ...out,
        'META-INF/manifest.xml': [Buffer.from(manifest, 'utf8'), { level: 0 }],
    }));
}

// Tiny, purpose-built source documents (a few KB) carrying one recognizable marker, so the
// committed encrypted fixtures stay small rather than baking in the large shared test.docx/test.odt.
const MARKER = 'SECRET CONTENT 42';

function minimalDocx(marker) {
    const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${marker}</w:t></w:r></w:p></w:body></w:document>`;
    return Buffer.from(zipSync({
        '[Content_Types].xml': Buffer.from(ct, 'utf8'),
        '_rels/.rels': Buffer.from(rels, 'utf8'),
        'word/document.xml': Buffer.from(doc, 'utf8'),
    }));
}

function minimalOdt(marker) {
    // Pad content.xml with many varied paragraphs so it compresses to well over 1 KiB: that makes the
    // ODF decryptor's checksum branch (only used above ~1 KiB compressed) run against this fixture.
    let filler = '';
    for (let i = 0; i < 400; i++) filler += `<text:p>Line ${i} lorem ipsum dolor sit amet consectetur ${i * 7} adipiscing ${(i * 131) % 1000}</text:p>`;
    const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2"><office:body><office:text><text:p>${marker}</text:p>${filler}</office:text></office:body></office:document-content>`;
    const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2"/>`;
    const manifest = `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>`;
    return Buffer.from(zipSync({
        mimetype: [Buffer.from('application/vnd.oasis.opendocument.text', 'utf8'), { level: 0 }],
        'content.xml': Buffer.from(content, 'utf8'),
        'styles.xml': Buffer.from(styles, 'utf8'),
        'META-INF/manifest.xml': Buffer.from(manifest, 'utf8'),
    }));
}

// ── main ───────────────────────────────────────────────────────────────────────
fs.mkdirSync(OUT_DIR, { recursive: true });
fs.writeFileSync(path.join(OUT_DIR, 'agile.docx'), agileEncrypt(minimalDocx(MARKER), PASSWORD));
fs.writeFileSync(path.join(OUT_DIR, 'standard.docx'), standardEncrypt(minimalDocx(MARKER), PASSWORD));
fs.writeFileSync(path.join(OUT_DIR, 'encrypted.odt'), odfEncrypt(minimalOdt(MARKER), PASSWORD));
console.log('wrote', fs.readdirSync(OUT_DIR).join(', '), `(password: ${PASSWORD}, marker: "${MARKER}")`);
