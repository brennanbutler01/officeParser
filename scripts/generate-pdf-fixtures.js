/**
 * Generates the deterministic PDF fixtures used by the parser test suite.
 *
 * These are written as raw PDF syntax (no external tooling) so the bytes are stable across machines
 * and CI. Run once and commit the output:
 *
 *     node scripts/generate-pdf-fixtures.js
 *
 * Outputs (under test/files/pdf/):
 *   - columns_untagged.pdf : two text columns, no structure tree, for reading-order tests.
 *   - rotated.pdf          : the same content on a /Rotate 90 page.
 *   - encrypted.pdf        : single page, RC4 40-bit standard security, user password "test123".
 *
 * The encryption is the PDF standard security handler, revision 2 (RC4, 40-bit). RC4 is implemented
 * inline because modern OpenSSL/Node no longer expose it, and qpdf is not assumed to be installed.
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const OUT_DIR = path.join(__dirname, '..', 'test', 'files', 'pdf');

// ── raw PDF assembly ─────────────────────────────────────────────────────────

/** Escapes a string for a PDF literal `( )` string. */
function esc(s) {
    return s.replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');
}

/** Builds a content stream that draws each { x, y, size, text } line with Helvetica. */
function contentStream(lines) {
    let s = 'BT\n';
    for (const l of lines) {
        s += `/F1 ${l.size} Tf\n1 0 0 1 ${l.x} ${l.y} Tm\n(${esc(l.text)}) Tj\n`;
    }
    s += 'ET\n';
    return Buffer.from(s, 'latin1');
}

/**
 * Assembles a one-page PDF from a content-stream buffer.
 * @param {Buffer} content
 * @param {{ rotate?: number, encrypt?: { userPassword: string } }} opts
 */
function buildPdf(content, opts = {}) {
    const rotate = opts.rotate || 0;
    const objects = []; // objects[i] = Buffer for object (i+1)

    objects[0] = Buffer.from('<< /Type /Catalog /Pages 2 0 R >>', 'latin1');
    objects[1] = Buffer.from('<< /Type /Pages /Kids [3 0 R] /Count 1 >>', 'latin1');
    const rotateEntry = rotate ? ` /Rotate ${rotate}` : '';
    objects[2] = Buffer.from(
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792]${rotateEntry} ` +
        `/Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>`, 'latin1');
    // object 4 (content) filled below, possibly encrypted
    objects[4] = Buffer.from('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>', 'latin1');

    let encryptRef = '';
    let idHex = '30313233343536373839414243444546'; // fixed 16-byte /ID
    if (opts.encrypt) {
        const enc = buildStandardEncryption(opts.encrypt.userPassword, idHex);
        objects[5] = Buffer.from(
            `<< /Filter /Standard /V 1 /R 2 /O <${enc.O}> /U <${enc.U}> /P -1 >>`, 'latin1');
        // encrypt the content stream with the object key for object 4, gen 0
        content = enc.encryptObject(content, 4, 0);
        encryptRef = ` /Encrypt 6 0 R`;
    }

    const streamObj = Buffer.concat([
        Buffer.from(`<< /Length ${content.length} >>\nstream\n`, 'latin1'),
        content,
        Buffer.from('\nendstream', 'latin1'),
    ]);
    objects[3] = streamObj;

    // Serialize with an xref table.
    const header = Buffer.from('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n', 'latin1');
    let body = header;
    const offsets = [];
    for (let i = 0; i < objects.length; i++) {
        offsets[i] = body.length;
        body = Buffer.concat([body, Buffer.from(`${i + 1} 0 obj\n`, 'latin1'), objects[i], Buffer.from('\nendobj\n', 'latin1')]);
    }
    const xrefOffset = body.length;
    let xref = `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
    for (let i = 0; i < objects.length; i++) {
        xref += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`;
    }
    let trailer = `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R${encryptRef} /ID [<${idHex}> <${idHex}>] >>\nstartxref\n${xrefOffset}\n%%EOF`;
    return Buffer.concat([body, Buffer.from(xref, 'latin1'), Buffer.from(trailer, 'latin1')]);
}

// ── RC4 + PDF standard security handler (revision 2, 40-bit) ─────────────────

function rc4(key, data) {
    const s = new Uint8Array(256);
    for (let i = 0; i < 256; i++) s[i] = i;
    let j = 0;
    for (let i = 0; i < 256; i++) {
        j = (j + s[i] + key[i % key.length]) & 0xff;
        [s[i], s[j]] = [s[j], s[i]];
    }
    const out = Buffer.alloc(data.length);
    let a = 0, b = 0;
    for (let k = 0; k < data.length; k++) {
        a = (a + 1) & 0xff;
        b = (b + s[a]) & 0xff;
        [s[a], s[b]] = [s[b], s[a]];
        out[k] = data[k] ^ s[(s[a] + s[b]) & 0xff];
    }
    return out;
}

const PAD = Buffer.from([
    0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41, 0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
    0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80, 0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A,
]);

function padPassword(pw) {
    const bytes = Buffer.from(pw, 'latin1');
    return Buffer.concat([bytes, PAD]).subarray(0, 32);
}

function md5(...buffers) {
    const h = crypto.createHash('md5');
    for (const b of buffers) h.update(b);
    return h.digest();
}

/** Builds /O, /U and a per-object encryptor for revision 2 (40-bit RC4). */
function buildStandardEncryption(userPassword, idHex) {
    const userPad = padPassword(userPassword);
    const ownerPad = padPassword(userPassword); // owner == user for the fixture
    const ownerKey = md5(ownerPad).subarray(0, 5);
    const O = rc4(ownerKey, userPad);

    const pLe = Buffer.alloc(4);
    pLe.writeInt32LE(-1, 0);
    const idBytes = Buffer.from(idHex, 'hex');
    const fileKey = md5(userPad, O, pLe, idBytes).subarray(0, 5);
    const U = rc4(fileKey, PAD);

    function encryptObject(data, objNum, gen) {
        const ext = Buffer.from([objNum & 0xff, (objNum >> 8) & 0xff, (objNum >> 16) & 0xff, gen & 0xff, (gen >> 8) & 0xff]);
        const objKey = md5(fileKey, ext).subarray(0, Math.min(fileKey.length + 5, 16));
        return rc4(objKey, data);
    }

    return { O: O.toString('hex'), U: U.toString('hex'), encryptObject };
}

// ── fixture definitions ──────────────────────────────────────────────────────

function twoColumnLines() {
    const lines = [];
    const leftWords = ['ALPHA', 'BETA', 'GAMMA', 'DELTA', 'EPSILON', 'OMEGA'];
    const rightWords = ['ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'NINE'];
    let y = 720;
    for (let i = 0; i < leftWords.length; i++) {
        lines.push({ x: 72, y, size: 12, text: `LEFT COLUMN ${leftWords[i]}` });
        lines.push({ x: 330, y, size: 12, text: `RIGHT COLUMN ${rightWords[i]}` });
        y -= 24;
    }
    return lines;
}

function main() {
    fs.mkdirSync(OUT_DIR, { recursive: true });

    const columns = contentStream(twoColumnLines());
    fs.writeFileSync(path.join(OUT_DIR, 'columns_untagged.pdf'), buildPdf(columns));
    fs.writeFileSync(path.join(OUT_DIR, 'rotated.pdf'), buildPdf(columns, { rotate: 90 }));

    const secret = contentStream([{ x: 72, y: 700, size: 24, text: 'SECRET CONTENT' }]);
    fs.writeFileSync(path.join(OUT_DIR, 'encrypted.pdf'), buildPdf(secret, { encrypt: { userPassword: 'test123' } }));

    console.log('Wrote fixtures to', OUT_DIR);
}

main();
