/**
 * Generates the DOCX template fixture for the templating tests: `test/files/template/invoice.docx`.
 *
 * Purpose-built to exercise the hard cases of run-aware placeholder substitution:
 *  - `{{name}}` is deliberately SPLIT across three `<w:t>` runs (`{{` | `na` | `me}}`), the Word quirk
 *    the substitution has to survive.
 *  - `{{amount}}` sits in a BOLD run, to prove a value adopts the formatting of its placeholder's run.
 *  - `{{item}}` lives in a TABLE CELL, `{{company}}` in a HEADER part, `{{note}}` receives a multiline
 *    value (-> `<w:br/>`), and `{{missing}}` has no data field (for the onMissing modes).
 *
 * Run: `node scripts/generate-template-fixtures.mjs`
 */
import { zipSync } from 'fflate';
import fs from 'fs';
import path from 'path';

const OUT = path.join('test', 'files', 'template');
const S = s => Buffer.from(s, 'utf8');

const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`;
const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`;

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const r = (t, bold) => `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;

// {{name}} split across three runs on purpose.
const greeting = `<w:p>${r('Dear {{')}${r('na')}${r('me}}, ')}${r('welcome.')}</w:p>`;
const balance = `<w:p>${r('Your balance is ')}${r('{{amount}}', true)}${r(' as of {{date}}.')}</w:p>`;
const table = `<w:tbl><w:tr><w:tc><w:p>${r('Item')}</w:p></w:tc><w:tc><w:p>${r('{{item}}')}</w:p></w:tc></w:tr></w:tbl>`;
const missing = `<w:p>${r('Ref: {{missing}}')}</w:p>`;
const note = `<w:p>${r('Note: {{note}}')}</w:p>`;
const sect = `<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/></w:sectPr>`;
const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${greeting}${balance}${table}${missing}${note}${sect}</w:body></w:document>`;
const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W}><w:p>${r('{{company}} - Confidential')}</w:p></w:hdr>`;

fs.mkdirSync(OUT, { recursive: true });
fs.writeFileSync(path.join(OUT, 'invoice.docx'), Buffer.from(zipSync({
    '[Content_Types].xml': S(ct),
    '_rels/.rels': S(rels),
    'word/_rels/document.xml.rels': S(docRels),
    'word/document.xml': S(document),
    'word/header1.xml': S(header),
})));
console.log('wrote', fs.readdirSync(OUT).join(', '));
