/**
 * DOCX templating: fills `{{placeholder}}` tags in a Word document with caller data, preserving all
 * of the document's formatting and structure, and returns a new `.docx`.
 *
 * The one hard part is that Word freely splits a run of text across several `<w:r>`/`<w:t>` elements
 * (an inserted spell-check boundary, a rsid change), so a placeholder typed as `{{name}}` can arrive
 * as `{{`, `nam`, `e}}` in three separate `<w:t>`. We therefore work per paragraph: concatenate its
 * `<w:t>` texts, find placeholders in the joined string, and write the result back so that ordinary
 * text keeps its original run (and formatting) while a placeholder's value adopts the formatting of
 * the run where the placeholder began. Paragraphs can nest (a text box's `<w:p>` lives inside a run
 * of the outer `<w:p>`), so we scope substitution to the text *between* paragraph-boundary tags,
 * which both handles nesting and never lets a placeholder span a paragraph boundary.
 *
 * @module template/docxTemplate
 */

import { zipSync, Zippable } from 'fflate';
import { DecompressionLimits, OfficeParserConfig, TemplateData, TemplateValue } from '../types.js';
import { escapeXml, stripInvalidXmlChars } from '../utils/sanitize.js';
import { extractFiles } from '../utils/zipUtils.js';

/** The `word/*.xml` parts that carry body text a placeholder could live in. */
const TEXT_PART = /^word\/(document\d*|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/;

/** Escapes a regex metacharacter run so custom delimiters can be used literally. */
function escapeRegex(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Renders one data value to the inner XML of a `<w:t>`, turning newlines into `<w:br/>`. */
function valueToRunXml(value: TemplateValue): string {
    let text: string;
    if (value == null) text = '';
    else if (value instanceof Date) text = isNaN(value.getTime()) ? '' : value.toISOString();
    else text = String(value);
    // Strip XML-illegal control chars (DB-sourced data often carries them) BEFORE escaping, exactly
    // as the DOCX generator does; then a newline becomes a real line break: close this <w:t>, emit
    // <w:br/>, reopen a <w:t>.
    return escapeXml(stripInvalidXmlChars(text)).replace(/\r?\n/g, '</w:t><w:br/><w:t xml:space="preserve">');
}

/**
 * Replaces placeholders within one chunk of run-level XML (the content between two paragraph-boundary
 * tags). `resolve` returns the run-XML to substitute for a key, or `null` to leave the placeholder
 * text as-is (the `onMissing: 'keep'` case).
 */
function replaceInChunk(chunk: string, phRe: RegExp, resolve: (key: string) => string | null): string {
    // A `<w:t>` is either self-closing (`<w:t/>`, emitted by the OpenXML SDK / POI / docx4j for an
    // empty run) or a normal `<w:t ...>text</w:t>`; capture the inner text of the latter.
    const WT = /<w:t\b[^>]*?(?:\/>|>([\s\S]*?)<\/w:t>)/g;
    const segs: { start: number; end: number; inner: string }[] = [];
    let m: RegExpExecArray | null;
    while ((m = WT.exec(chunk))) segs.push({ start: m.index, end: m.index + m[0].length, inner: m[1] ?? '' });
    if (!segs.length) return chunk;

    const joined = segs.map(s => s.inner).join('');
    phRe.lastIndex = 0;
    if (!phRe.test(joined)) return chunk; // no placeholder here: leave it exactly as-is

    // Map each character position in the joined text back to the <w:t> segment it came from.
    const posSeg = new Int32Array(joined.length);
    let ci = 0;
    segs.forEach((s, si) => { for (let k = 0; k < s.inner.length; k++) posSeg[ci++] = si; });

    phRe.lastIndex = 0;
    const matches: { start: number; end: number; key: string }[] = [];
    while ((m = phRe.exec(joined))) matches.push({ start: m.index, end: m.index + m[0].length, key: m[1] });

    // Rebuild each segment's inner text. Ordinary characters stay in their own segment; a placeholder's
    // replacement is attributed wholesale to the segment where its opening delimiter began.
    const newInner = segs.map(() => '');
    let mi = 0;
    for (let p = 0; p < joined.length;) {
        if (mi < matches.length && p === matches[mi].start) {
            const mt = matches[mi++];
            const val = resolve(mt.key);
            newInner[posSeg[p]] += val === null ? joined.slice(mt.start, mt.end) : val;
            p = mt.end;
        } else {
            newInner[posSeg[p]] += joined[p];
            p++;
        }
    }

    // Splice the rebuilt `<w:t>` elements back in, from last to first so earlier offsets stay valid.
    let out = chunk;
    for (let si = segs.length - 1; si >= 0; si--) {
        out = out.slice(0, segs[si].start) + `<w:t xml:space="preserve">${newInner[si]}</w:t>` + out.slice(segs[si].end);
    }
    return out;
}

/**
 * Applies placeholder replacement to one XML part. The part is split at every paragraph-boundary tag
 * (`<w:p ...>`, `</w:p>`, `<w:p/>`); each piece between boundaries is one paragraph's own run content
 * (nested paragraphs, e.g. text boxes, become their own pieces), so joining a piece's `<w:t>` never
 * crosses a paragraph boundary. This is linear (no `[\s\S]*?`-to-`</w:p>` backtracking).
 */
function replaceInPart(xml: string, phRe: RegExp, resolve: (key: string) => string | null): string {
    const pieces = xml.split(/(<w:p\b[^>]*\/>|<w:p\b[^>]*>|<\/w:p>)/g);
    for (let i = 0; i < pieces.length; i++) {
        // Boundary tags (odd indices) have no `<w:t>`, so processing them is a harmless no-op; process
        // every piece uniformly rather than tracking parity.
        pieces[i] = replaceInChunk(pieces[i], phRe, resolve);
    }
    return pieces.join('');
}

/**
 * Renders one `.docx` from a template zip and a single data map. `entries` is the already-unzipped
 * template (shared across a batch); this function copies it, rewrites the text parts, and re-zips.
 */
export function renderDocxTemplate(
    entries: Record<string, Uint8Array>,
    data: TemplateData,
    opts: { start: string; end: string; onMissing: 'keep' | 'empty' | 'error'; mtime: Date; onFieldMissing: (key: string) => never },
): Uint8Array {
    // Match delimiters against the raw XML inner text, where Word stores e.g. `<<` as `&lt;&lt;`; so
    // build the pattern from the XML-escaped delimiters. Placeholder names allow Unicode letters/digits.
    const start = escapeRegex(escapeXml(opts.start));
    const end = escapeRegex(escapeXml(opts.end));
    const phRe = new RegExp(`${start}\\s*([\\p{L}\\p{N}_.\\-]+)\\s*${end}`, 'gu');
    const resolve = (key: string): string | null => {
        if (Object.prototype.hasOwnProperty.call(data, key)) return valueToRunXml(data[key]);
        if (opts.onMissing === 'empty') return '';
        if (opts.onMissing === 'error') opts.onFieldMissing(key);
        return null; // 'keep'
    };

    const out: Zippable = {};
    for (const [name, bytes] of Object.entries(entries)) {
        if (TEXT_PART.test(name)) {
            const xml = Buffer.from(bytes).toString('utf8');
            out[name] = [Buffer.from(replaceInPart(xml, phRe, resolve), 'utf8'), { mtime: opts.mtime }];
        } else {
            out[name] = [bytes, { mtime: opts.mtime }];
        }
    }
    return zipSync(out);
}

/**
 * Unzips a docx under decompression limits and confirms it is one (has `word/document.xml`); returns
 * null when it is not a docx. Uses the parser's streaming `extractFiles` guard so a zip-bomb template
 * is bounded exactly as an ordinary parsed document is.
 */
export async function openDocx(bytes: Uint8Array, limits?: DecompressionLimits, config?: OfficeParserConfig): Promise<Record<string, Uint8Array> | null> {
    const src = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let files;
    try { files = await extractFiles(src, () => true, limits ?? {}, config); }
    catch { return null; }
    const map: Record<string, Uint8Array> = {};
    for (const f of files) map[f.path] = f.content;
    return map['word/document.xml'] ? map : null;
}
