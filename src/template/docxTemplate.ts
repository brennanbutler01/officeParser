/**
 * DOCX templating: fills `{{placeholder}}` tags in a Word document with caller data, preserving all
 * of the document's formatting and structure, and returns a new `.docx`.
 *
 * The one hard part is that Word freely splits a run of text across several `<w:r>`/`<w:t>` elements
 * (an inserted spell-check boundary, a rsid change), so a placeholder typed as `{{name}}` can arrive
 * as `{{`, `nam`, `e}}` in three separate `<w:t>`. We therefore work per paragraph: concatenate its
 * `<w:t>` texts, find placeholders in the joined string, and write the result back so that ordinary
 * text keeps its original run (and formatting) while a placeholder's value adopts the formatting of
 * the run where the placeholder began. Nothing else in the package is touched.
 *
 * @module template/docxTemplate
 */

import { unzipSync, zipSync, Zippable } from 'fflate';
import { TemplateData, TemplateValue } from '../types.js';
import { escapeXml } from '../utils/sanitize.js';

/** The `word/*.xml` parts that carry body text a placeholder could live in. */
const TEXT_PART = /^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/;

/** Escapes a regex metacharacter run so custom delimiters can be used literally. */
function escapeRegex(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** Renders one data value to the inner XML of a `<w:t>`, turning newlines into `<w:br/>`. */
function valueToRunXml(value: TemplateValue): string {
    let text: string;
    if (value == null) text = '';
    else if (value instanceof Date) text = value.toISOString();
    else text = String(value);
    // A newline in a value becomes a real line break: close this <w:t>, emit <w:br/>, reopen a <w:t>.
    return escapeXml(text).replace(/\r?\n/g, '</w:t><w:br/><w:t xml:space="preserve">');
}

/**
 * Replaces placeholders within a single `<w:p>...</w:p>` paragraph. `resolve` returns the run-XML to
 * substitute for a key, or `null` to leave the placeholder text as-is (the `onMissing: 'keep'` case).
 */
function replaceInParagraph(pXml: string, phRe: RegExp, resolve: (key: string) => string | null): string {
    const WT = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g;
    const segs: { start: number; end: number; inner: string }[] = [];
    let m: RegExpExecArray | null;
    while ((m = WT.exec(pXml))) segs.push({ start: m.index, end: m.index + m[0].length, inner: m[1] });
    if (!segs.length) return pXml;

    const joined = segs.map(s => s.inner).join('');
    phRe.lastIndex = 0;
    if (!phRe.test(joined)) return pXml; // no placeholder in this paragraph: leave it exactly as-is

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
    let out = pXml;
    for (let si = segs.length - 1; si >= 0; si--) {
        out = out.slice(0, segs[si].start) + `<w:t xml:space="preserve">${newInner[si]}</w:t>` + out.slice(segs[si].end);
    }
    return out;
}

/** Applies placeholder replacement to every paragraph of one XML part. */
function replaceInPart(xml: string, phRe: RegExp, resolve: (key: string) => string | null): string {
    // Paragraphs never nest, so a non-greedy match to the first </w:p> is safe.
    return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, p => replaceInParagraph(p, phRe, resolve));
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
    const phRe = new RegExp(`${escapeRegex(opts.start)}\\s*([\\w.\\-]+)\\s*${escapeRegex(opts.end)}`, 'g');
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

/** Unzips a docx and confirms it is one (has `word/document.xml`); returns null when it is not a docx. */
export function openDocx(bytes: Uint8Array): Record<string, Uint8Array> | null {
    let entries: Record<string, Uint8Array>;
    try { entries = unzipSync(bytes); } catch { return null; }
    return entries['word/document.xml'] ? entries : null;
}
