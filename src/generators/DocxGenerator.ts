import { zipSync, Zippable } from 'fflate';
import { ConversionResult, DocxGeneratorConfig, GeneratorConfig, ImageMode, OfficeContentNode, OfficeMetadata, OfficeParserAST, OfficeWarningType, TextFormatting } from '../types.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { escapeXml, isSafeStyleMapTag, sanitizeDocxUrl, stripInvalidXmlChars } from '../utils/sanitize.js';
import { BaseGenerator } from './BaseGenerator.js';

/** Decodes base64 to bytes, cross-env (atob exists in Node and browsers). */
function decodeBase64(b64: string): Uint8Array {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

/** Image MIME to file extension for a packaged media part. */
const MIME_EXT: Record<string, string> = {
    'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/png': 'png', 'image/gif': 'gif',
    'image/svg+xml': 'svg', 'image/webp': 'webp', 'image/bmp': 'bmp', 'image/tiff': 'tiff',
};

/** Page dimensions in twips (1/20 pt), portrait. */
const PAGE_SIZES: Record<string, { w: number; h: number }> = {
    A4: { w: 11906, h: 16838 },
    Letter: { w: 12240, h: 15840 },
    Legal: { w: 12240, h: 20160 },
};

const EMU_PER_PT = 12700;
const EMU_PER_IN = 914400;
const TWIPS_PER_PT = 20;

/** Parses a CSS-ish length ('12pt', '1in', '2cm', '10mm', '96px', or a bare number = px) to points. */
function lengthToPt(value: string | number | undefined): number | null {
    if (value == null) return null;
    if (typeof value === 'number') return Number.isFinite(value) ? value * 0.75 : null; // px -> pt
    const m = /^\s*(-?[\d.]+)\s*(pt|in|cm|mm|px|%)?\s*$/.exec(value);
    if (!m) return null;
    const n = parseFloat(m[1]);
    if (!Number.isFinite(n)) return null;
    switch (m[2]) {
        case 'pt': return n;
        case 'in': return n * 72;
        case 'cm': return n * 28.3465;
        case 'mm': return n * 2.83465;
        case '%': return null; // handled by the caller (fraction of content width)
        default: return n * 0.75; // px at 96dpi
    }
}

/** Sniffs intrinsic pixel size from the header of PNG / JPEG / GIF bytes. */
function sniffImageSize(bytes: Uint8Array): { w: number; h: number } | null {
    if (bytes.length < 24) return null;
    // PNG: signature then IHDR width/height (big-endian) at offset 16.
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47) {
        const w = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
        const h = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
        if (w > 0 && h > 0) return { w, h };
    }
    // GIF: logical screen width/height (little-endian) at offset 6.
    if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
        const w = bytes[6] | (bytes[7] << 8);
        const h = bytes[8] | (bytes[9] << 8);
        if (w > 0 && h > 0) return { w, h };
    }
    // JPEG: scan segments for a SOF marker carrying height/width.
    if (bytes[0] === 0xFF && bytes[1] === 0xD8) {
        let o = 2;
        while (o + 9 < bytes.length) {
            if (bytes[o] !== 0xFF) { o++; continue; }
            const marker = bytes[o + 1];
            if (marker >= 0xC0 && marker <= 0xCF && marker !== 0xC4 && marker !== 0xC8 && marker !== 0xCC) {
                const h = (bytes[o + 5] << 8) | bytes[o + 6];
                const w = (bytes[o + 7] << 8) | bytes[o + 8];
                if (w > 0 && h > 0) return { w, h };
                return null;
            }
            const len = (bytes[o + 2] << 8) | bytes[o + 3];
            if (len <= 0) return null;
            o += 2 + len;
        }
    }
    return null;
}

/** Admonition type to accent color (matches the HTML generator's palette family). */
const ADMONITION_COLOR: Record<string, string> = {
    note: '0969DA', tip: '1A7F37', important: '8250DF', warning: '9A6700', caution: 'CF222E',
};

/** Escapes text content for an XML part after stripping XML-illegal characters. */
function xmlText(s: string | undefined): string {
    return escapeXml(stripInvalidXmlChars(s ?? ''));
}

/** Validates a `#RRGGBB`/`RRGGBB`/`#RGB` hex color to bare uppercase `RRGGBB`, or null. */
function hexColor(v: string | undefined): string | null {
    if (!v) return null;
    let h = v.trim().replace(/^#/, '');
    if (/^[0-9A-Fa-f]{3}$/.test(h)) h = h.split('').map(c => c + c).join('');
    return /^[0-9A-Fa-f]{6}$/.test(h) ? h.toUpperCase() : null;
}

/** Sanitizes a bookmark/anchor name to Word's `[A-Za-z0-9_]`, leading letter/underscore, <=40 chars. */
function toBookmarkNameRaw(name: string): string {
    let s = (name || '').replace(/[^A-Za-z0-9_]/g, '_');
    if (!/^[A-Za-z_]/.test(s)) s = '_' + s;
    return s.slice(0, 40) || '_';
}

interface Rel { id: string; type: string; target: string; mode?: string; }
interface MediaPart { name: string; bytes: Uint8Array; ext: string; contentType: string; }
interface NoteEntry { key: string; kind: 'footnote' | 'endnote'; node: OfficeContentNode; }
interface CommentEntry { id: number; node: OfficeContentNode; }

/**
 * Generates a Word (.docx / WordprocessingML) document from any AST, as a `Uint8Array`.
 *
 * The package is hand-built with `fflate.zipSync` (the same approach as `EpubGenerator`), with no
 * new dependencies, browser-safe, and reproducible (pinned zip mtimes, deterministic ids). It emits
 * exactly the WML dialect `WordParser` reads back, so a generated document round-trips through the
 * parser with near-zero loss (the basis of its test strategy).
 */
export class DocxGenerator extends BaseGenerator<'docx'> {
    private rels: Rel[] = [];
    private relByTarget = new Map<string, string>();
    private media: MediaPart[] = [];
    private mediaByAttachment = new Map<string, string>();
    private usedExtensions = new Set<string>();
    private bookmarkIds = new Map<string, number>();
    private usedBookmarkNames = new Set<string>();
    private bookmarkCounter = 0;
    private footnotes: NoteEntry[] = [];
    private endnotes: NoteEntry[] = [];
    private comments: CommentEntry[] = [];
    private commentCounter = 0;
    private numbering: { listId: string; numId: number; levels: Map<number, 'ordered' | 'unordered'>; startAt: Map<number, number> }[] = [];
    private numByListId = new Map<string, number>();
    private needsNumbering = false;
    private mathWarned = false;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'docx'>) {
        super('docx', ast, config);
    }

    // ── registries ────────────────────────────────────────────────────────────

    private addRel(type: string, target: string, mode?: string): string {
        const dedupeKey = `${type}::${target}::${mode ?? ''}`;
        const existing = this.relByTarget.get(dedupeKey);
        if (existing) return existing;
        const id = `rId${this.rels.length + 100}`; // offset so document-level ids never clash with .rels
        this.rels.push({ id, type, target, mode });
        this.relByTarget.set(dedupeKey, id);
        return id;
    }

    /** Returns a stable, unique bookmark id for a sanitized name (same name -> same id). */
    private bookmarkId(rawName: string): { id: number; name: string } {
        const base = toBookmarkNameRaw(rawName);
        const existing = this.bookmarkIds.get(base);
        if (existing !== undefined) return { id: existing, name: base };
        let name = base, i = 2;
        while (this.usedBookmarkNames.has(name)) name = `${base.slice(0, 37)}_${i++}`;
        this.usedBookmarkNames.add(name);
        const id = this.bookmarkCounter++;
        this.bookmarkIds.set(name, id);
        return { id, name };
    }

    /** Packages an image attachment once, returning the relationship id or null (skipped). */
    private mediaRel(attachmentName: string): { rid: string; cx: number; cy: number; intrinsic: { w: number; h: number } | null } | null {
        const att = (this.ast.attachments || []).find(a => a.name === attachmentName);
        if (!att || !att.data) return null;
        const ext = MIME_EXT[(att.mimeType || '').toLowerCase()];
        if (!ext) { this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName, reason: 'unsupported mime' }); return null; }
        let bytes: Uint8Array;
        try { bytes = decodeBase64(att.data); } catch { this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName }); return null; }
        let name = this.mediaByAttachment.get(attachmentName);
        if (!name) {
            name = `image${this.media.length + 1}.${ext}`;
            const contentType = (att.mimeType || 'image/png');
            this.media.push({ name, bytes, ext, contentType });
            this.mediaByAttachment.set(attachmentName, name);
            this.usedExtensions.add(ext);
        }
        const rid = this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', `media/${name}`);
        return { rid, cx: 0, cy: 0, intrinsic: sniffImageSize(bytes) };
    }

    // ── metadata / reproducibility ─────────────────────────────────────────────

    private resolveModified(): { iso: string; mtime: Date } {
        const raw: unknown = this.effectiveMetadata.modified;
        let resolved: Date | null = null;
        if (raw instanceof Date && !isNaN(raw.getTime())) resolved = raw;
        else if (typeof raw === 'string' && raw !== '') { const p = new Date(raw); if (!isNaN(p.getTime())) resolved = p; }
        resolved ??= new Date();
        const MIN = Date.UTC(1980, 0, 1), MAX = Date.UTC(2099, 11, 31, 23, 59, 59);
        const t = resolved.getTime();
        const clamped = t < MIN ? new Date(MIN) : t > MAX ? new Date(MAX) : resolved;
        return { iso: resolved.toISOString().replace(/\.\d+Z$/, 'Z'), mtime: clamped };
    }

    // ── entry point ─────────────────────────────────────────────────────────────

    async generate(): Promise<ConversionResult<'docx'>> {
        this.warnUnrepresentableCustomMetadata('docx');
        const { iso, mtime } = this.resolveModified();

        // Pre-pass: assign numbering ids in first-seen listId order (deterministic).
        this.prescanNumbering(this.ast.content);

        // Render body blocks (this also populates rels, media, notes, comments, bookmarks).
        const metaBlock = this.config.renderMetadata ? this.renderMetadataBlock() : '';
        const body = await this.renderBlocks(this.ast.content);

        const hasHeader = !!(this.ast.auxiliary?.headers?.length);
        const hasFooter = !!(this.ast.auxiliary?.footers?.length);
        const headerXml = hasHeader ? await this.renderHdrFtr('hdr', this.ast.auxiliary!.headers!) : '';
        const footerXml = hasFooter ? await this.renderHdrFtr('ftr', this.ast.auxiliary!.footers!) : '';
        const headerRid = hasHeader ? this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', 'header1.xml') : '';
        const footerRid = hasFooter ? this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', 'footer1.xml') : '';

        const documentXml = this.buildDocumentXml(metaBlock + body, headerRid, footerRid);
        const parts = this.assembleParts({ documentXml, iso, hasHeader, hasFooter, headerXml, footerXml });

        const zipFiles: Zippable = {};
        for (const [path, content] of Object.entries(parts)) zipFiles[path] = new TextEncoder().encode(content);
        for (const m of this.media) zipFiles[`word/media/${m.name}`] = m.bytes;

        const bytes = zipSync(zipFiles, { mtime });
        return { value: bytes, messages: this.messages };
    }

    // ── block / inline traversal ─────────────────────────────────────────────────

    private async renderBlocks(nodes: OfficeContentNode[] | undefined): Promise<string> {
        let out = '';
        let prevPaginated: string | null = null;
        for (const node of nodes || []) {
            checkAbortSignal(this.config.abortSignal);
            const override = await this.handleOnNode(node);
            if (override === false) continue;
            if (typeof override === 'string') { out += override; prevPaginated = null; continue; }
            // Preserve source pagination: a page break between consecutive page (or slide) nodes.
            if ((node.type === 'page' || node.type === 'slide') && prevPaginated === node.type) {
                out += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
            }
            out += await this.renderBlockNode(node);
            prevPaginated = (node.type === 'page' || node.type === 'slide') ? node.type : null;
        }
        return out;
    }

    private async renderBlockNode(node: OfficeContentNode): Promise<string> {
        switch (node.type) {
            case 'paragraph': return this.paragraph(node);
            case 'heading': return this.paragraph(node, this.headingStyle(node));
            case 'list': return this.listItem(node);
            case 'table': return this.table(node);
            case 'image': return this.wrapP(await this.imageRun(node), this.jc((node.metadata as any)?.align));
            case 'code': return this.codeBlock(node);
            case 'break': return this.blockBreak(node);
            case 'note': return ''; // notes ride node.notes, never appear as standalone blocks
            case 'comment': return this.standaloneComment(node);
            case 'admonition': return this.admonition(node);
            case 'chart': return this.chart(node);
            case 'embed': return this.embed(node);
            case 'definitionTerm': return this.paragraph(node, 'DefinitionTerm');
            case 'definitionDescription': return this.paragraphIndented(node, 720);
            case 'header': case 'footer': return this.renderBlocks(node.children); // inline page furniture
            case 'sheet': return this.sheet(node);
            case 'row': case 'cell': case 'text': // stray inline/table-part at block level: wrap
                return this.wrapP(await this.renderInline([node]));
            case 'page': case 'slide': case 'drawing': case 'definitionList':
                return this.renderBlocks(node.children); // containers: flatten (page/slide breaks handled by caller sequence)
            case 'slideMaster': return ''; // template furniture, dropped
            default: {
                // Exhaustiveness guard: a new OfficeContentNodeType must be classified here.
                const _exhaustive: never = node;
                void _exhaustive;
                return '';
            }
        }
    }

    // ── paragraph family ─────────────────────────────────────────────────────────

    private async paragraph(node: OfficeContentNode, styleOverride?: string): Promise<string> {
        const meta = node.metadata as any;
        const semantic = this.getSemanticMapping(node);
        const style = styleOverride ?? this.styleForTag(semantic?.tag) ?? this.knownStyle(meta?.style);
        const pPr = this.buildPPr({ style, meta });
        const bookmarks = this.bookmarksFor(node);
        const inner = await this.renderInline(node.children && node.children.length ? node.children : [{ type: 'text', text: node.text || '' } as OfficeContentNode]);
        return `<w:p>${pPr}${bookmarks.start}${inner}${bookmarks.end}</w:p>`;
    }

    private async paragraphIndented(node: OfficeContentNode, leftTwips: number): Promise<string> {
        const inner = await this.renderInline(node.children || []);
        return `<w:p><w:pPr><w:ind w:left="${leftTwips}"/></w:pPr>${inner}</w:p>`;
    }

    private buildPPr(opts: { style?: string; meta?: any; numPr?: string; extra?: string }): string {
        const { style, meta, numPr, extra } = opts;
        if (this.config.includeFormatting === false && !style && !numPr && !extra) return '';
        let s = '';
        if (style) s += `<w:pStyle w:val="${escapeXml(style)}"/>`;
        if (numPr) s += numPr;
        if (extra) s += extra;
        if (this.config.includeFormatting !== false && meta) {
            const ind = meta.paragraphIndentation;
            if (ind && (ind.left || ind.right || ind.firstLine || ind.hanging)) {
                const attrs = [
                    ind.left ? `w:left="${ind.left | 0}"` : '',
                    ind.right ? `w:right="${ind.right | 0}"` : '',
                    ind.firstLine ? `w:firstLine="${ind.firstLine | 0}"` : '',
                    ind.hanging ? `w:hanging="${ind.hanging | 0}"` : '',
                ].filter(Boolean).join(' ');
                if (attrs) s += `<w:ind ${attrs}/>`;
            }
            const j = this.jcVal(meta.alignment);
            if (j) s += `<w:jc w:val="${j}"/>`;
        }
        return s ? `<w:pPr>${s}</w:pPr>` : '';
    }

    private bookmarksFor(node: OfficeContentNode): { start: string; end: string } {
        if (this.config.ignoreInternalLinks) return { start: '', end: '' };
        const names: string[] = [...(((node.metadata as any)?.anchorIds) || [])];
        if (this.config.generateIds && node.type === 'heading' && node.text) {
            const slug = this.slugify(node.text);
            if (slug && !names.includes(slug)) names.push(slug);
        }
        let start = '', end = '';
        for (const raw of names) {
            const { id, name } = this.bookmarkId(raw);
            start += `<w:bookmarkStart w:id="${id}" w:name="${escapeXml(name)}"/>`;
            end += `<w:bookmarkEnd w:id="${id}"/>`;
        }
        return { start, end };
    }

    private headingStyle(node: OfficeContentNode): string {
        const level = Math.min(6, Math.max(1, (node.metadata as any)?.level || 1));
        return `Heading${level}`;
    }

    private jcVal(alignment: string | undefined): string | null {
        if (alignment === 'center') return 'center';
        if (alignment === 'right') return 'right';
        if (alignment === 'justify') return 'both';
        return null;
    }
    private jc(alignment: string | undefined): string {
        const v = this.jcVal(alignment);
        return v ? `<w:pPr><w:jc w:val="${v}"/></w:pPr>` : '';
    }

    private wrapP(inner: string, pPr = ''): string {
        return `<w:p>${pPr}${inner}</w:p>`;
    }

    // ── inline runs ──────────────────────────────────────────────────────────────

    private async renderInline(nodes: OfficeContentNode[]): Promise<string> {
        let out = '';
        // Group consecutive children that share the same link into one hyperlink wrapper.
        let i = 0;
        const items = nodes || [];
        while (i < items.length) {
            const node = items[i];
            const override = await this.handleOnNode(node);
            if (override === false) { i++; continue; }
            if (typeof override === 'string') { out += override; i++; continue; }
            const link = (node.type === 'text') ? (node.metadata as any)?.link : undefined;
            if (link && node.type === 'text') {
                let j = i;
                const group: OfficeContentNode[] = [];
                while (j < items.length && items[j].type === 'text' && (items[j].metadata as any)?.link === link) {
                    group.push(items[j]); j++;
                }
                out += this.hyperlink(link, (node.metadata as any)?.linkType, group);
                i = j;
                continue;
            }
            out += await this.inlineNode(node);
            i++;
        }
        return out;
    }

    private async inlineNode(node: OfficeContentNode): Promise<string> {
        switch (node.type) {
            case 'text': return this.textRuns(node) + await this.noteRefs(node) + await this.commentRefs(node);
            case 'code': return this.inlineCode(node);
            case 'image': return await this.imageRun(node);
            case 'break': return this.inlineBreak(node);
            case 'paragraph': case 'heading': // block content nested inline (HTML): render its runs
                return this.renderInline(node.children || [{ type: 'text', text: node.text || '' } as OfficeContentNode]);
            default:
                // Anything else nested inline: fall back to its plain text as a run.
                return this.run(this.getNodeText(node), undefined);
        }
    }

    private run(text: string, fmt: TextFormatting | undefined): string {
        if (!text) return '';
        const rPr = this.buildRPr(fmt);
        // Split on newlines into <w:br/>-separated segments.
        const segs = text.split(/\r\n|\r|\n/);
        let t = '';
        segs.forEach((seg, idx) => {
            if (idx > 0) t += '<w:br/>';
            if (seg) t += `<w:t xml:space="preserve">${xmlText(seg)}</w:t>`;
        });
        return `<w:r>${rPr}${t}</w:r>`;
    }

    private textRuns(node: OfficeContentNode): string {
        return this.run(node.text || '', node.formatting);
    }

    private buildRPr(fmt: TextFormatting | undefined, styleId?: string): string {
        let s = '';
        if (styleId) s += `<w:rStyle w:val="${styleId}"/>`;
        if (fmt && this.config.includeFormatting !== false) {
            if (fmt.font) s += `<w:rFonts w:ascii="${escapeXml(fmt.font)}" w:hAnsi="${escapeXml(fmt.font)}"/>`;
            if (fmt.bold) s += '<w:b/>';
            if (fmt.italic) s += '<w:i/>';
            if (fmt.strikethrough) s += '<w:strike/>';
            const c = hexColor(fmt.color);
            if (c) s += `<w:color w:val="${c}"/>`;
            const sz = lengthToPt(fmt.size);
            if (sz && sz > 0) s += `<w:sz w:val="${Math.round(sz * 2)}"/>`;
            if (fmt.underline) s += '<w:u w:val="single"/>';
            const bg = hexColor(fmt.backgroundColor);
            if (bg) s += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
            if (fmt.superscript) s += '<w:vertAlign w:val="superscript"/>';
            else if (fmt.subscript) s += '<w:vertAlign w:val="subscript"/>';
        }
        return s ? `<w:rPr>${s}</w:rPr>` : '';
    }

    private hyperlink(link: string, linkType: string | undefined, group: OfficeContentNode[]): string {
        const runs = group.map(n => this.styledRun(n.text || '', n.formatting, 'Hyperlink')).join('');
        const internal = linkType === 'internal' || link.startsWith('#');
        if (internal) {
            if (this.config.ignoreInternalLinks) return group.map(n => this.textRuns(n)).join('');
            const { name } = this.bookmarkId(link.replace(/^#/, ''));
            return `<w:hyperlink w:anchor="${escapeXml(name)}">${runs}</w:hyperlink>`;
        }
        const safe = sanitizeDocxUrl(link);
        if (!safe) return group.map(n => this.textRuns(n)).join('');
        const rid = this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', escapeXml(safe), 'External');
        return `<w:hyperlink r:id="${rid}">${runs}</w:hyperlink>`;
    }

    private styledRun(text: string, fmt: TextFormatting | undefined, styleId: string): string {
        if (!text) return '';
        const rPr = this.buildRPr(fmt, styleId);
        return `<w:r>${rPr}<w:t xml:space="preserve">${xmlText(text)}</w:t></w:r>`;
    }

    // ── notes & comments ─────────────────────────────────────────────────────────

    private async noteRefs(node: OfficeContentNode): Promise<string> {
        if (!node.notes || !node.notes.length) return '';
        let out = '';
        for (const note of node.notes) {
            const kind: 'footnote' | 'endnote' = (note.metadata as any)?.noteType === 'endnote' ? 'endnote' : 'footnote';
            const key = this.getFootnoteKey(note);
            if (!this.noteBodies.has(key)) {
                (kind === 'endnote' ? this.endnotes : this.footnotes).push({ key, kind, node });
                // Style the note body paragraphs as FootnoteText; a body always ends with a w:p.
                const body = (await this.renderBlocks(note.children)) || '<w:p/>';
                this.noteBodies.set(key, this.styleNoteBody(body));
            }
            const id = this.noteId(key);
            const tag = kind === 'endnote' ? 'w:endnoteReference' : 'w:footnoteReference';
            out += `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><${tag} w:id="${id}"/></w:r>`;
        }
        return out;
    }

    private styleNoteBody(body: string): string {
        // Give each top-level body paragraph the FootnoteText style (only when it has no pPr already).
        return body.replace(/<w:p>(?!<w:pPr>)/g, '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>');
    }

    private noteIdMap = new Map<string, number>();
    private noteIdCounter = 2; // ids -1/0/1 reserved by the separators
    private noteId(key: string): number {
        let id = this.noteIdMap.get(key);
        if (id === undefined) { id = this.noteIdCounter++; this.noteIdMap.set(key, id); }
        return id;
    }

    private async commentRefs(node: OfficeContentNode): Promise<string> {
        if (!node.comments || !node.comments.length) return '';
        let out = '';
        for (const c of node.comments) {
            const id = this.commentCounter++;
            this.comments.push({ id, node: c });
            this.commentBodies.set(id, (await this.renderBlocks(c.children)) || '<w:p/>');
            out += `<w:r><w:commentReference w:id="${id}"/></w:r>`;
        }
        return out;
    }

    private async standaloneComment(node: OfficeContentNode): Promise<string> {
        const id = this.commentCounter++;
        this.comments.push({ id, node });
        this.commentBodies.set(id, (await this.renderBlocks(node.children)) || '<w:p/>');
        return `<w:p><w:r><w:commentReference w:id="${id}"/></w:r></w:p>`;
    }

    // ── lists / numbering ────────────────────────────────────────────────────────

    private prescanNumbering(nodes: OfficeContentNode[] | undefined): void {
        for (const node of nodes || []) {
            if (node.type === 'list') {
                const meta = node.metadata as any;
                const listId = String(meta?.listId ?? 'pdf-list');
                let entry = this.numByListId.get(listId) !== undefined
                    ? this.numbering.find(n => n.listId === listId)!
                    : undefined;
                if (!entry) {
                    const numId = this.numbering.length + 1;
                    entry = { listId, numId, levels: new Map(), startAt: new Map() };
                    this.numbering.push(entry);
                    this.numByListId.set(listId, numId);
                    this.needsNumbering = true;
                }
                const lvl = Math.max(0, Math.min(8, meta?.indentation | 0));
                if (!entry.levels.has(lvl)) {
                    entry.levels.set(lvl, meta?.listType === 'ordered' ? 'ordered' : 'unordered');
                    if (typeof meta?.itemIndex === 'number') entry.startAt.set(lvl, meta.itemIndex + 1);
                }
            }
            this.prescanNumbering(node.children);
        }
    }

    private async listItem(node: OfficeContentNode): Promise<string> {
        const meta = node.metadata as any;
        const listId = String(meta?.listId ?? 'pdf-list');
        const numId = this.numByListId.get(listId) ?? 1;
        const lvl = Math.max(0, Math.min(8, meta?.indentation | 0));
        const numPr = `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
        const pPr = this.buildPPr({ style: 'ListParagraph', numPr, meta });
        let prefix = '';
        if (meta?.isTask) prefix = this.run(meta.checked ? '☑ ' : '☐ ', undefined);
        const inner = await this.renderInline(node.children || [{ type: 'text', text: node.text || '' } as OfficeContentNode]);
        return `<w:p>${pPr}${prefix}${inner}</w:p>`;
    }

    private buildNumberingXml(): string {
        let abstracts = '', nums = '';
        for (const entry of this.numbering) {
            const aId = entry.numId - 1;
            let levels = '';
            for (let l = 0; l <= 8; l++) {
                const type = entry.levels.get(l) ?? entry.levels.get(0) ?? 'unordered';
                const fmt = type === 'ordered' ? 'decimal' : 'bullet';
                const text = type === 'ordered' ? `%${l + 1}.` : '•';
                levels += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>`
                    + `<w:lvlText w:val="${escapeXml(text)}"/><w:lvlJc w:val="left"/>`
                    + `<w:pPr><w:ind w:left="${720 * (l + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
            }
            abstracts += `<w:abstractNum w:abstractNumId="${aId}"><w:multiLevelType w:val="multilevel"/>${levels}</w:abstractNum>`;
            let overrides = '';
            for (const [lvl, start] of entry.startAt) {
                if (start > 1) overrides += `<w:lvlOverride w:ilvl="${lvl}"><w:startOverride w:val="${start}"/></w:lvlOverride>`;
            }
            nums += `<w:num w:numId="${entry.numId}"><w:abstractNumId w:val="${aId}"/>${overrides}</w:num>`;
        }
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abstracts}${nums}</w:numbering>`;
    }

    // ── tables ───────────────────────────────────────────────────────────────────

    private async table(node: OfficeContentNode): Promise<string> {
        const rows = (node.children || []).filter(r => r.type === 'row');
        if (!rows.length) return '';
        // Grid-occupancy pass: compute grid width and per-row rendered cells with synthesized merges.
        const cols = this.gridWidth(rows);
        const contentWidth = this.contentWidthTwips();
        const colW = Math.max(1, Math.floor(contentWidth / Math.max(1, cols)));
        const tblGrid = `<w:tblGrid>${Array.from({ length: cols }, () => `<w:gridCol w:w="${colW}"/>`).join('')}</w:tblGrid>`;
        const tableAlign = this.jcVal((node.metadata as any)?.align);
        const tblPr = `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>`
            + (tableAlign ? `<w:jc w:val="${tableAlign}"/>` : '')
            + `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('')}</w:tblBorders></w:tblPr>`;

        const active = new Map<number, number>(); // grid col -> remaining vMerge rows
        let trs = '';
        for (const row of rows) {
            const cells = (row.children || []).filter(c => c.type === 'cell');
            let tcs = '';
            let col = 0, ci = 0;
            while (ci < cells.length || [...active.entries()].some(([c]) => c >= col)) {
                if ((active.get(col) || 0) > 0) {
                    // continuation cell for an active vertical merge
                    tcs += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/><w:vMerge/></w:tcPr><w:p/></w:tc>`;
                    active.set(col, active.get(col)! - 1);
                    if (active.get(col)! <= 0) active.delete(col);
                    col++;
                    continue;
                }
                if (ci >= cells.length) break;
                const cell = cells[ci++];
                const cmeta = cell.metadata as any;
                const colSpan = Math.max(1, Math.min(cols, cmeta?.colSpan || 1));
                const rowSpan = Math.max(1, Math.min(1000, cmeta?.rowSpan || 1));
                let tcPr = `<w:tcW w:w="0" w:type="auto"/>`;
                if (colSpan > 1) tcPr += `<w:gridSpan w:val="${colSpan}"/>`;
                if (rowSpan > 1) { tcPr += `<w:vMerge w:val="restart"/>`; for (let k = 0; k < colSpan; k++) active.set(col + k, rowSpan - 1); }
                const bg = hexColor(cmeta?.backgroundColor);
                if (bg) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
                let inner = await this.renderBlocks(cell.children);
                // A w:tc must end with a w:p: append one only when empty or ending in a nested table.
                const trimmed = inner.trimEnd();
                if (!trimmed) inner = '<w:p/>';
                else if (trimmed.endsWith('</w:tbl>')) inner = trimmed + '<w:p/>';
                tcs += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${inner}</w:tc>`;
                col += colSpan;
            }
            trs += `<w:tr>${tcs}</w:tr>`;
        }
        return `<w:tbl>${tblPr}${tblGrid}${trs}</w:tbl>`;
    }

    private gridWidth(rows: OfficeContentNode[]): number {
        let max = 1;
        const active = new Map<number, number>();
        for (const row of rows) {
            let col = 0;
            const cells = (row.children || []).filter(c => c.type === 'cell');
            let ci = 0;
            while (ci < cells.length || [...active.values()].some(v => v > 0)) {
                if ((active.get(col) || 0) > 0) { active.set(col, active.get(col)! - 1); if (active.get(col)! <= 0) active.delete(col); col++; continue; }
                if (ci >= cells.length) break;
                const cmeta = cells[ci++].metadata as any;
                const colSpan = Math.max(1, Math.min(1000, cmeta?.colSpan || 1));
                const rowSpan = Math.max(1, Math.min(1000, cmeta?.rowSpan || 1));
                if (rowSpan > 1) for (let k = 0; k < colSpan; k++) active.set(col + k, rowSpan - 1);
                col += colSpan;
            }
            max = Math.max(max, col);
        }
        return Math.min(1000, max);
    }

    private async sheet(node: OfficeContentNode): Promise<string> {
        const name = (node.metadata as any)?.sheetName;
        const heading = name ? `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">${xmlText(name)}</w:t></w:r></w:p>` : '';
        const table = await this.table({ type: 'table', children: (node.children || []).filter(c => c.type === 'row') } as OfficeContentNode);
        return heading + table;
    }

    // ── images ───────────────────────────────────────────────────────────────────

    private async imageRun(node: OfficeContentNode): Promise<string> {
        const mode: ImageMode = this.imageMode();
        if (mode === 'none') return '';
        const meta = node.metadata as any;
        const ocr = node.text || '';
        if (mode === 'ocrtext-only') return ocr ? this.run(ocr, undefined) : '';

        const rel = meta?.attachmentName ? this.mediaRel(meta.attachmentName) : null;
        let drawing = '';
        if (rel) {
            const { cx, cy } = this.imageEmu(node, rel.intrinsic);
            const id = this.media.length; // stable per-doc id
            const alt = xmlText(meta?.altText || '');
            drawing = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`
                + `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="image${id}" descr="${alt}"/>`
                + `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`
                + `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`
                + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`
                + `<pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}"/><pic:cNvPicPr/></pic:nvPicPr>`
                + `<pic:blipFill><a:blip r:embed="${rel.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
                + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
                + `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
        } else if (meta?.url) {
            // Remote-only image: degrade to a link on the alt text (never fetch bytes: SSRF).
            const safe = sanitizeDocxUrl(meta.url);
            if (safe) { const rid = this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', escapeXml(safe), 'External'); drawing = `<w:hyperlink r:id="${rid}">${this.styledRun(meta.altText || safe, undefined, 'Hyperlink')}</w:hyperlink>`; }
        }
        if (mode === 'image+ocrtext' && ocr) return drawing + this.run('\n' + ocr, undefined);
        return drawing;
    }

    private imageEmu(node: OfficeContentNode, intrinsic: { w: number; h: number } | null): { cx: number; cy: number } {
        const contentWidthEmu = Math.round(this.contentWidthTwips() / TWIPS_PER_PT * EMU_PER_PT);
        const meta = node.metadata as any;
        // 1. explicit width
        if (meta?.width) {
            const pct = /^\s*([\d.]+)\s*%\s*$/.exec(meta.width);
            let cx: number | null = null;
            if (pct) cx = Math.round(contentWidthEmu * Math.min(100, parseFloat(pct[1])) / 100);
            else { const pt = lengthToPt(meta.width); if (pt) cx = Math.round(pt * EMU_PER_PT); }
            if (cx) { const ar = intrinsic ? intrinsic.h / intrinsic.w : 0.75; return this.capEmu(cx, Math.round(cx * ar), contentWidthEmu); }
        }
        // 2. PDF-sourced on-page bounds (points)
        if (node.bounds && node.bounds.width > 0 && node.bounds.height > 0) {
            return this.capEmu(Math.round(node.bounds.width * EMU_PER_PT), Math.round(node.bounds.height * EMU_PER_PT), contentWidthEmu);
        }
        // 3. intrinsic pixels at 96dpi
        if (intrinsic) {
            const cx = Math.round(intrinsic.w / 96 * EMU_PER_IN);
            const cy = Math.round(intrinsic.h / 96 * EMU_PER_IN);
            return this.capEmu(cx, cy, contentWidthEmu);
        }
        // 4. fallback 3in x 2.25in
        return { cx: 3 * EMU_PER_IN, cy: Math.round(2.25 * EMU_PER_IN) };
    }
    private capEmu(cx: number, cy: number, maxCx: number): { cx: number; cy: number } {
        if (cx > maxCx && cx > 0) { cy = Math.round(cy * maxCx / cx); cx = maxCx; }
        return { cx: Math.max(1, cx), cy: Math.max(1, cy) };
    }

    private contentWidthTwips(): number {
        const cfg = this.config.docxConfig;
        const size = PAGE_SIZES[cfg.pageSize] || PAGE_SIZES.A4;
        const w = cfg.landscape ? size.h : size.w;
        const left = Math.round((cfg.margin.left ?? 72) * TWIPS_PER_PT);
        const right = Math.round((cfg.margin.right ?? 72) * TWIPS_PER_PT);
        return Math.max(720, w - left - right);
    }

    // ── breaks / code / admonition / chart / embed ────────────────────────────────

    private blockBreak(node: OfficeContentNode): string {
        const t = (node.metadata as any)?.breakType;
        if (t === 'page') return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
        if (t === 'column') return `<w:p><w:r><w:br w:type="column"/></w:r></w:p>`;
        if (t === 'thematic') return `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`;
        if (t === 'lastRenderedPage') return '';
        return `<w:p><w:r><w:br/></w:r></w:p>`;
    }
    private inlineBreak(node: OfficeContentNode): string {
        const t = (node.metadata as any)?.breakType;
        if (t === 'page') return `<w:r><w:br w:type="page"/></w:r>`;
        if (t === 'column') return `<w:r><w:br w:type="column"/></w:r>`;
        if (t === 'carriageReturn') return `<w:r><w:cr/></w:r>`;
        return `<w:r><w:br/></w:r>`;
    }

    private codeBlock(node: OfficeContentNode): string {
        if ((node.metadata as any)?.math && !this.mathWarned) { this.mathWarned = true; this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'docx', feature: 'math' }); }
        const text = node.text || this.getNodeText(node);
        const segs = text.split(/\r\n|\r|\n/);
        const runs = segs.map((seg, i) => `${i > 0 ? '<w:br/>' : ''}${seg ? `<w:t xml:space="preserve">${xmlText(seg)}</w:t>` : ''}`).join('');
        return `<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr><w:r>${runs}</w:r></w:p>`;
    }
    private inlineCode(node: OfficeContentNode): string {
        return `<w:r><w:rPr><w:rStyle w:val="CodeChar"/></w:rPr><w:t xml:space="preserve">${xmlText(node.text || '')}</w:t></w:r>`;
    }

    private async admonition(node: OfficeContentNode): Promise<string> {
        const meta = node.metadata as any;
        const color = ADMONITION_COLOR[meta?.admonitionType] || '0969DA';
        const title = meta?.title || (meta?.admonitionType ? meta.admonitionType[0].toUpperCase() + meta.admonitionType.slice(1) : 'Note');
        const titleP = `<w:p><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="4" w:color="${color}"/></w:pBdr><w:ind w:left="360"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="${color}"/></w:rPr><w:t xml:space="preserve">${xmlText(title)}</w:t></w:r></w:p>`;
        const body = await this.renderBlocks(node.children);
        // Indent the body paragraphs with a left border to visually box the admonition.
        const boxed = body.replace(/<w:p>(?!<w:pPr>)/g, `<w:p><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="4" w:color="${color}"/></w:pBdr><w:ind w:left="360"/></w:pPr>`);
        return titleP + boxed;
    }

    private async chart(node: OfficeContentNode): Promise<string> {
        if (this.config.includeCharts === false) return '';
        const meta = node.metadata as any;
        const att = (this.ast.attachments || []).find(a => a.name === meta?.attachmentName);
        const data = att?.chartData;
        if (!data) return `<w:p><w:r><w:t xml:space="preserve">[Chart: ${xmlText(meta?.attachmentName || '')}]</w:t></w:r></w:p>`;
        const caption = data.title ? `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${xmlText(data.title)}</w:t></w:r></w:p>` : '';
        // Build a table: header = series names, first col = labels.
        const rows: OfficeContentNode[] = [];
        const header: OfficeContentNode = { type: 'row', children: [cellOf(''), ...data.dataSets.map(d => cellOf(d.name || ''))] };
        rows.push(header);
        (data.labels || []).forEach((label, i) => {
            rows.push({ type: 'row', children: [cellOf(label), ...data.dataSets.map(d => cellOf(String(d.values?.[i] ?? '')))] });
        });
        const table = await this.table({ type: 'table', children: rows } as OfficeContentNode);
        return caption + table;
    }

    private embed(node: OfficeContentNode): string {
        const meta = node.metadata as any;
        const url = meta?.url ? sanitizeDocxUrl(meta.url) : '';
        if (!url) {
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'docx', feature: 'embed' });
            const fallback = meta?.label || node.text || this.getNodeText(node);
            return fallback ? `<w:p>${this.styledRun(fallback, undefined, '')}</w:p>` : '';
        }
        const rid = this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', escapeXml(url), 'External');
        return `<w:p><w:hyperlink r:id="${rid}">${this.styledRun(meta?.label || url, undefined, 'Hyperlink')}</w:hyperlink></w:p>`;
    }

    // ── styleMap helpers ─────────────────────────────────────────────────────────

    private styleForTag(tag: string | undefined): string | undefined {
        if (!tag || !isSafeStyleMapTag(tag)) return undefined;
        const map: Record<string, string> = {
            h1: 'Heading1', h2: 'Heading2', h3: 'Heading3', h4: 'Heading4', h5: 'Heading5', h6: 'Heading6',
            blockquote: 'Quote', pre: 'Code',
        };
        return map[tag.toLowerCase()];
    }
    private knownStyle(style: string | undefined): string | undefined {
        const known = new Set(['Quote', 'IntenseQuote', 'Title', 'ListParagraph', 'Code', 'DefinitionTerm']);
        return style && known.has(style) ? style : undefined;
    }

    // ── header / footer / metadata block ─────────────────────────────────────────

    private async renderHdrFtr(kind: 'hdr' | 'ftr', nodes: OfficeContentNode[]): Promise<string> {
        const inner = await this.renderBlocks(nodes) || '<w:p/>';
        const tag = kind === 'hdr' ? 'w:hdr' : 'w:ftr';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<${tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `
            + `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${inner}</${tag}>`;
    }

    private renderMetadataBlock(): string {
        const m = this.effectiveMetadata;
        let out = '';
        if (m.title) out += `<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t xml:space="preserve">${xmlText(m.title)}</w:t></w:r></w:p>`;
        const bits = [m.author, m.modified instanceof Date ? m.modified.toISOString().slice(0, 10) : undefined].filter(Boolean).join(' — ');
        if (bits) out += `<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">${xmlText(bits)}</w:t></w:r></w:p>`;
        return out;
    }

    // ── part assembly ─────────────────────────────────────────────────────────────

    private buildDocumentXml(body: string, headerRid: string, footerRid: string): string {
        const cfg = this.config.docxConfig;
        const size = PAGE_SIZES[cfg.pageSize] || PAGE_SIZES.A4;
        const pw = cfg.landscape ? size.h : size.w;
        const ph = cfg.landscape ? size.w : size.h;
        const orient = cfg.landscape ? ' w:orient="landscape"' : '';
        const mar = cfg.margin;
        const pgMar = `<w:pgMar w:top="${Math.round((mar.top ?? 72) * TWIPS_PER_PT)}" w:right="${Math.round((mar.right ?? 72) * TWIPS_PER_PT)}" `
            + `w:bottom="${Math.round((mar.bottom ?? 72) * TWIPS_PER_PT)}" w:left="${Math.round((mar.left ?? 72) * TWIPS_PER_PT)}" w:header="708" w:footer="708" w:gutter="0"/>`;
        const sectPr = `<w:sectPr>`
            + (headerRid ? `<w:headerReference w:type="default" r:id="${headerRid}"/>` : '')
            + (footerRid ? `<w:footerReference w:type="default" r:id="${footerRid}"/>` : '')
            + `<w:pgSz w:w="${pw}" w:h="${ph}"${orient}/>${pgMar}</w:sectPr>`;
        const bodyContent = body.trim() ? body : '<w:p/>';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `
            + `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" `
            + `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" `
            + `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" `
            + `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`
            + `<w:body>${bodyContent}${sectPr}</w:body></w:document>`;
    }

    private assembleParts(o: { documentXml: string; iso: string; hasHeader: boolean; hasFooter: boolean; headerXml: string; footerXml: string }): Record<string, string> {
        const parts: Record<string, string> = {};
        const hasFootnotes = this.footnotes.length > 0;
        const hasEndnotes = this.endnotes.length > 0;
        const hasComments = this.comments.length > 0;

        // Fixed-part relationships (styles + numbering + notes + comments), added to the doc rels.
        this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'styles.xml');
        if (this.needsNumbering) this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', 'numbering.xml');
        if (hasFootnotes) this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes', 'footnotes.xml');
        if (hasEndnotes) this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes', 'endnotes.xml');
        if (hasComments) this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'comments.xml');

        parts['[Content_Types].xml'] = this.contentTypes({ hasFootnotes, hasEndnotes, hasComments });
        parts['_rels/.rels'] = ROOT_RELS;
        parts['word/document.xml'] = o.documentXml;
        parts['word/_rels/document.xml.rels'] = this.documentRels();
        parts['word/styles.xml'] = STYLES_XML;
        if (this.needsNumbering) parts['word/numbering.xml'] = this.buildNumberingXml();
        if (hasFootnotes) parts['word/footnotes.xml'] = this.buildNotesXml('footnote', this.footnotes);
        if (hasEndnotes) parts['word/endnotes.xml'] = this.buildNotesXml('endnote', this.endnotes);
        if (hasComments) parts['word/comments.xml'] = this.buildCommentsXml();
        if (o.hasHeader) parts['word/header1.xml'] = o.headerXml;
        if (o.hasFooter) parts['word/footer1.xml'] = o.footerXml;
        parts['docProps/core.xml'] = this.coreProps(o.iso);
        parts['docProps/app.xml'] = APP_XML;
        const custom = this.customProps();
        if (custom) parts['docProps/custom.xml'] = custom;
        return parts;
    }

    private contentTypes(o: { hasFootnotes: boolean; hasEndnotes: boolean; hasComments: boolean }): string {
        const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
        let defaults = `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`;
        for (const ext of this.usedExtensions) {
            const ct = ext === 'jpg' ? 'image/jpeg' : ext === 'svg' ? 'image/svg+xml' : ext === 'tiff' ? 'image/tiff' : `image/${ext}`;
            defaults += `<Default Extension="${ext}" ContentType="${ct}"/>`;
        }
        let overrides = `<Override PartName="/word/document.xml" ContentType="${wml}.document.main+xml"/>`
            + `<Override PartName="/word/styles.xml" ContentType="${wml}.styles+xml"/>`;
        if (this.needsNumbering) overrides += `<Override PartName="/word/numbering.xml" ContentType="${wml}.numbering+xml"/>`;
        if (o.hasFootnotes) overrides += `<Override PartName="/word/footnotes.xml" ContentType="${wml}.footnotes+xml"/>`;
        if (o.hasEndnotes) overrides += `<Override PartName="/word/endnotes.xml" ContentType="${wml}.endnotes+xml"/>`;
        if (o.hasComments) overrides += `<Override PartName="/word/comments.xml" ContentType="${wml}.comments+xml"/>`;
        if (this.ast.auxiliary?.headers?.length) overrides += `<Override PartName="/word/header1.xml" ContentType="${wml}.header+xml"/>`;
        if (this.ast.auxiliary?.footers?.length) overrides += `<Override PartName="/word/footer1.xml" ContentType="${wml}.footer+xml"/>`;
        overrides += `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`
            + `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`;
        if (this.customProps()) overrides += `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>`;
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${overrides}</Types>`;
    }

    private documentRels(): string {
        const rels = this.rels.map(r => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"${r.mode ? ` TargetMode="${r.mode}"` : ''}/>`).join('');
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
    }

    private buildNotesXml(kind: 'footnote' | 'endnote', notes: NoteEntry[]): string {
        const tag = kind === 'footnote' ? 'w:footnote' : 'w:endnote';
        const root = kind === 'footnote' ? 'w:footnotes' : 'w:endnotes';
        const sepTag = kind === 'footnote' ? 'w:footnote' : 'w:endnote';
        let entries = `<${sepTag} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></${sepTag}>`
            + `<${sepTag} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></${sepTag}>`;
        // Render each note body synchronously via a cached string (children already rendered in a pre-pass not available;
        // so render here). Since renderBlocks is async, we handle it in the caller; store pre-rendered bodies.
        for (const n of notes) {
            const id = this.noteId(n.key);
            const body = (this.noteBodies.get(n.key)) ?? '<w:p/>';
            entries += `<${tag} w:id="${id}">${body}</${tag}>`;
        }
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${entries}</${root}>`;
    }

    private noteBodies = new Map<string, string>();
    private buildCommentsXml(): string {
        let entries = '';
        for (const c of this.comments) {
            const meta = c.node.metadata as any;
            const body = this.commentBodies.get(c.id) ?? '<w:p/>';
            entries += `<w:comment w:id="${c.id}" w:author="${xmlText(meta?.author || '')}" w:date="${xmlText(meta?.date || '')}" w:initials="${xmlText(meta?.initials || '')}">${body}</w:comment>`;
        }
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${entries}</w:comments>`;
    }
    private commentBodies = new Map<number, string>();

    private coreProps(iso: string): string {
        const m = this.effectiveMetadata;
        const created = m.created instanceof Date ? m.created.toISOString().replace(/\.\d+Z$/, 'Z') : iso;
        const el = (tag: string, val: string | undefined) => val ? `<${tag}>${xmlText(val)}</${tag}>` : '';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" `
            + `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
            + el('dc:title', m.title) + el('dc:creator', m.author) + el('dc:subject', m.subject)
            + el('dc:description', m.description) + el('cp:keywords', m.keywords) + el('cp:lastModifiedBy', m.lastModifiedBy)
            + `<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>`
            + `<dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified></cp:coreProperties>`;
    }

    private customProps(): string | null {
        const props = this.effectiveMetadata.customProperties;
        if (!props || !Object.keys(props).length) return null;
        let entries = '', pid = 2;
        for (const [key, value] of Object.entries(props)) {
            let v: string;
            if (typeof value === 'boolean') v = `<vt:bool>${value ? 'true' : 'false'}</vt:bool>`;
            else if (typeof value === 'number') v = `<vt:r8>${value}</vt:r8>`;
            else if (value instanceof Date) v = `<vt:filetime>${value.toISOString().replace(/\.\d+Z$/, 'Z')}</vt:filetime>`;
            else v = `<vt:lpwstr>${xmlText(String(value))}</vt:lpwstr>`;
            entries += `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid++}" name="${xmlText(key)}">${v}</property>`;
        }
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" `
            + `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${entries}</Properties>`;
    }
}

function cellOf(text: string): OfficeContentNode {
    return { type: 'cell', text, children: [{ type: 'paragraph', children: [{ type: 'text', text } as OfficeContentNode] } as OfficeContentNode], metadata: { row: 0, col: 0 } as any };
}

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>`;

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>officeParser</Application></Properties>`;

const STYLES_XML = buildStyles();
function buildStyles(): string {
    const heading = (n: number, sz: number) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${sz}"/></w:rPr></w:style>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
        + `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
        + `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>`
        + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`
        + heading(1, 32) + heading(2, 28) + heading(3, 26) + heading(4, 24) + heading(5, 22) + heading(6, 22)
        + `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="56"/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="IntenseQuote"><w:name w:val="Intense Quote"/><w:basedOn w:val="Quote"/><w:qFormat/><w:rPr><w:i/><w:b/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:contextualSpacing/></w:pPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr></w:style>`
        + `<w:style w:type="character" w:styleId="CodeChar"><w:name w:val="Code Char"/><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="DefinitionTerm"><w:name w:val="Definition Term"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>`
        + `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>`
        + `<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="Footnote Reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="Footnote Text"/><w:basedOn w:val="Normal"/></w:style>`
        + `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>`
        + `</w:styles>`;
}
