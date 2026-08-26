/**
 * PDF Parser
 *
 * Extracts text, structure, metadata, images, links and attachments from PDF files using PDF.js
 * (pdfjs-dist).
 *
 * **Features**
 * - High-fidelity text assembly: baseline line clustering, gap-based word spacing (no more glued or
 *   broken words), super/subscript detection, hyphenation repair, and paragraph reconstruction.
 * - Reading-order recovery for multi-column and float-beside-text pages via a recursive XY-cut.
 * - Semantic structure from tagged PDFs (headings, tables, lists, footnotes) via the structure tree,
 *   with a geometric fallback for untagged files. (Tagged path: {@link module:parsers/pdf/structTree}.)
 * - Per-node page geometry (`bounds`) and page dimensions, on by default, opt out with
 *   `ignorePositions`.
 * - Password-protected documents via `pdfParserConfig.password`, or the `pdfParserConfig.onPassword`
 *   callback to supply one lazily/interactively.
 * - Comprehensive metadata (including the document outline/bookmarks, page labels, and permissions),
 *   hyperlink extraction, image extraction with optional OCR, and embedded file attachments.
 *
 * **Pipeline**
 * 1. Open the document (one worker, one pass).
 * 2. Extract metadata and embedded attachments.
 * 3. For each selected page, collect normalized text runs, images, annotations and (when tagged) the
 *    structure tree in a single pass.
 * 4. Post-process purely in memory: resolve fonts once per document, build lines, segment blocks,
 *    reconstruct paragraphs/headings (or walk the structure tree), and interleave images.
 *
 * @module PdfParser
 * @see https://mozilla.github.io/pdf.js/ PDF.js documentation
 */

import { zlibSync } from 'fflate';
import { DEFAULT_OFFICE_PARSER_CONFIG } from '../defaults.js';
import { FullOfficeParserConfig, ImageMetadata, OfficeAttachment, OfficeContentNode, OfficeErrorType, OfficeMetadata, OfficeParserAST, OfficeWarningType, TextFormatting, TextMetadata } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { parseOfficeDate } from '../utils/dateUtils.js';
import { assertNode, isBrowser } from '../utils/envUtils.js';
import { checkAbortSignal, getOfficeError, logWarning } from '../utils/errorUtils.js';
import { createAttachment } from '../utils/imageUtils.js';
import { loadPdfJs } from '../utils/moduleLoader.js';
import { performOcr } from '../utils/ocrUtils.js';
import { computeRunBox, roundBounds, rotateBoundsToRendered, unionAll } from './pdf/geometry.js';
import { PageExtract, PdfImage, PdfLayoutConfig, RawRun, ResolvedFont } from './pdf/pdfTypes.js';
import { blockToNodes, buildLines, computeDocContext, DocContext, PageContext, runsToParagraph, segmentIntoBlocks } from './pdf/textLayout.js';
import { buildTaggedNodes } from './pdf/structTree.js';

/** Type guard for a pdf.js TextItem (marked-content items lack `str`/`transform`). */
function isTextItem(item: any): item is { str: string; transform: number[]; width: number; height: number; fontName: string; dir?: string; hasEOL?: boolean } {
    return item && typeof item.str === 'string' && Array.isArray(item.transform) && item.transform.length >= 6;
}

/** A link annotation matched to text by geometric overlap. */
interface ResolvedLink {
    /** Rect in layout (rotation-0) viewport space: [minX, minY, maxX, maxY]. */
    rect: [number, number, number, number];
    meta: TextMetadata;
}

/** Precomputed CRC-32 table (polynomial 0xEDB88320) for PNG chunk checksums. */
const CRC32_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c >>> 0;
    }
    return table;
})();

/** CRC-32 over a byte range, as PNG requires over each chunk's type+data. */
function crc32(bytes: Uint8Array): number {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC32_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
}

/**
 * Encodes raw RGBA pixel data into a PNG buffer (8-bit RGB, alpha flattened against white).
 *
 * PNG replaced an earlier uncompressed BMP encoder: a scanned page as BMP is many megabytes, and
 * inlining that as a base64 `data:` URI produced multi-megabyte single lines that broke downstream
 * consumers (e.g. Markdown renderers). Deflate typically shrinks a scanned page by an order of
 * magnitude, and PNG is a real image format that Tesseract and browsers both accept, so OCR and
 * embedding still work. Uses fflate's `zlibSync` (the same browser-safe dependency the ZIP reader
 * uses) rather than Node's `zlib`, so the browser bundle needs no polyfill.
 */
function encodePng(width: number, height: number, data: Uint8Array | Uint8ClampedArray): Buffer {
    // Raw image data: one filter byte (0 = none) per scanline, then RGB triples, top-to-bottom.
    const raw = new Uint8Array(height * (1 + width * 3));
    let o = 0;
    for (let y = 0; y < height; y++) {
        raw[o++] = 0; // filter: none
        for (let x = 0; x < width; x++) {
            const i = (y * width + x) * 4;
            const alpha = data[i + 3] / 255;
            raw[o++] = Math.round(data[i + 0] * alpha + 255 * (1 - alpha));
            raw[o++] = Math.round(data[i + 1] * alpha + 255 * (1 - alpha));
            raw[o++] = Math.round(data[i + 2] * alpha + 255 * (1 - alpha));
        }
    }
    const idatData = zlibSync(raw, { level: 6 });

    // Assemble chunks: each is length(4 BE) + type(4) + data + CRC32(4 BE) over type+data.
    const chunk = (type: string, body: Uint8Array): Buffer => {
        const typeBytes = Buffer.from(type, 'ascii');
        const out = Buffer.alloc(12 + body.length);
        out.writeUInt32BE(body.length, 0);
        typeBytes.copy(out, 4);
        Buffer.from(body).copy(out, 8);
        out.writeUInt32BE(crc32(out.subarray(4, 8 + body.length)), 8 + body.length);
        return out;
    };

    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;   // bit depth
    ihdr[9] = 2;   // color type: 2 = truecolor RGB
    ihdr[10] = 0;  // compression: deflate
    ihdr[11] = 0;  // filter: adaptive
    ihdr[12] = 0;  // interlace: none

    const signature = Buffer.from([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);
    return Buffer.concat([
        signature,
        chunk('IHDR', ihdr),
        chunk('IDAT', idatData),
        chunk('IEND', new Uint8Array(0)),
    ]);
}

/**
 * Converts raw PDF image pixel data (grayscale/RGB/RGBA) to a normalized RGBA buffer. The result is
 * re-encoded to PNG by {@link encodePng}, which is what enables OCR on PDF images.
 */
function convertToRgbaBuffer(data: Uint8Array | Uint8ClampedArray, width: number, height: number, kind?: number): Buffer {
    let rgbaData: Uint8ClampedArray;
    if (kind === 1) {
        rgbaData = new Uint8ClampedArray(width * height * 4);
        for (let i = 0; i < width * height; i++) {
            const gray = data[i];
            rgbaData[i * 4] = gray;
            rgbaData[i * 4 + 1] = gray;
            rgbaData[i * 4 + 2] = gray;
            rgbaData[i * 4 + 3] = 255;
        }
    } else if (kind === 2 || data.length === width * height * 3) {
        rgbaData = new Uint8ClampedArray(width * height * 4);
        for (let i = 0; i < width * height; i++) {
            rgbaData[i * 4] = data[i * 3];
            rgbaData[i * 4 + 1] = data[i * 3 + 1];
            rgbaData[i * 4 + 2] = data[i * 3 + 2];
            rgbaData[i * 4 + 3] = 255;
        }
    } else {
        rgbaData = data instanceof Uint8ClampedArray ? data : new Uint8ClampedArray(data);
    }
    return Buffer.from(rgbaData.buffer, rgbaData.byteOffset, rgbaData.byteLength);
}

/** Reads and coerces the PDF-specific knobs from a resolved parser config. */
function resolvePdfLayoutConfig(config: FullOfficeParserConfig): PdfLayoutConfig {
    const p = config.pdfParserConfig ?? {};
    const num = (v: unknown, fallback: number) => {
        const n = typeof v === 'number' ? v : Number(v);
        return Number.isFinite(n) ? n : fallback;
    };
    return {
        useTags: p.useTags !== false,
        detectColumns: p.detectColumns !== false,
        mergeHyphenatedWords: p.mergeHyphenatedWords !== false,
        lineToleranceFactor: num(p.lineToleranceFactor, 0.35),
        spaceToleranceFactor: num(p.spaceToleranceFactor, 0.25),
        headingDetection: p.headingDetection ?? 'auto',
        disableTextNormalization: !!p.disableTextNormalization,
        includePositions: !config.ignorePositions,
    };
}

/** Parses a 1-based page-range spec like "1-3,7" against a page count. Empty/invalid means all. */
function parsePageRange(spec: string | undefined, numPages: number): number[] {
    const all = () => Array.from({ length: numPages }, (_, i) => i + 1);
    if (!spec || !spec.trim()) return all();
    const pages = new Set<number>();
    for (const part of spec.split(',')) {
        const token = part.trim();
        if (!token) continue;
        const range = token.match(/^(\d+)\s*-\s*(\d+)$/);
        if (range) {
            const lo = parseInt(range[1], 10), hi = parseInt(range[2], 10);
            if (lo < 1 || hi < lo) return all();
            for (let p = lo; p <= hi; p++) if (p <= numPages) pages.add(p);
        } else if (/^\d+$/.test(token)) {
            const p = parseInt(token, 10);
            if (p >= 1 && p <= numPages) pages.add(p);
        } else {
            return all();
        }
    }
    return pages.size ? [...pages].sort((a, b) => a - b) : all();
}

/** Resolves a font id once per document into name/weight/style plus ascent/descent for geometry. */
async function resolveFont(fontKey: string, commonObjs: any, styles: Record<string, any>): Promise<ResolvedFont> {
    const style = styles[fontKey] || {};
    const resolved: ResolvedFont = {
        bold: false,
        italic: false,
        ascent: typeof style.ascent === 'number' ? style.ascent : 0.8,
        descent: typeof style.descent === 'number' ? style.descent : -0.2,
        vertical: !!style.vertical,
    };
    try {
        if (commonObjs?.has?.(fontKey)) {
            const fontData: any = await new Promise((resolve) => commonObjs.get(fontKey, (d: any) => resolve(d)));
            const rawName: string | undefined = typeof fontData?.name === 'string' ? fontData.name : undefined;
            if (rawName) {
                resolved.name = rawName.replace(/^[A-Z]{6}\+/, '');
                const lower = rawName.toLowerCase();
                resolved.bold = lower.includes('bold') || (typeof fontData.black === 'boolean' && fontData.black);
                resolved.italic = lower.includes('italic') || lower.includes('oblique');
            }
        }
    } catch {
        // Font lookup failed; ascent/descent from styles are enough to proceed.
    }
    if (typeof style.fontFamily === 'string' && !resolved.name) resolved.name = style.fontFamily;
    return resolved;
}

/** Applies a 6-element affine matrix to a point. */
function applyMatrix(m: number[], x: number, y: number): [number, number] {
    return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

/** Converts a PDF-space rect [x1,y1,x2,y2] to a normalized layout-viewport rect. */
function toViewportRect(viewport: any, rect: number[]): [number, number, number, number] {
    const m = viewport.transform;
    const [ax, ay] = applyMatrix(m, rect[0], rect[1]);
    const [bx, by] = applyMatrix(m, rect[2], rect[3]);
    return [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)];
}

/** Resolves a page's Link annotations into geometry + hyperlink metadata (honoring config flags). */
async function resolveAnnotations(
    page: any, viewport: any, pdfDocument: any, config: FullOfficeParserConfig, destCache: Map<string, string | null>,
): Promise<ResolvedLink[]> {
    const out: ResolvedLink[] = [];
    let annots: any[];
    try {
        annots = await page.getAnnotations();
    } catch (e) {
        logWarning(OfficeWarningType.ANNOTATION_EXTRACTION_FAILED, config, page.pageNumber, e);
        return out;
    }
    for (const annot of annots) {
        if (annot.subtype !== 'Link' || !annot.rect) continue;
        const url: string | undefined = annot.url || annot.unsafeUrl || annot.data?.url;
        let meta: TextMetadata | undefined;
        if (url) {
            const internal = url.startsWith('#');
            if (internal && config.ignoreInternalLinks) continue;
            meta = { link: url, linkType: internal ? 'internal' : 'external' };
        } else if (annot.dest) {
            if (config.ignoreInternalLinks) continue;
            const href = await resolveDest(annot.dest, pdfDocument, destCache);
            meta = { link: href, linkType: 'internal' };
        }
        if (meta) out.push({ rect: toViewportRect(viewport, annot.rect), meta });
    }
    return out;
}

/** Resolves an internal destination to `#page=N` when cheaply possible, else `#internal`. */
async function resolveDest(dest: string | unknown[], pdfDocument: any, cache: Map<string, string | null>): Promise<string> {
    try {
        let explicit = dest;
        if (typeof dest === 'string') {
            if (cache.has(dest)) { const c = cache.get(dest); return c ?? '#internal'; }
            if (cache.size >= 2000) return '#internal';
            explicit = await pdfDocument.getDestination(dest);
            if (!explicit) { cache.set(dest, null); return '#internal'; }
        }
        if (Array.isArray(explicit) && explicit[0]) {
            const idx = await pdfDocument.getPageIndex(explicit[0]);
            const href = `#page=${idx + 1}`;
            if (typeof dest === 'string') cache.set(dest, href);
            return href;
        }
    } catch {
        // fall through
    }
    if (typeof dest === 'string') cache.set(dest, null);
    return '#internal';
}

/** Finds the hyperlink metadata for a run by intersecting its box with the page's link rects. */
function linkForBox(x: number, yTop: number, w: number, h: number, links: ResolvedLink[]): TextMetadata | undefined {
    const minX = x, maxX = x + w, minY = yTop, maxY = yTop + h;
    for (const l of links) {
        const [lx1, ly1, lx2, ly2] = l.rect;
        if (minX < lx2 && maxX > lx1 && minY < ly2 && maxY > ly1) return l.meta;
    }
    return undefined;
}

/** Maps the granted-permission flags from `getPermissions()` to readable action names. */
function permissionNames(pdfjs: any, perms: number[]): string[] {
    const F = pdfjs.PermissionFlag || {};
    const pairs: [number | undefined, string][] = [
        [F.PRINT, 'print'], [F.MODIFY_CONTENTS, 'modify'], [F.COPY, 'copy'],
        [F.MODIFY_ANNOTATIONS, 'annotate'], [F.FILL_INTERACTIVE_FORMS, 'fillForms'],
        [F.COPY_FOR_ACCESSIBILITY, 'copyForAccessibility'], [F.ASSEMBLE, 'assemble'],
        [F.PRINT_HIGH_QUALITY, 'printHighQuality'],
    ];
    const set = new Set(perms);
    return pairs.filter(([flag]) => flag !== undefined && set.has(flag)).map(([, name]) => name);
}

/** Lists optional-content (layer) names and default visibility, defensively across pdf.js shapes. */
function listOptionalContentLayers(oc: any): { name: string; visible: boolean }[] {
    const out: { name: string; visible: boolean }[] = [];
    if (!oc || typeof oc.getGroups !== 'function') return out;
    let groups: any;
    try { groups = oc.getGroups(); } catch { return out; }
    if (!groups) return out;
    const entries: [string, any][] = groups instanceof Map
        ? [...groups.entries()]
        : Object.keys(groups).map(k => [k, groups[k]] as [string, any]);
    for (const [id, g] of entries) {
        const name = (g && (g.name || (g.data && g.data.name))) || id;
        let visible = true;
        try { if (typeof oc.isVisible === 'function') visible = !!oc.isVisible(id); } catch { /* keep default */ }
        out.push({ name: String(name), visible });
    }
    return out;
}

/** Builds the document outline (bookmarks) as a tree of `list` nodes carrying destination links. */
async function buildOutline(
    pdfDocument: any, destCache: Map<string, string | null>, config: FullOfficeParserConfig,
): Promise<OfficeContentNode[] | undefined> {
    let outline: any[] | null;
    try { outline = await pdfDocument.getOutline(); } catch { return undefined; }
    if (!Array.isArray(outline) || !outline.length) return undefined;

    const convert = async (items: any[], depth: number): Promise<OfficeContentNode[]> => {
        const nodes: OfficeContentNode[] = [];
        for (const item of items) {
            const title = typeof item?.title === 'string' ? item.title : '';
            let link: string | undefined;
            let linkType: 'internal' | 'external' | undefined;
            if (item?.url) { link = item.url; linkType = 'external'; }
            else if (item?.dest != null) { link = await resolveDest(item.dest, pdfDocument, destCache); linkType = 'internal'; }
            const nested = Array.isArray(item?.items) && item.items.length ? await convert(item.items, depth + 1) : [];
            const label: OfficeContentNode = { type: 'text', text: title };
            if (link) label.metadata = { link, linkType };
            nodes.push({
                type: 'list',
                text: title,
                children: [label, ...nested],
                metadata: { listType: 'unordered', indentation: depth, alignment: 'left', listId: 'pdf-outline', itemIndex: nodes.length },
            });
        }
        return nodes;
    };

    const nodes = await convert(outline, 0);
    return nodes.length ? nodes : undefined;
}

/** Collects one page's text runs, images and annotations into a PageExtract. */
async function collectPage(
    pdfjs: any, pdfDocument: any, pageNumber: number, config: FullOfficeParserConfig,
    pdfCfg: PdfLayoutConfig, fontCache: Map<string, ResolvedFont>, destCache: Map<string, string | null>,
): Promise<PageExtract> {
    const page = await pdfDocument.getPage(pageNumber);
    const rotation = ((page.rotate % 360) + 360) % 360;
    const layoutViewport = page.getViewport({ scale: 1, rotation: 0 });
    const authoredW = layoutViewport.width;
    const authoredH = layoutViewport.height;
    const width = rotation % 180 === 0 ? authoredW : authoredH;
    const height = rotation % 180 === 0 ? authoredH : authoredW;

    const textContent = await page.getTextContent({ includeMarkedContent: true, disableNormalization: pdfCfg.disableTextNormalization });
    const styles: Record<string, any> = textContent.styles || {};

    // Resolve every font on the page once (document-scoped cache).
    const seen = new Set<string>();
    for (const item of textContent.items) if (isTextItem(item) && item.fontName) seen.add(item.fontName);
    for (const key of seen) if (!fontCache.has(key)) fontCache.set(key, await resolveFont(key, page.commonObjs, styles));

    const links = await resolveAnnotations(page, layoutViewport, pdfDocument, config, destCache);

    // Walk items, tracking the marked-content stack for mcid / Artifact scope.
    const stack: { id: string | null; tag: string | null }[] = [];
    const runs: RawRun[] = [];
    for (const item of textContent.items) {
        if (!isTextItem(item)) {
            const type = (item as any).type as string | undefined;
            if (type === 'beginMarkedContent' || type === 'beginMarkedContentProps') {
                stack.push({ id: (item as any).id ?? null, tag: (item as any).tag ?? null });
            } else if (type === 'endMarkedContent') {
                stack.pop();
            }
            continue;
        }
        if (!item.str) continue;
        const font = fontCache.get(item.fontName) || { bold: false, italic: false, ascent: 0.8, descent: -0.2, vertical: false };
        const m = pdfjs.Util.transform(layoutViewport.transform, item.transform);
        const box = computeRunBox(m, item.width || 0, font.ascent, font.descent);

        const formatting: TextFormatting = {};
        if (font.name) formatting.font = font.name;
        if (font.bold) formatting.bold = true;
        if (font.italic) formatting.italic = true;
        formatting.size = String(Math.round(box.fontSize * 2) / 2);

        let mcid: string | null = null, inArtifact = false;
        for (let i = stack.length - 1; i >= 0; i--) {
            if (stack[i].tag === 'Artifact') inArtifact = true;
            if (mcid === null && stack[i].id !== null) mcid = stack[i].id;
        }

        const dir = (item.dir === 'rtl' || item.dir === 'ttb') ? item.dir : 'ltr';
        const link = linkForBox(box.x, box.yTop, box.width, box.height, links);

        runs.push({
            text: item.str,
            x: box.x, yTop: box.yTop, yBaseline: box.yBaseline, width: box.width, height: box.height,
            fontSize: box.fontSize,
            fontKey: item.fontName,
            dir,
            hasEOL: !!item.hasEOL,
            angle: box.angle === -1 ? 0 : box.angle,
            mcid, inArtifact,
            formatting,
            link,
        });
    }

    const images = (config.extractAttachments || config.ocr)
        ? await collectImages(pdfjs, page, layoutViewport, config, pageNumber)
        : [];

    let structTree: unknown | null = null;
    if (pdfCfg.useTags) {
        try { structTree = await page.getStructTree(); } catch { structTree = null; }
    }

    if (typeof page.cleanup === 'function') { try { page.cleanup(); } catch { /* best effort */ } }

    return { pageNumber, width, height, authoredW, authoredH, rotation, runs, images, structTree };
}

/** Extracts images from a page's operator list, positioned in layout-viewport space. */
async function collectImages(pdfjs: any, page: any, viewport: any, config: FullOfficeParserConfig, pageNumber: number): Promise<PdfImage[]> {
    const images: PdfImage[] = [];
    try {
        const ops = await page.getOperatorList();
        const fnArray = ops.fnArray;
        const argsArray = ops.argsArray;
        for (let j = 0; j < fnArray.length; j++) {
            const fn = fnArray[j];
            if (fn === pdfjs.OPS.dependency) {
                for (const dep of argsArray[j]) {
                    try {
                        if (page.objs.has(dep)) continue;
                        await new Promise<void>((resolve) => {
                            const timeout = setTimeout(resolve, 500);
                            page.objs.get(dep, () => { clearTimeout(timeout); resolve(); });
                        });
                    } catch (e) {
                        logWarning(OfficeWarningType.DEPENDENCY_LOAD_FAILED, config, dep, e);
                    }
                }
            }
            if (fn === pdfjs.OPS.paintImageXObject || fn === pdfjs.OPS.paintXObject) {
                const imgName = argsArray[j][0];
                try {
                    let hasObj = page.objs.has(imgName);
                    let targetObjs = page.objs;
                    if (!hasObj && page.commonObjs.has(imgName)) { hasObj = true; targetObjs = page.commonObjs; }
                    if (!hasObj) continue;
                    const imgObj: any = await new Promise((resolve) => targetObjs.get(imgName, (d: any) => resolve(d)));
                    if (isBrowser && !imgObj.data && imgObj.bitmap) {
                        try {
                            const canvas = document.createElement('canvas');
                            canvas.width = imgObj.width; canvas.height = imgObj.height;
                            const ctx = canvas.getContext('2d');
                            if (ctx) { ctx.drawImage(imgObj.bitmap, 0, 0); imgObj.data = ctx.getImageData(0, 0, imgObj.width, imgObj.height).data; imgObj.kind = 3; }
                        } catch (e) {
                            logWarning(OfficeWarningType.IMAGE_PROCESSING_FAILED, config, undefined, e);
                        }
                    }
                    if (imgObj?.data && imgObj.width > 0 && imgObj.height > 0) {
                        // Nearest preceding CTM gives the image placement (unit square mapped by [a,b,c,d,e,f]).
                        let ctm: number[] | null = null;
                        for (let k = j - 1; k >= 0; k--) {
                            if (fnArray[k] === pdfjs.OPS.transform) { ctm = argsArray[k]; break; }
                        }
                        const bounds = ctm ? imageBounds(viewport, ctm) : { x: 0, y: 0, width: 0, height: 0 };
                        images.push({ name: imgName, bounds, data: imgObj.data, pixelWidth: imgObj.width, pixelHeight: imgObj.height, kind: imgObj.kind });
                    }
                } catch {
                    // Image access failed, continue.
                }
            }
        }
    } catch (e) {
        logWarning(OfficeWarningType.IMAGE_EXTRACTION_FAILED, config, `from page ${pageNumber}`, e);
    }
    return images;
}

/** Maps an image CTM to a layout-viewport box (the image fills the unit square under the CTM). */
function imageBounds(viewport: any, ctm: number[]): { x: number; y: number; width: number; height: number } {
    const [a, b, c, d, e, f] = ctm;
    const xs = [e, e + a, e + c, e + a + c];
    const ys = [f, f + b, f + d, f + b + d];
    const rect = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
    const [vx1, vy1, vx2, vy2] = toViewportRect(viewport, rect);
    return { x: vx1, y: vy1, width: vx2 - vx1, height: vy2 - vy1 };
}

/**
 * Parses a PDF file and extracts content.
 *
 * @param buffer - The PDF file buffer
 * @param config - Parser configuration
 * @returns Promise resolving to the parsed AST
 */
export const parsePdf = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    checkAbortSignal(config.abortSignal);
    const pdfjs = await loadPdfJs();
    const pdfCfg = resolvePdfLayoutConfig(config);

    // --- Worker configuration ---
    const workerSrc = config.pdfWorkerSrc;
    if (isBrowser) {
        pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;
    } else {
        assertNode('pdf-worker-auto-resolution');
        let resolved = false;
        if (workerSrc !== DEFAULT_OFFICE_PARSER_CONFIG.pdfWorkerSrc && workerSrc !== '') {
            pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;
            resolved = true;
        } else if ((globalThis as any).pdfjsWorker) {
            resolved = true;
        } else {
            try {
                // @ts-ignore - 'require' is available in Node.js/CommonJS environment
                const localWorkerPath = require.resolve('pdfjs-dist/legacy/build/pdf.worker.mjs');
                const { pathToFileURL } = await import('url');
                pdfjs.GlobalWorkerOptions.workerSrc = pathToFileURL(localWorkerPath).href;
                resolved = true;
            } catch (e) {
                logWarning(OfficeWarningType.PDF_WORKER_FALLBACK, config, undefined, e);
            }
        }
        if (!resolved) pdfjs.GlobalWorkerOptions.workerSrc = workerSrc;
    }

    const onPassword = config.pdfParserConfig?.onPassword;
    // Cap callback-driven retries so an onPassword that keeps returning a wrong password can't loop.
    const MAX_PASSWORD_ATTEMPTS = 3;
    let password: string | undefined = config.pdfParserConfig?.password || undefined;
    let passwordAttempts = 0;

    // Open the document, retrying with an onPassword-supplied password when the PDF is encrypted.
    while (true) {
        checkAbortSignal(config.abortSignal);
        // A fresh copy per attempt: pdf.js transfers the data buffer to its worker, detaching it, so
        // reusing the same Uint8Array on a password retry would fail with a transfer error.
        const loadingTask = pdfjs.getDocument({
            data: new Uint8Array(buffer),
            verbosity: 0,
            isEvalSupported: false,
            password,
        });

        let pdfDocument;
        try {
            pdfDocument = await loadingTask.promise;
        } catch (e: any) {
            try { await loadingTask.destroy(); } catch { /* best effort cleanup */ }
            if (e?.name === 'PasswordException') {
                const need = pdfjs.PasswordResponses?.NEED_PASSWORD;
                const reason: 'required' | 'incorrect' = e.code === need ? 'required' : 'incorrect';
                if (onPassword && passwordAttempts < MAX_PASSWORD_ATTEMPTS) {
                    passwordAttempts++;
                    const supplied = await onPassword(reason);
                    if (supplied) { password = supplied; continue; }
                }
                throw getOfficeError(reason === 'required' ? OfficeErrorType.PDF_PASSWORD_REQUIRED : OfficeErrorType.PDF_PASSWORD_INCORRECT, config);
            }
            const message = e instanceof Error ? e.message : String(e);
            if (message.includes('workerSrc') || message.includes('No "GlobalWorkerOptions.workerSrc" specified')) {
                throw getOfficeError(OfficeErrorType.PDF_WORKER_MISSING, config);
            }
            throw e;
        }

        try {
            return await buildAst(pdfjs, pdfDocument, config, pdfCfg);
        } finally {
            try { await loadingTask.destroy(); } catch { /* best effort cleanup */ }
        }
    }
};

/** Assembles the AST from an opened document. Separated so the caller can guarantee task cleanup. */
async function buildAst(pdfjs: any, pdfDocument: any, config: FullOfficeParserConfig, pdfCfg: PdfLayoutConfig): Promise<OfficeParserAST> {
    const content: OfficeContentNode[] = [];
    const attachments: OfficeAttachment[] = [];
    const numPages = pdfDocument.numPages;

    // --- Metadata ---
    const meta = await pdfDocument.getMetadata().catch(() => ({ info: {} }));
    const info = (meta.info || {}) as Record<string, unknown>;
    const metadata: OfficeMetadata = {
        pages: numPages,
        title: info?.Title as string | undefined,
        author: info?.Author as string | undefined,
        subject: info?.Subject as string | undefined,
        description: info?.Keywords as string | undefined,
        created: parseOfficeDate(info?.CreationDate as string | undefined),
        modified: parseOfficeDate(info?.ModDate as string | undefined),
    };
    if (typeof info?.Language === 'string') metadata.language = info.Language as string;

    const standardPdfInfoKeys = new Set([
        'Title', 'Author', 'Subject', 'Keywords', 'Creator', 'Producer',
        'CreationDate', 'ModDate', 'Trapped', 'IsAcroFormPresent', 'IsXFAPresent',
        'IsCollectionPresent', 'IsSignaturesPresent', 'PDFFormatVersion'
    ]);
    if (info) {
        metadata.nativeProperties = {};
        for (const [key, val] of Object.entries(info)) {
            if (key === 'Custom' && typeof val === 'object' && !Array.isArray(val) && !(val instanceof Date) && val !== null) {
                for (const [ck, cv] of Object.entries(val)) metadata.nativeProperties[ck] = cv;
            } else {
                metadata.nativeProperties[key] = val;
            }
        }
        const customProperties: Record<string, string | number | boolean | Date> = {};
        for (const key of Object.keys(info)) {
            if (standardPdfInfoKeys.has(key)) continue;
            const val = info[key];
            if (val === null || val === undefined) continue;
            if (key === 'Custom' && typeof val === 'object' && !Array.isArray(val) && !(val instanceof Date)) {
                for (const [ck, cv] of Object.entries(val)) {
                    if (cv === null || cv === undefined) continue;
                    if (typeof cv === 'string' || typeof cv === 'number' || typeof cv === 'boolean' || cv instanceof Date) customProperties[ck] = cv;
                }
                continue;
            }
            if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean' || val instanceof Date) customProperties[key] = val;
        }
        if (Object.keys(customProperties).length > 0) metadata.customProperties = customProperties;
    }
    if (meta.metadata) {
        if (!metadata.nativeProperties) metadata.nativeProperties = {};
        const xmp: any = meta.metadata;
        metadata.nativeProperties['XMP'] = typeof xmp.getAll === 'function' ? xmp.getAll() : xmp;
    }

    // Tagged flag for consumers.
    let markInfo: any = null;
    try { markInfo = await pdfDocument.getMarkInfo(); } catch { markInfo = null; }
    if (!metadata.nativeProperties) metadata.nativeProperties = {};
    metadata.nativeProperties['tagged'] = !!(markInfo && markInfo.Marked);
    if (markInfo) metadata.nativeProperties['markInfo'] = markInfo;

    // Document permissions (which user actions the file allows). `null` means all actions allowed;
    // pdf.js cannot validate signatures, so we only report presence, never validity.
    try {
        const perms: number[] | null = await pdfDocument.getPermissions();
        metadata.nativeProperties['permissions'] = perms ? permissionNames(pdfjs, perms) : 'all';
    } catch { /* not available */ }

    // Optional-content group (layer) names and default visibility, for consumers that care which
    // layers exist. Text inside hidden layers is still extracted (getTextContent ignores visibility).
    try {
        const oc = await pdfDocument.getOptionalContentConfig();
        const layers = listOptionalContentLayers(oc);
        if (layers.length) metadata.nativeProperties['layers'] = layers;
    } catch { /* no optional content */ }

    // Printed page labels (e.g. roman-numeral front matter), distinct from the physical page index.
    let pageLabels: (string | null)[] | null = null;
    try { pageLabels = await pdfDocument.getPageLabels(); } catch { pageLabels = null; }

    // --- Embedded file attachments ---
    try {
        const embeddedFiles = await pdfDocument.getAttachments();
        if (embeddedFiles && config.extractAttachments) {
            for (const name in embeddedFiles) {
                const file = embeddedFiles[name];
                attachments.push(createAttachment(file.filename, Buffer.from(file.content)));
            }
        }
    } catch (e) {
        logWarning(OfficeWarningType.ATTACHMENT_EXTRACTION_FAILED, config, undefined, e);
    }

    // --- Single collection pass ---
    const pageNumbers = parsePageRange(config.pdfParserConfig?.pageRange, numPages);
    const fontCache = new Map<string, ResolvedFont>();
    const destCache = new Map<string, string | null>();
    const extracts: PageExtract[] = [];
    for (const pageNum of pageNumbers) {
        checkAbortSignal(config.abortSignal);
        try {
            extracts.push(await collectPage(pdfjs, pdfDocument, pageNum, config, pdfCfg, fontCache, destCache));
        } catch (e: any) {
            logWarning(OfficeWarningType.PAGE_LOAD_FAILED, config, pageNum, e);
        }
    }

    const allRuns = extracts.flatMap(e => e.runs);
    const docCtx = computeDocContext(allRuns, pdfCfg, config.newlineDelimiter);

    // Tagged-structure trust: the document must declare it is tagged and not flag it as suspect.
    const docTagged = !!(markInfo && markInfo.Marked);
    const docTrusted = docTagged && !markInfo.Suspects;
    let structWarned = false;
    const warnStruct = (reason: string) => {
        if (structWarned) return;
        structWarned = true;
        logWarning(OfficeWarningType.PDF_STRUCT_TREE_UNRELIABLE, config, reason);
    };
    if (pdfCfg.useTags && docTagged && !docTrusted) warnStruct('the document flags its tags as suspect');

    const auxHeaders: OfficeContentNode[] = [];
    const auxFooters: OfficeContentNode[] = [];

    let imageCounter = 0;
    for (const extract of extracts) {
        checkAbortSignal(config.abortSignal);
        const pageCtx: PageContext = { pageNumber: extract.pageNumber, authoredW: extract.authoredW, authoredH: extract.authoredH, rotation: extract.rotation };
        const bodyRuns = extract.runs.filter(r => !r.inArtifact);

        let pageContent: OfficeContentNode[] | null = null;

        // Tagged path: use the structure tree when trusted and it covers most of the page's text.
        if (pdfCfg.useTags && docTrusted && extract.structTree) {
            const runsByMcid = new Map<string, RawRun[]>();
            for (const r of bodyRuns) {
                if (!r.mcid) continue;
                const list = runsByMcid.get(r.mcid);
                if (list) list.push(r); else runsByMcid.set(r.mcid, [r]);
            }
            const { nodes, coveredMcids } = buildTaggedNodes(extract.structTree, runsByMcid, pageCtx, docCtx, { ignoreNotes: config.ignoreNotes });
            const textMcids = new Set(bodyRuns.filter(r => r.text.trim() && r.mcid).map(r => r.mcid as string));
            let coveredText = 0;
            for (const m of textMcids) if (coveredMcids.has(m)) coveredText++;
            const coverage = textMcids.size ? coveredText / textMcids.size : 1;
            if (coverage >= 0.7) {
                pageContent = nodes;
                // Stitch in any runs the tags did not cover, via the geometric path.
                const leftover = bodyRuns.filter(r => r.text.trim() && (!r.mcid || !coveredMcids.has(r.mcid)));
                if (leftover.length) {
                    pageContent.push(...geometricNodes(leftover, pageCtx, docCtx, pdfCfg));
                    warnStruct('some text on a page was outside the tag tree');
                }
            } else {
                warnStruct('the tag tree covered too little of the page text');
            }
        }

        // Geometric fallback (untagged, distrusted, or low coverage).
        if (!pageContent) pageContent = geometricNodes(bodyRuns, pageCtx, docCtx, pdfCfg);

        // Images: emit as attachments/OCR, then splice each into the flow before the first text node
        // that sits lower on the page, so reading order is preserved without reordering text.
        for (const img of extract.images) {
            const node = await emitImage(img, extract.pageNumber, ++imageCounter, config, attachments, pageCtx, pdfCfg);
            if (!node) continue;
            const y = node.bounds?.y ?? Infinity;
            const at = pageContent.findIndex(n => n.type !== 'image' && (n.bounds?.y ?? Infinity) > y);
            if (at < 0) pageContent.push(node); else pageContent.splice(at, 0, node);
        }

        // Artifact runs (running headers/footers, decorations) route to auxiliary unless dropped.
        if (!config.ignoreHeadersAndFooters) {
            const artifacts = extract.runs.filter(r => r.inArtifact && r.text.trim());
            classifyArtifacts(artifacts, extract, pageCtx, docCtx, pdfCfg, auxHeaders, auxFooters);
        }

        const pageNode: OfficeContentNode = {
            type: 'page',
            children: pageContent,
            text: pageContent.map(n => n.text).join(config.newlineDelimiter),
            metadata: { pageNumber: extract.pageNumber },
        };
        if (pdfCfg.includePositions && pageNode.type === 'page' && pageNode.metadata) {
            pageNode.metadata.pageWidth = Math.round(extract.width * 100) / 100;
            pageNode.metadata.pageHeight = Math.round(extract.height * 100) / 100;
            if (extract.rotation) pageNode.metadata.rotation = extract.rotation;
        }
        if (pageNode.type === 'page' && pageNode.metadata) {
            const label = pageLabels?.[extract.pageNumber - 1];
            if (label && label !== String(extract.pageNumber)) pageNode.metadata.pageLabel = label;
        }
        content.push(pageNode);
    }

    // Document outline (bookmarks / TOC), when present, into auxiliary.
    const outline = config.ignoreInternalLinks ? undefined : await buildOutline(pdfDocument, destCache, config);

    const auxiliary = (auxHeaders.length || auxFooters.length || outline)
        ? { headers: auxHeaders, footers: auxFooters, ...(outline ? { outline } : {}) }
        : undefined;

    return createAST('pdf', metadata, content, attachments, config, auxiliary);
}

/** Runs the geometric (untagged) assembly path over a set of runs. */
function geometricNodes(runs: RawRun[], pageCtx: PageContext, docCtx: DocContext, pdfCfg: PdfLayoutConfig): OfficeContentNode[] {
    const lines = buildLines(runs, pdfCfg);
    const out: OfficeContentNode[] = [];
    for (const block of segmentIntoBlocks(lines, pdfCfg)) out.push(...blockToNodes(block, pageCtx, docCtx));
    // Rescue rotated text (90/180/270), which the horizontal line builder skips, so it is not
    // silently dropped. It is appended after the main flow, in the source content order (which is
    // usually the correct reading order); precise visual ordering of rotated text is a limitation.
    out.push(...rotatedTextNodes(runs, pageCtx, pdfCfg));
    return out;
}

/** Recovers rotated (90/180/270) runs as trailing paragraphs, one per angle, in content order. */
function rotatedTextNodes(runs: RawRun[], pageCtx: PageContext, pdfCfg: PdfLayoutConfig): OfficeContentNode[] {
    const out: OfficeContentNode[] = [];
    for (const angle of [90, 270, 180] as const) {
        const group = runs.filter(r => r.angle === angle && r.text.trim().length > 0);
        if (!group.length) continue;
        let text = '';
        for (const r of group) {
            if (text && !/\s$/.test(text) && !/^\s/.test(r.text)) text += ' ';
            text += r.text;
        }
        text = text.replace(/\s+/g, ' ').trim();
        if (!text) continue;
        const node: OfficeContentNode = { type: 'paragraph', text, children: [{ type: 'text', text }] };
        const box = unionAll(group.map(r => ({ x: r.x, y: r.yTop, width: r.width, height: r.height })));
        if (pdfCfg.includePositions && box) {
            const rendered = roundBounds(rotateBoundsToRendered(box, pageCtx.rotation, pageCtx.authoredW, pageCtx.authoredH));
            node.bounds = rendered;
            if (node.children && node.children[0]) node.children[0].bounds = rendered;
        }
        out.push(node);
    }
    return out;
}

/** Sorts artifact text in the top/bottom margins into header/footer paragraph groups. */
function classifyArtifacts(
    runs: RawRun[], extract: PageExtract, pageCtx: PageContext, docCtx: DocContext, pdfCfg: PdfLayoutConfig,
    headers: OfficeContentNode[], footers: OfficeContentNode[],
): void {
    if (!runs.length) return;
    const h = extract.authoredH;
    const headerRuns = runs.filter(r => r.yTop < 0.15 * h);
    const footerRuns = runs.filter(r => r.yTop > 0.85 * h);
    for (const node of geometricNodes(headerRuns, pageCtx, docCtx, pdfCfg)) headers.push(retypeAsHeaderFooter(node, 'header'));
    for (const node of geometricNodes(footerRuns, pageCtx, docCtx, pdfCfg)) footers.push(retypeAsHeaderFooter(node, 'footer'));
}

/** Wraps a paragraph produced from margin artifacts as a header/footer node. */
function retypeAsHeaderFooter(node: OfficeContentNode, type: 'header' | 'footer'): OfficeContentNode {
    return { type, text: node.text, children: node.children, bounds: node.bounds, metadata: { type: 'default' } };
}

/** Encodes/OCRs one image and returns its positioned image node (or null when nothing was emitted). */
async function emitImage(
    img: PdfImage, pageNumber: number, index: number, config: FullOfficeParserConfig,
    attachments: OfficeAttachment[], page: PageContext, pdfCfg: PdfLayoutConfig,
): Promise<OfficeContentNode | null> {
    if (!config.extractAttachments) return null;
    const attachmentName = `pdf_image_p${pageNumber}_${index}.png`;
    try {
        const rgba = convertToRgbaBuffer(img.data, img.pixelWidth, img.pixelHeight, img.kind);
        const png = encodePng(img.pixelWidth, img.pixelHeight, new Uint8Array(rgba));
        const attachment = createAttachment(attachmentName, png);
        attachment.mimeType = 'image/png';
        if (config.ocr && img.pixelWidth >= 10 && img.pixelHeight >= 10) {
            try { attachment.ocrText = (await performOcr(png, { ...config.ocrConfig })).trim(); }
            catch (e) { logWarning(OfficeWarningType.OCR_FAILED, config, attachmentName, e); }
        }
        attachments.push(attachment);
        const metadata: ImageMetadata = { attachmentName };
        const node: OfficeContentNode = { type: 'image', text: attachment.ocrText || '', metadata };
        if (pdfCfg.includePositions) node.bounds = roundBounds(rotateBoundsToRendered(img.bounds, page.rotation, page.authoredW, page.authoredH));
        return node;
    } catch (e) {
        logWarning(OfficeWarningType.IMAGE_EXTRACTION_FAILED, config, attachmentName, e);
        return null;
    }
}
