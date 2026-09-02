/**
 * Shared helpers for the office-package generators (DOCX, ODT, EPUB).
 *
 * These are the pure, format-agnostic pieces every hand-built package writer needs: base64 image
 * decoding, image-size sniffing, length/colour/date normalisation, and the reproducible zip instant.
 * They live here so DocxGenerator, OdtGenerator and EpubGenerator share one implementation instead
 * of drifting copies.
 *
 * @module officeGenUtils
 */

import { OfficeContentNode } from '../types.js';

/**
 * Whether a table row should be treated as a header row, mirroring the HtmlGenerator heuristic so
 * every generator agrees. A field a parser actually sets (never the test-only `isHeader`): the row's
 * or a cell's `style` containing "header" (PDF `TH` cells set `style: 'header'`), an explicit
 * `isHeader` flag, or - for the first row only, where it is a reliable signal - every cell bold.
 */
export function isHeaderRow(row: OfficeContentNode, isFirstRow: boolean): boolean {
    const cells = (row.children || []).filter(c => c.type === 'cell');
    if (!cells.length) return false;
    const meta = row.metadata as any;
    if (meta?.isHeader) return true;
    if (typeof meta?.style === 'string' && meta.style.toLowerCase().includes('header')) return true;
    const cellIsHeader = (c: OfficeContentNode) => {
        const cm = c.metadata as any;
        if (cm?.isHeader) return true;
        return typeof cm?.style === 'string' && cm.style.toLowerCase().includes('header');
    };
    if (cells.every(cellIsHeader)) return true;
    if (isFirstRow && cells.every(c => !!c.children?.length && c.children.every(ch => ch.formatting?.bold === true))) return true;
    return false;
}

/**
 * Percent-encodes URL-unsafe characters (space, `<>"{}|^[]` and the like) for an href/Target sink.
 * `encodeURI` leaves existing `%xx` escapes and structural characters (`/?:@&=#`) intact, so it is
 * idempotent for already-valid URLs; Word rejects some of those raw characters in a relationship
 * Target with an "unreadable content" dialog. Falls back to the raw string if encoding throws.
 */
export function encUrl(url: string): string {
    try { return encodeURI(url); } catch { return url; }
}

/** Decodes a base64 string to raw bytes, cross-env (atob exists in Node 16+ and browsers). */
export function decodeBase64(b64: string): Uint8Array {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
}

/** Image MIME to file extension for a packaged media part. */
export const MIME_EXT: Record<string, string> = {
    'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/png': 'png', 'image/gif': 'gif',
    'image/svg+xml': 'svg', 'image/webp': 'webp', 'image/bmp': 'bmp', 'image/tiff': 'tiff',
};

/** Admonition type to accent color (matches the HTML generator's palette family). */
export const ADMONITION_COLOR: Record<string, string> = {
    note: '0969DA', tip: '1A7F37', important: '8250DF', warning: '9A6700', caution: 'CF222E',
};

/** Parses a CSS-ish length ('12pt', '1in', '2cm', '10mm', '96px', or a bare number = px) to points. */
export function lengthToPt(value: string | number | undefined): number | null {
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
export function sniffImageSize(bytes: Uint8Array): { w: number; h: number } | null {
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

/** Validates a `#RRGGBB`/`RRGGBB`/`#RGB` hex color to bare uppercase `RRGGBB`, or null. */
export function hexColor(v: string | undefined): string | null {
    if (!v) return null;
    let h = v.trim().replace(/^#/, '');
    if (/^[0-9A-Fa-f]{3}$/.test(h)) h = h.split('').map(c => c + c).join('');
    return /^[0-9A-Fa-f]{6}$/.test(h) ? h.toUpperCase() : null;
}

/** Sanitizes a bookmark/anchor name to `[A-Za-z0-9_]`, leading letter/underscore, <=40 chars. */
export function toBookmarkNameRaw(name: string): string {
    let s = (name || '').replace(/[^A-Za-z0-9_]/g, '_');
    if (!/^[A-Za-z_]/.test(s)) s = '_' + s;
    return s.slice(0, 40) || '_';
}

/** Normalizes a Date or date-like string to a W3CDTF instant (`YYYY-MM-DDThh:mm:ssZ`), or null. */
export function toW3CDTF(v: unknown): string | null {
    let d: Date | null = null;
    if (v instanceof Date && !isNaN(v.getTime())) d = v;
    else if (typeof v === 'string' && v.trim() !== '') { const p = new Date(v); if (!isNaN(p.getTime())) d = p; }
    return d ? d.toISOString().replace(/\.\d+Z$/, 'Z') : null;
}

/**
 * Resolves the modification instant used for both a package's visible date property and every zip
 * entry's mtime, from a raw `effectiveMetadata.modified` value.
 *
 * Prefers a real Date, then a date-like string, and only falls back to the current time when neither
 * exists (the sole non-reproducible path, hence last). `iso` is `YYYY-MM-DDThh:mm:ssZ` (UTC, whole
 * seconds). `mtime` is additionally clamped to zip's DOS-timestamp range (1980-2099): fflate throws
 * rather than clamping, and an unset/epoch-zero date is the common out-of-range case.
 */
export function resolveZipInstant(raw: unknown): { iso: string; mtime: Date } {
    let resolved: Date | null = null;
    if (raw instanceof Date && !isNaN(raw.getTime())) resolved = raw;
    else if (typeof raw === 'string' && raw !== '') { const p = new Date(raw); if (!isNaN(p.getTime())) resolved = p; }
    resolved ??= new Date();
    const MIN = Date.UTC(1980, 0, 1), MAX = Date.UTC(2099, 11, 31, 23, 59, 59);
    const t = resolved.getTime();
    const clamped = t < MIN ? new Date(MIN) : t > MAX ? new Date(MAX) : resolved;
    return { iso: resolved.toISOString().replace(/\.\d+Z$/, 'Z'), mtime: clamped };
}
