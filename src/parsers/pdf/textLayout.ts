/**
 * The geometric PDF text-assembly pipeline (the "untagged" path).
 *
 * Stages, in order: build visual lines from normalized runs (baseline clustering + gap-based word
 * spacing + super/subscript attachment), segment a page into reading-order blocks (recursive
 * XY-cut), then reconstruct paragraphs and headings from each block. The line builder is shared with
 * the tagged path (structTree.ts) so formatting, spacing and bounds are identical on both.
 *
 * Everything here is pure: plain data in, AST nodes out, no pdf.js objects.
 *
 * @module parsers/pdf/textLayout
 */

import { ListMetadata, NodeBounds, OfficeContentNode, TextFormatting, TextMetadata } from '../../types.js';
import { rotateBoundsToRendered, roundBounds, unionAll, unionBounds } from './geometry.js';
import { PdfLayoutConfig, PdfLine, RawRun, TextFragment } from './pdfTypes.js';

// ── small numeric helpers ───────────────────────────────────────────────────

function median(values: number[]): number {
    if (!values.length) return 0;
    const s = [...values].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
}

/** Rounds a font size to the nearest 0.5, the granularity used for size comparisons. */
function roundSize(n: number): number {
    return Math.round(n * 2) / 2;
}

/** Char-count-weighted mode of run font sizes (rounded to 0.5). */
function dominantFontSize(runs: RawRun[]): number {
    const weight = new Map<number, number>();
    for (const r of runs) {
        const s = roundSize(r.fontSize);
        weight.set(s, (weight.get(s) || 0) + Math.max(1, r.text.length));
    }
    let best = 0, bestW = -1;
    for (const [s, w] of weight) if (w > bestW) { bestW = w; best = s; }
    return best || 12;
}

// ── formatting/link equality (constrains run merging to same-line, same-style) ──

function sameFormatting(a: TextFormatting, b: TextFormatting): boolean {
    const ka = Object.keys(a), kb = Object.keys(b);
    if (ka.length !== kb.length) return false;
    for (const k of ka) if ((a as any)[k] !== (b as any)[k]) return false;
    return true;
}

function sameLink(a: TextMetadata | undefined, b: TextMetadata | undefined): boolean {
    if (!a && !b) return true;
    if (!a || !b) return false;
    return a.link === b.link && a.linkType === b.linkType;
}

// ── line building ────────────────────────────────────────────────────────────

interface Cluster {
    runs: RawRun[];
    baseline: number;
    weight: number;
    fontSize: number;
    minX: number;
    maxX: number;
}

/**
 * Groups normalized angle-0 runs into visual lines by baseline proximity, attaches stray
 * super/subscript fragments, splits lines at wide horizontal gaps (so columns and table cells become
 * separate atoms), then builds fragments with gap-derived spacing. Returns lines top-to-bottom.
 */
export function buildLines(runs: RawRun[], cfg: PdfLayoutConfig): PdfLine[] {
    const items = runs.filter(r => r.angle === 0 && r.text.length > 0);
    if (!items.length) return [];
    items.sort((a, b) => a.yBaseline - b.yBaseline || a.x - b.x);

    let clusters: Cluster[] = [];
    for (const r of items) {
        let best: Cluster | undefined;
        let bestDiff = Infinity;
        for (const c of clusters) {
            const diff = Math.abs(r.yBaseline - c.baseline);
            const tol = Math.max(1.0, cfg.lineToleranceFactor * Math.min(r.fontSize, c.fontSize));
            if (diff <= tol && diff < bestDiff) { best = c; bestDiff = diff; }
        }
        if (best) addToCluster(best, r);
        else clusters.push({ runs: [r], baseline: r.yBaseline, weight: Math.max(1, r.text.length), fontSize: r.fontSize, minX: r.x, maxX: r.x + r.width });
    }

    clusters = mergeSuperSubClusters(clusters);

    const lines = clusters.flatMap(c => clusterToLines(c.runs, cfg));
    lines.sort((a, b) => a.baseline - b.baseline);
    return lines;
}

function addToCluster(c: Cluster, r: RawRun): void {
    const w = Math.max(1, r.text.length);
    c.baseline = (c.baseline * c.weight + r.yBaseline * w) / (c.weight + w);
    c.weight += w;
    c.runs.push(r);
    if (r.fontSize > c.fontSize) c.fontSize = r.fontSize;
    c.minX = Math.min(c.minX, r.x);
    c.maxX = Math.max(c.maxX, r.x + r.width);
}

/**
 * Folds small clusters that are really super/subscripts (a fraction of the size, baseline offset in
 * the super/subscript band, horizontally inside a larger cluster) back into that larger line. Fixes
 * the case where a raised "script" run would otherwise orphan onto its own line.
 */
function mergeSuperSubClusters(clusters: Cluster[]): Cluster[] {
    const removed = new Set<Cluster>();
    for (const s of clusters) {
        if (removed.has(s)) continue;
        // Candidate small clusters only: short and small.
        const sChars = s.runs.reduce((n, r) => n + r.text.replace(/\s/g, '').length, 0);
        if (sChars > 12) continue;
        let target: Cluster | undefined;
        let bestDiff = Infinity;
        for (const l of clusters) {
            if (l === s || removed.has(l)) continue;
            if (s.fontSize > 0.8 * l.fontSize) continue;
            const diff = Math.abs(s.baseline - l.baseline);
            if (diff > 0.7 * l.fontSize) continue;
            const within = s.minX >= l.minX - 2 && s.maxX <= l.maxX + 2;
            const overlap = s.minX < l.maxX && s.maxX > l.minX;
            if (!within && !overlap) continue;
            if (diff < bestDiff) { bestDiff = diff; target = l; }
        }
        if (target) { for (const r of s.runs) addToCluster(target, r); removed.add(s); }
    }
    return clusters.filter(c => !removed.has(c));
}

/**
 * Splits one baseline cluster into visual sub-lines at wide horizontal gaps, then builds each. A
 * wide gap (much larger than a space) marks a column or table-cell boundary, so the pieces become
 * separate atoms the block segmenter can place independently.
 */
function clusterToLines(clusterRuns: RawRun[], cfg: PdfLayoutConfig): PdfLine[] {
    const runs = [...clusterRuns].sort((a, b) => a.x - b.x);
    const fs = dominantFontSize(runs);
    const wideGap = Math.max(18, 4 * cfg.spaceToleranceFactor * fs, 1.5 * fs);
    const segments: RawRun[][] = [];
    let cur: RawRun[] = [];
    let last: RawRun | undefined;
    for (const r of runs) {
        const isBlank = r.text.trim().length === 0;
        // pdf.js synthesizes a wide whitespace run to bridge a gutter; treat that (or a plain
        // coordinate gap) as a column/cell boundary.
        const coordGap = last ? r.x - (last.x + last.width) : 0;
        if (cur.length && ((isBlank && r.width >= wideGap) || coordGap >= wideGap)) {
            segments.push(cur);
            cur = [];
        }
        if (isBlank && r.width >= wideGap) { last = r; continue; } // drop the separating whitespace
        cur.push(r);
        last = r;
    }
    if (cur.length) segments.push(cur);
    return segments.filter(seg => seg.some(r => r.text.trim().length)).map(seg => clusterToLine(seg, cfg));
}

function clusterToLine(clusterRuns: RawRun[], cfg: PdfLayoutConfig): PdfLine {
    const runs = [...clusterRuns].sort((a, b) => a.x - b.x);
    const lineFontSize = dominantFontSize(runs);

    // baseline as char-weighted mean of the dominant-size runs (ignores raised/lowered runs)
    let bSum = 0, bW = 0;
    for (const r of runs) {
        if (roundSize(r.fontSize) >= lineFontSize) { const w = Math.max(1, r.text.length); bSum += r.yBaseline * w; bW += w; }
    }
    const baseline = bW ? bSum / bW : median(runs.map(r => r.yBaseline));

    let rtlChars = 0, totalChars = 0, boldChars = 0;
    for (const r of runs) {
        const n = r.text.replace(/\s/g, '').length;
        totalChars += n;
        if (r.dir === 'rtl') rtlChars += n;
        if (r.formatting.bold) boldChars += n;
    }
    const dir: 'ltr' | 'rtl' = rtlChars * 2 > totalChars ? 'rtl' : 'ltr';

    const fragments: TextFragment[] = [];
    let prev: RawRun | undefined;
    for (const r of runs) {
        const fmt: TextFormatting = { ...r.formatting };
        const off = baseline - r.yBaseline; // >0 raised, <0 lowered
        const rel = lineFontSize ? Math.abs(off) / lineFontSize : 0;
        // A small-font run offset from the baseline is super/subscript. Real producers lower a
        // subscript far less than they raise a superscript (often ~0.04em vs ~0.3em), so use a much
        // smaller threshold for the downward case.
        if (r.fontSize <= 0.75 * lineFontSize && rel <= 0.7) {
            if (off > 0 && rel >= 0.15) fmt.superscript = true;
            else if (off < 0 && rel >= 0.03) fmt.subscript = true;
        }

        let lead = '';
        if (prev) {
            const gap = r.x - (prev.x + prev.width);
            const ref = Math.min(prev.fontSize, r.fontSize) || lineFontSize;
            // A super/subscript run attaches to the preceding text with no space ("super"+"script" ->
            // "superscript", "H"+"2" -> "H2"); only a genuinely wide gap (a real word boundary) still
            // inserts one. A normal run uses the ordinary small-gap threshold.
            const isSupSub = !!(fmt.superscript || fmt.subscript);
            const threshold = isSupSub ? Math.max(ref, cfg.spaceToleranceFactor * lineFontSize) : cfg.spaceToleranceFactor * ref;
            if (gap >= threshold) lead = ' ';
        }

        const box: NodeBounds = { x: r.x, y: r.yTop, width: r.width, height: r.height };
        const last = fragments[fragments.length - 1];
        if (last && sameFormatting(last.formatting, fmt) && sameLink(last.link, r.link)) {
            last.text += lead + r.text;
            last.bounds = unionBounds(last.bounds, box);
        } else {
            fragments.push({ text: lead + r.text, bounds: box, formatting: fmt, link: r.link });
        }
        prev = r;
    }
    if (fragments.length) fragments[0].text = fragments[0].text.replace(/^\s+/, '');
    // Drop fragments left empty after trimming (e.g. a bullet-glyph run), so generators do not emit
    // stray formatting markers such as an empty `` ` `` for a zero-length monospace run.
    const keptFragments = fragments.filter(f => f.text.length > 0);

    const bounds = unionAll((keptFragments.length ? keptFragments : fragments).map(f => f.bounds))!;
    const visible = keptFragments.map(f => f.text).join('').replace(/\s+$/, '');
    const endsWithHyphen = /[-­]$/.test(visible);

    return {
        fragments: keptFragments, dir, baseline, fontSize: lineFontSize,
        bold: totalChars > 0 && boldChars * 2 > totalChars,
        endsWithHyphen,
        x: bounds.x, yTop: bounds.y, width: bounds.width, height: bounds.height,
    };
}

// ── block segmentation (recursive XY-cut) ────────────────────────────────────

/**
 * Splits a page's lines into reading-order blocks with a recursive XY-cut. Horizontal cuts (row
 * bands) are taken liberally; vertical cuts (column gutters) require a genuine full-height gap so
 * text is never reordered without cause. When `detectColumns` is off, the page is one block.
 */
export function segmentIntoBlocks(lines: PdfLine[], cfg: PdfLayoutConfig): PdfLine[][] {
    if (!cfg.detectColumns || lines.length <= 1) return lines.length ? [lines] : [];
    const medianH = median(lines.map(l => l.height)) || 12;
    const out: PdfLine[][] = [];
    xyCut(lines, medianH, out, 0);
    return out;
}

function xyCut(lines: PdfLine[], medianH: number, out: PdfLine[][], depth: number): void {
    if (lines.length <= 1 || depth > 40) { if (lines.length) out.push(lines); return; }

    const h = bestHorizontalGap(lines);
    const v = bestVerticalGutter(lines);
    const H_MIN = 0.6 * medianH;
    const V_MIN = Math.max(12, 1.5 * medianH);

    const hOk = h && h.gap >= H_MIN;
    const vOk = v && v.gap >= V_MIN;

    // Prefer a horizontal band split; only take a vertical gutter when it is clearly the stronger cut
    // and no adequate horizontal split exists, so single-column text is never carved into columns.
    if (hOk && (!vOk || h!.gap >= v!.gap)) {
        const top = lines.filter(l => l.yTop + l.height <= h!.at + 0.1);
        const bottom = lines.filter(l => l.yTop + l.height > h!.at + 0.1);
        if (top.length && bottom.length) { xyCut(top, medianH, out, depth + 1); xyCut(bottom, medianH, out, depth + 1); return; }
    }
    if (vOk) {
        const left = lines.filter(l => l.x + l.width <= v!.at + 0.1);
        const right = lines.filter(l => l.x + l.width > v!.at + 0.1);
        if (left.length && right.length) {
            const rtl = lines.filter(l => l.dir === 'rtl').length * 2 > lines.length;
            const first = rtl ? right : left;
            const second = rtl ? left : right;
            xyCut(first, medianH, out, depth + 1);
            xyCut(second, medianH, out, depth + 1);
            return;
        }
    }
    out.push(lines);
}

/** Largest empty horizontal band; `at` is the y at which to split (bottom edge of the upper part). */
function bestHorizontalGap(lines: PdfLine[]): { gap: number; at: number } | null {
    const sorted = [...lines].sort((a, b) => a.yTop - b.yTop);
    let maxBottom = sorted[0].yTop + sorted[0].height;
    let best: { gap: number; at: number } | null = null;
    for (let i = 1; i < sorted.length; i++) {
        const gap = sorted[i].yTop - maxBottom;
        if (gap > 0 && (!best || gap > best.gap)) best = { gap, at: maxBottom };
        maxBottom = Math.max(maxBottom, sorted[i].yTop + sorted[i].height);
    }
    return best;
}

/**
 * Widest vertical gutter that fully separates the lines into a left group and a right group.
 * `at` is the x of the split. A gutter only counts if no line straddles it.
 */
function bestVerticalGutter(lines: PdfLine[]): { gap: number; at: number } | null {
    const intervals = lines.map(l => [l.x, l.x + l.width] as const).sort((a, b) => a[0] - b[0]);
    let maxRight = intervals[0][1];
    let best: { gap: number; at: number } | null = null;
    for (let i = 1; i < intervals.length; i++) {
        const gap = intervals[i][0] - maxRight;
        if (gap > 0) {
            // valid only if every line is entirely left of maxRight or entirely right of intervals[i][0]
            const split = (maxRight + intervals[i][0]) / 2;
            const straddles = lines.some(l => l.x < split && l.x + l.width > split);
            if (!straddles && (!best || gap > best.gap)) best = { gap, at: split };
        }
        maxRight = Math.max(maxRight, intervals[i][1]);
    }
    return best;
}

// ── paragraph + heading reconstruction ───────────────────────────────────────

/** Document-wide facts needed to classify headings and render text. */
export interface DocContext {
    cfg: PdfLayoutConfig;
    bodyFontSize: number;
    /** Heading level (1..6) for a rounded font size, or 0 when that size is body text. */
    headingLevelForSize: (roundedSize: number) => number;
    newlineDelimiter: string;
    /** Monotonic counter giving each reconstructed geometric list a unique `listId`. */
    listCounter: { n: number };
}

/** Per-page facts needed to place bounds in rendered space. */
export interface PageContext {
    pageNumber: number;
    authoredW: number;
    authoredH: number;
    rotation: number;
}

function emitBounds(box: NodeBounds, page: PageContext, cfg: PdfLayoutConfig): NodeBounds | undefined {
    if (!cfg.includePositions) return undefined;
    return roundBounds(rotateBoundsToRendered(box, page.rotation, page.authoredW, page.authoredH));
}

/**
 * Computes the document body font size and a size-to-heading-level ranking from all runs. Heading
 * levels are ranks of the distinct "large" sizes (descending), so headings resolve even when their
 * sizes cluster closely.
 */
export function computeDocContext(allRuns: RawRun[], cfg: PdfLayoutConfig, newlineDelimiter: string): DocContext {
    const bodyFontSize = dominantFontSize(allRuns) || 12;
    const headingSizes = new Set<number>();
    if (cfg.headingDetection !== 'off') {
        for (const s of new Set(allRuns.map(r => roundSize(r.fontSize)))) {
            if (s >= 1.2 * bodyFontSize) headingSizes.add(s);
        }
    }
    const ranked = [...headingSizes].sort((a, b) => b - a);
    const levelForSize = new Map<number, number>();
    ranked.forEach((s, i) => levelForSize.set(s, Math.min(6, i + 1)));
    return {
        cfg, bodyFontSize, newlineDelimiter,
        headingLevelForSize: (s: number) => levelForSize.get(s) ?? 0,
        listCounter: { n: 0 },
    };
}

/** Descriptor for one reconstructed paragraph before it becomes an AST node. */
interface ParaGroup {
    lines: PdfLine[];
    fontSize: number;
    bold: boolean;
}

function groupParagraphs(block: PdfLine[]): ParaGroup[] {
    const lines = [...block].sort((a, b) => a.baseline - b.baseline);
    const deltas: number[] = [];
    for (let i = 1; i < lines.length; i++) deltas.push(lines[i].baseline - lines[i - 1].baseline);
    const leading = median(deltas) || (lines[0]?.height ?? 12) * 1.2;
    const blockLeft = Math.min(...lines.map(l => l.x));
    const spaceGuess = 0.25 * (lines[0]?.fontSize ?? 12);

    const groups: ParaGroup[] = [];
    let cur: PdfLine[] = [];
    for (let i = 0; i < lines.length; i++) {
        const l = lines[i];
        if (cur.length) {
            const prev = cur[cur.length - 1];
            const gap = l.baseline - prev.baseline;
            const sizeShift = Math.abs(l.fontSize - prev.fontSize) / Math.max(prev.fontSize, 1) > 0.15;
            const indent = l.x - blockLeft > 4 * spaceGuess && !prev.endsWithHyphen;
            if (gap > 1.45 * leading || sizeShift || indent) { groups.push(finishGroup(cur)); cur = []; }
        }
        cur.push(l);
    }
    if (cur.length) groups.push(finishGroup(cur));
    return groups;
}

function finishGroup(lines: PdfLine[]): ParaGroup {
    const runsLike = lines.flatMap(l => l.fragments.map(f => ({ text: f.text, fontSize: l.fontSize, bold: l.bold })));
    const weight = new Map<number, number>();
    for (const f of runsLike) weight.set(roundSize(f.fontSize), (weight.get(roundSize(f.fontSize)) || 0) + Math.max(1, f.text.length));
    let size = 12, bestW = -1;
    for (const [s, w] of weight) if (w > bestW) { bestW = w; size = s; }
    const boldChars = lines.filter(l => l.bold).reduce((s, l) => s + (l.width), 0);
    const totalChars = lines.reduce((s, l) => s + l.width, 0);
    return { lines, fontSize: size, bold: totalChars > 0 && boldChars * 2 > totalChars };
}

// ── geometric list detection ─────────────────────────────────────────────────

/** A recognized list marker at the start of a paragraph group. */
interface ListMarker {
    type: 'ordered' | 'unordered';
    /** The exact leading token (incl. trailing whitespace) to strip from the item text. */
    raw: string;
    /** The ordinal an ordered marker represents (1-based), or null for unordered/unparseable. */
    number: number | null;
    /** True for a real bullet glyph (•, ◦, ▪ …) rather than the ambiguous "-"/"*" forms. */
    strongBullet: boolean;
}

const BULLET_GLYPHS = '•◦▪‣·○●■□∙-';
const UNORDERED_RE = new RegExp(`^([${BULLET_GLYPHS}]|\\*)[ \\t\\u00a0]+`);
const ORDERED_RE = /^\(?([0-9]{1,3}|[a-zA-Z]|[ivxlcdmIVXLCDM]{1,7})[.)][ \t ]+/;

/** Escapes a literal string for safe insertion into a RegExp. */
function escapeRe(s: string): string { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

function romanValue(s: string): number {
    const map: Record<string, number> = { i: 1, v: 5, x: 10, l: 50, c: 100, d: 500, m: 1000 };
    const t = s.toLowerCase();
    let total = 0, prev = 0;
    for (let i = t.length - 1; i >= 0; i--) {
        const v = map[t[i]];
        if (!v) return 0;
        total += v < prev ? -v : v;
        prev = v;
    }
    return total;
}

/** Detects a list marker at the start of a group's first line, or null. */
function detectListMarker(group: ParaGroup): ListMarker | null {
    const first = (group.lines[0]?.fragments[0]?.text ?? '').replace(/^\s+/, '');
    if (!first) return null;
    const u = UNORDERED_RE.exec(first);
    if (u) return { type: 'unordered', raw: u[0], number: null, strongBullet: u[1] !== '-' && u[1] !== '*' };
    const o = ORDERED_RE.exec(first);
    if (o) {
        const core = o[1];
        let n: number | null = null;
        if (/^[0-9]+$/.test(core)) n = parseInt(core, 10);
        else if (/^[ivxlcdm]+$/i.test(core)) { const r = romanValue(core); n = r > 0 ? r : null; }
        else if (/^[a-z]$/i.test(core)) n = core.toLowerCase().charCodeAt(0) - 96;
        return { type: 'ordered', raw: o[0], number: n, strongBullet: false };
    }
    return null;
}

/** Buckets marker left-edges into 0-based indent levels (each distinct column ~a nesting level). */
function indentLevels(lefts: number[]): (x: number) => number {
    const sorted = [...new Set(lefts.map(x => Math.round(x)))].sort((a, b) => a - b);
    const levels: number[] = [];
    for (const x of sorted) {
        if (!levels.length || x - levels[levels.length - 1] > 12) levels.push(x);
    }
    return (x: number) => {
        let lvl = 0;
        for (let i = 0; i < levels.length; i++) if (x >= levels[i] - 6) lvl = i;
        return lvl;
    };
}

/** Removes a list marker (e.g. "•", "1.", "(a)") plus its trailing space from a node's aggregate
 *  text and its first text run. */
function stripMarkerFrom(node: OfficeContentNode, raw: string): void {
    const re = new RegExp('^\\s*' + escapeRe(raw.trim()) + '[ \\t\\u00a0]*');
    if (node.text) node.text = node.text.replace(re, '');
    const strip = (n: OfficeContentNode): boolean => {
        if (n.type === 'text' && typeof n.text === 'string' && re.test(n.text)) { n.text = n.text.replace(re, ''); return true; }
        for (const c of n.children || []) if (strip(c)) return true;
        return false;
    };
    for (const c of node.children || []) if (strip(c)) break;
}

// ── geometric table detection ─────────────────────────────────────────────────

/** Builds a table cell node from a single visual line (one row's cell). */
function cellFromLine(line: PdfLine, row: number, col: number, page: PageContext, cfg: PdfLayoutConfig): OfficeContentNode {
    const children: OfficeContentNode[] = [];
    for (const f of line.fragments) {
        const child: OfficeContentNode = { type: 'text', text: f.text, formatting: Object.keys(f.formatting).length ? f.formatting : undefined };
        const b = emitBounds(f.bounds, page, cfg); if (b) child.bounds = b;
        if (f.link) child.metadata = f.link;
        children.push(child);
    }
    const cell: OfficeContentNode = { type: 'cell', text: line.fragments.map(f => f.text).join('').trim(), children, metadata: { row, col } };
    const cb = emitBounds({ x: line.x, y: line.yTop, width: line.width, height: line.height }, page, cfg);
    if (cb) cell.bounds = cb;
    return cell;
}

/**
 * Recovers grid tables from a page's lines *before* block/column segmentation, which would otherwise
 * carve a table's columns into separate reading blocks. Lines are grouped into rows by baseline; a
 * run of at least three consecutive rows that each carry two or more cells, whose cell left-edges
 * cluster into three or more stable columns, and whose cells are short (median ≤ 25 characters)
 * becomes a `table`. The short-cell and ≥3-column guards keep a multi-column article of long wrapped
 * lines from being mistaken for a table. Returns each table with the y it should be spliced in at,
 * plus the set of lines it consumed so the caller lays out the remaining lines normally.
 */
export function detectTables(lines: PdfLine[], page: PageContext, doc: DocContext): { tables: { node: OfficeContentNode; y: number }[]; consumed: Set<PdfLine> } {
    const tables: { node: OfficeContentNode; y: number }[] = [];
    const consumed = new Set<PdfLine>();
    // Only short segments can be table cells: a long line is prose, and letting it join a row would
    // bleed neighbouring text (and side-by-side tables) into the grid. So it never becomes a cell.
    const lineLen = (l: PdfLine) => l.fragments.reduce((n, f) => n + f.text.trim().length, 0);
    const items = lines.filter(l => l.fragments.length && lineLen(l) <= 40);
    if (items.length < 6) return { tables, consumed };

    // Group lines into rows by shared baseline; sort each row left-to-right.
    const sorted = [...items].sort((a, b) => a.baseline - b.baseline || a.x - b.x);
    const rows: PdfLine[][] = [];
    for (const l of sorted) {
        const last = rows[rows.length - 1];
        const tol = 0.5 * (last ? Math.min(last[0].fontSize, l.fontSize) : l.fontSize);
        if (last && Math.abs(last[0].baseline - l.baseline) <= tol) last.push(l);
        else rows.push([l]);
    }
    for (const r of rows) r.sort((a, b) => a.x - b.x);

    let i = 0;
    while (i < rows.length) {
        if (rows[i].length < 2) { i++; continue; }
        let j = i;
        while (j < rows.length && rows[j].length >= 2) j++;
        const run = rows.slice(i, j);
        i = j;
        if (run.length < 3) continue;

        const cells = run.flat();
        // Cluster cell left-edges into columns.
        const colStarts: number[] = [];
        for (const x of cells.map(c => c.x).sort((a, b) => a - b)) {
            if (!colStarts.length || x - colStarts[colStarts.length - 1] > 12) colStarts.push(x);
        }
        if (colStarts.length < 3) continue;
        // Short-cell guard: long lines are prose columns, not table cells.
        if (median(cells.map(c => c.fragments.reduce((n, f) => n + f.text.trim().length, 0))) > 25) continue;

        const colOf = (x: number) => {
            let best = 0, bd = Infinity;
            for (let k = 0; k < colStarts.length; k++) { const d = Math.abs(x - colStarts[k]); if (d < bd) { bd = d; best = k; } }
            return best;
        };
        const rowNodes: OfficeContentNode[] = run.map((r, ri) => {
            // Assign columns strictly left-to-right and strictly increasing, so two cells (e.g. a
            // right-aligned number whose left edge lands near a neighbour's column) never collide on
            // one column, which would render as an extra shifted cell.
            let lastCol = -1;
            const cellNodes = r.map(line => {
                const col = Math.max(colOf(line.x), lastCol + 1);
                lastCol = col;
                return cellFromLine(line, ri, col, page, doc.cfg);
            });
            const rowNode: OfficeContentNode = { type: 'row', children: cellNodes, text: cellNodes.map(c => c.text || '').join(' ').trim() };
            const rb = unionAll(cellNodes.map(c => c.bounds)); if (rb) rowNode.bounds = rb;
            return rowNode;
        });
        const table: OfficeContentNode = { type: 'table', children: rowNodes, text: rowNodes.map(r => r.text || '').join('\n') };
        const tb = unionAll(rowNodes.map(r => r.bounds)); if (tb) table.bounds = tb;
        tables.push({ node: table, y: run[0][0].yTop });
        for (const l of cells) consumed.add(l);
    }
    return { tables, consumed };
}

// ── geometric list detection ─────────────────────────────────────────────────

/** A marker that stands alone as its own atom (e.g. "1.1.1.", "i.", "•"), split from its item text. */
const STANDALONE_MARKER_RE = /^(?:[•◦▪‣·○●■□∙]|\(?(?:\d+(?:\.\d+)*|[a-zA-Z]|[ivxlcdmIVXLCDM]{1,7})[.)])$/;

/** Parses a standalone marker token into a {@link ListMarker} (raw empty - nothing to strip). */
function parseStandaloneMarker(t: string): ListMarker | null {
    if (/^[•◦▪‣·○●■□∙]$/.test(t)) return { type: 'unordered', raw: '', number: null, strongBullet: true };
    const m = /^\(?(\d+(?:\.\d+)*|[a-zA-Z]|[ivxlcdmIVXLCDM]{1,7})[.)]$/.exec(t);
    if (!m) return null;
    const core = m[1];
    const last = core.includes('.') ? (core.split('.').filter(Boolean).pop() ?? core) : core;
    let num: number | null = null;
    if (/^\d+$/.test(last)) num = parseInt(last, 10);
    else if (/^[ivxlcdm]+$/i.test(last)) { const r = romanValue(last); num = r > 0 ? r : null; }
    else if (/^[a-z]$/i.test(last)) num = last.toLowerCase().charCodeAt(0) - 96;
    return { type: 'ordered', raw: '', number: num, strongBullet: false };
}

/** All the visible text of a paragraph group, concatenated and trimmed. */
function groupText(g: ParaGroup): string {
    return g.lines.flatMap(l => l.fragments.map(f => f.text)).join('').trim();
}

/**
 * Turns a block's lines into paragraph, heading and list nodes. Line groups are formed as
 * paragraphs (with hyphenation repair), then two list shapes are recovered geometrically:
 *
 * 1. A marker that prefixes its item text ("• item", "1. item").
 * 2. A marker drawn as its own atom in the hanging indent, with the item text a separate group on
 *    the same baseline to its right ("1.1.1." | "Three") - the usual multi-level / roman-numeral case
 *    a PDF producer emits.
 *
 * To keep prose safe, an ambiguous "-"/"*" prefix or a lone numbered prefix only counts as a list
 * when a neighbour carries a compatible marker; an unambiguous bullet glyph, or a separate marker
 * atom sitting beside its text, counts on its own. Nesting level comes from the marker's indent.
 */
export function blockToNodes(block: PdfLine[], page: PageContext, doc: DocContext): OfficeContentNode[] {
    const groups = groupParagraphs(block);
    const N = groups.length;
    const baseOf = (g: ParaGroup) => g.lines[0]?.baseline ?? 0;
    const leftOf = (g: ParaGroup) => Math.min(...g.lines.map(l => l.x));

    const dropped = new Array<boolean>(N).fill(false);
    const attached: (ListMarker | null)[] = new Array(N).fill(null); // separate marker merged onto a group
    const markerLeft: (number | null)[] = new Array(N).fill(null);   // the marker's own x (for nesting)

    // Fold a standalone marker atom into the item-text group sharing its baseline, just to its right.
    for (let i = 0; i < N; i++) {
        if (dropped[i]) continue;
        const t = groupText(groups[i]);
        if (!STANDALONE_MARKER_RE.test(t)) continue;
        const mk = parseStandaloneMarker(t);
        if (!mk) continue;
        const yb = baseOf(groups[i]), xr = leftOf(groups[i]), fs = groups[i].fontSize || 12;
        let best = -1, bestDx = Infinity;
        for (let j = 0; j < N; j++) {
            if (j === i || dropped[j] || attached[j]) continue;
            if (Math.abs(baseOf(groups[j]) - yb) > 0.6 * fs) continue;
            const dx = leftOf(groups[j]) - xr;
            if (dx <= 0) continue; // the text must be to the marker's right
            if (dx < bestDx) { bestDx = dx; best = j; }
        }
        if (best >= 0 && bestDx < 6 * fs) { attached[best] = mk; markerLeft[best] = xr; dropped[i] = true; }
    }

    const prefix: (ListMarker | null)[] = groups.map((g, i) => (dropped[i] || attached[i]) ? null : detectListMarker(g));
    const effective: (ListMarker | null)[] = groups.map((_, i) => attached[i] || prefix[i]);
    // A group starting with a number is not a list item when it is really a numbered heading (its font
    // is heading-sized) or a table-of-contents entry (dot leaders "... 3"). Stripping the number there
    // would corrupt a heading or TOC. Bullet glyphs are exempt (a heading rarely starts with one).
    const headingSized = (g: ParaGroup) => doc.cfg.headingDetection !== 'off' && g.fontSize >= 1.25 * doc.bodyFontSize;
    const isItem = effective.map((m, i) => {
        if (dropped[i] || !m) return false;
        if (m.type === 'ordered' && (headingSized(groups[i]) || /\.{4,}/.test(groupText(groups[i])))) return false;
        if (attached[i] || m.strongBullet) return true;         // separate atom or real bullet: strong
        const prev = effective[i - 1], next = effective[i + 1]; // else require a run of compatible markers
        return (!!prev && prev.type === m.type) || (!!next && next.type === m.type);
    });

    const itemLefts = groups.map((g, i) => isItem[i] ? (markerLeft[i] ?? leftOf(g)) : null).filter((x): x is number => x != null);
    const levelFor = indentLevels(itemLefts);

    const nodes: OfficeContentNode[] = [];
    let listId: string | null = null;
    let itemInList = 0;
    let prevWasItem = false;
    for (let i = 0; i < N; i++) {
        if (dropped[i]) continue;
        if (!isItem[i]) {
            listId = null; prevWasItem = false;
            const node = paragraphNode(groups[i], page, doc);
            if (node && (node.text || '').trim().length) nodes.push(node);
            continue;
        }
        if (!prevWasItem) { listId = `pdf-geo-list-${++doc.listCounter.n}`; itemInList = 0; }
        const marker = effective[i]!;
        const p = paragraphNode(groups[i], page, doc, 0);
        if (!p) { prevWasItem = false; continue; }
        if (marker.raw) stripMarkerFrom(p, marker.raw);
        if (!(p.text || '').trim().length) { prevWasItem = false; continue; }
        const meta: ListMetadata = {
            listType: marker.type,
            indentation: levelFor(markerLeft[i] ?? leftOf(groups[i])),
            alignment: 'left',
            listId: listId!,
            // Number ordered items from the marker the PDF rendered (so an interrupted list resumes);
            // fall back to the running position within this list.
            itemIndex: marker.type === 'ordered' && marker.number != null && marker.number >= 1 ? marker.number - 1 : itemInList,
        };
        itemInList++;
        prevWasItem = true;
        const item: OfficeContentNode = { type: 'list', text: p.text, children: p.children, metadata: meta };
        if (p.bounds) item.bounds = p.bounds;
        nodes.push(item);
    }
    return nodes;
}

/**
 * Builds a single paragraph or heading node from a flat set of runs (all treated as one block).
 * Used by the tagged path, where the structure tree already says these runs form one P/Hn, so the
 * shared line/spacing/bounds logic runs but paragraph splitting does not.
 *
 * @param forcedLevel - heading level to force (0 = paragraph); omit for auto heading detection.
 */
export function runsToParagraph(runs: RawRun[], page: PageContext, doc: DocContext, forcedLevel?: number): OfficeContentNode | null {
    const lines = buildLines(runs, doc.cfg);
    if (!lines.length) return null;
    return paragraphNode(finishGroup(lines), page, doc, forcedLevel);
}

function paragraphNode(group: ParaGroup, page: PageContext, doc: DocContext, forcedLevel?: number): OfficeContentNode | null {
    const cfg = doc.cfg;
    const children: OfficeContentNode[] = [];
    const boxes: NodeBounds[] = [];
    let text = '';

    for (let li = 0; li < group.lines.length; li++) {
        const line = group.lines[li];
        const prev = group.lines[li - 1];
        if (li > 0) {
            // Join with the previous line: de-hyphenate or insert a space.
            const prevText = prev.fragments.map(f => f.text).join('');
            const softHyphen = /­$/.test(prevText);
            const hardHyphen = /-$/.test(prevText);
            const nextStartsLower = /^[\p{Ll}]/u.test(line.fragments[0]?.text ?? '');
            if (cfg.mergeHyphenatedWords && (softHyphen || (hardHyphen && nextStartsLower))) {
                // drop the trailing hyphen from the last emitted child and join with no space
                const lastChild = children[children.length - 1];
                if (lastChild?.text) { lastChild.text = lastChild.text.replace(/[-­]$/, ''); }
                text = text.replace(/[-­]$/, '');
            } else {
                const lastChild = children[children.length - 1];
                if (lastChild?.text && !lastChild.text.endsWith(' ')) lastChild.text += ' ';
                text += ' ';
            }
        }
        for (const frag of line.fragments) {
            const child: OfficeContentNode = {
                type: 'text',
                text: frag.text,
                formatting: Object.keys(frag.formatting).length ? frag.formatting : undefined,
            };
            const b = emitBounds(frag.bounds, page, cfg);
            if (b) child.bounds = b;
            if (frag.link) child.metadata = frag.link;
            children.push(child);
            boxes.push(frag.bounds);
            text += frag.text;
        }
    }

    if (!children.length) return null;
    const level = forcedLevel !== undefined ? forcedLevel : (doc.cfg.headingDetection === 'off' ? 0 : headingLevel(group, doc));
    const container = emitBounds(unionAll(boxes)!, page, cfg);

    if (level > 0) {
        const node: OfficeContentNode = { type: 'heading', text, children, metadata: { level } };
        if (container) node.bounds = container;
        return node;
    }
    const node: OfficeContentNode = { type: 'paragraph', text, children };
    if (container) node.bounds = container;
    return node;
}

function headingLevel(group: ParaGroup, doc: DocContext): number {
    if (doc.cfg.headingDetection === 'off') return 0;
    const size = group.fontSize;
    if (size < 1.2 * doc.bodyFontSize) return 0;
    const chars = group.lines.reduce((s, l) => s + l.fragments.reduce((t, f) => t + f.text.length, 0), 0);
    if (group.lines.length > 3 || chars >= 200) return 0;
    if (!(group.bold || size >= 1.35 * doc.bodyFontSize)) return 0;
    return doc.headingLevelForSize(size) || 1;
}
