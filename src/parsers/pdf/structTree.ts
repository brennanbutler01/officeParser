/**
 * The tagged-PDF path: turns a page's structure tree into semantic AST nodes.
 *
 * pdf.js hands back a per-page tree of `{ role, children }` internal nodes and
 * `{ type: 'content', id }` leaves, where `id` is a marked-content id. Every text run collected in
 * the main pass already carries its `mcid`, so joining a leaf to its runs is a map lookup. Block
 * text (P, Hn, list items, cells) is rebuilt with the same line/spacing/bounds logic as the
 * geometric path via {@link runsToParagraph}, so both paths produce identical run shapes.
 *
 * @module parsers/pdf/structTree
 */

import { CellMetadata, ListMetadata, NoteMetadata, OfficeContentNode } from '../../types.js';
import { unionAll } from './geometry.js';
import { RawRun } from './pdfTypes.js';
import { DocContext, PageContext, runsToParagraph } from './textLayout.js';

interface StructNode {
    role?: string;
    type?: string;
    id?: string;
    alt?: string;
    children?: StructNode[];
}

/** Options that gate which tagged roles are emitted. */
export interface TaggedOptions {
    ignoreNotes: boolean;
}

/** Result of walking one page's structure tree. */
export interface TaggedResult {
    nodes: OfficeContentNode[];
    /** Marked-content ids that were consumed, for the coverage/trust check. */
    coveredMcids: Set<string>;
}

interface WalkCtx {
    runsByMcid: Map<string, RawRun[]>;
    page: PageContext;
    doc: DocContext;
    covered: Set<string>;
    opts: TaggedOptions;
    listCounter: { n: number };
}

const HEADING = /^H([1-6])$/;
const role = (n: StructNode): string => n.role || n.type || '';

/** Walks a page's struct tree into ordered AST nodes. */
export function buildTaggedNodes(
    structTree: any, runsByMcid: Map<string, RawRun[]>, page: PageContext, doc: DocContext, opts: TaggedOptions,
): TaggedResult {
    const covered = new Set<string>();
    const ctx: WalkCtx = { runsByMcid, page, doc, covered, opts, listCounter: { n: 0 } };
    const nodes = structTree ? walkChildren(structTree as StructNode, ctx, 0) : [];
    return { nodes, coveredMcids: covered };
}

function walkChildren(node: StructNode, ctx: WalkCtx, sectionDepth: number): OfficeContentNode[] {
    const out: OfficeContentNode[] = [];
    for (const child of node.children || []) out.push(...walkNode(child, ctx, sectionDepth));
    return out;
}

function walkNode(node: StructNode, ctx: WalkCtx, sectionDepth: number): OfficeContentNode[] {
    if (node.type === 'content') {
        const p = paragraphFrom(node, ctx, 0);
        return p ? [p] : [];
    }
    const r = role(node);
    const hm = HEADING.exec(r);
    if (hm) {
        const level = ctx.doc.cfg.headingDetection === 'off' ? 0 : parseInt(hm[1], 10);
        const p = paragraphFrom(node, ctx, level);
        return p ? [p] : [];
    }

    switch (r) {
        case 'H': return blockWithNotes(node, ctx, Math.min(6, Math.max(1, sectionDepth)));
        case 'P': case 'Caption': case 'Title': case 'Lbl': case 'LBody': return blockWithNotes(node, ctx, 0);
        // A table-of-contents entry (TOCI) wraps a Link + Span + leader dots + page number; emit it as
        // one paragraph instead of letting the default recursion shatter each content leaf into its own.
        case 'TOCI': {
            const nodes = blockWithNotes(node, ctx, 0);
            for (const n of nodes) stripDotLeaders(n);
            return nodes;
        }
        case 'Table': { const t = buildTable(node, ctx); return t ? [t] : []; }
        case 'L': return buildList(node, ctx, 0);
        case 'Note': case 'FENote': {
            if (ctx.opts.ignoreNotes) { markCovered(node, ctx); return []; }
            const n = buildNote(node, ctx);
            return n ? [n] : [];
        }
        case 'Sect': case 'Part': case 'Art': case 'Section':
            return walkChildren(node, ctx, sectionDepth + 1);
        case 'Formula': {
            // Real math reconstruction from positioned glyphs is out of scope; at least mark the block
            // as a formula so consumers can tell it apart from ordinary prose.
            const nodes = blockWithNotes(node, ctx, 0);
            for (const n of nodes) if (n.type === 'paragraph') n.metadata = { ...(n.metadata || {}), style: 'formula' };
            return nodes;
        }
        case 'Private': return [];
        // Transparent containers: recurse and flatten.
        default:
            return walkChildren(node, ctx, sectionDepth);
    }
}

/** Role subtrees whose text must never be folded into a parent paragraph. */
const ALWAYS_SKIP = new Set(['Note', 'FENote']);

/** Gathers runs under a node from content leaves, skipping notes and any extra role subtrees. */
function collectRuns(node: StructNode, ctx: WalkCtx, skip?: Set<string>): RawRun[] {
    const out: RawRun[] = [];
    const recurse = (n: StructNode) => {
        if (n.type === 'content' && n.id) {
            const runs = ctx.runsByMcid.get(n.id);
            if (runs) { out.push(...runs); ctx.covered.add(n.id); }
            return;
        }
        for (const c of n.children || []) {
            if (ALWAYS_SKIP.has(role(c)) || (skip && skip.has(role(c)))) continue;
            recurse(c);
        }
    };
    recurse(node);
    return out;
}

/** Emits a block paragraph/heading followed by any footnotes nested inside it. */
function blockWithNotes(node: StructNode, ctx: WalkCtx, level: number): OfficeContentNode[] {
    const out: OfficeContentNode[] = [];
    const p = paragraphFrom(node, ctx, level);
    if (p) out.push(p);
    for (const noteNode of findNotes(node)) {
        if (ctx.opts.ignoreNotes) { markCovered(noteNode, ctx); continue; }
        const n = buildNote(noteNode, ctx);
        if (n) out.push(n);
    }
    return out;
}

/** Finds Note/FENote subtrees anywhere under a node. */
function findNotes(node: StructNode): StructNode[] {
    const out: StructNode[] = [];
    const recurse = (n: StructNode) => {
        for (const c of n.children || []) {
            if (ALWAYS_SKIP.has(role(c))) out.push(c);
            else recurse(c);
        }
    };
    recurse(node);
    return out;
}

/** Marks every content leaf under a node as covered without emitting it (e.g. ignored notes). */
function markCovered(node: StructNode, ctx: WalkCtx): void {
    if (node.type === 'content' && node.id) ctx.covered.add(node.id);
    for (const c of node.children || []) markCovered(c, ctx);
}

function paragraphFrom(node: StructNode, ctx: WalkCtx, level: number): OfficeContentNode | null {
    const runs = collectRuns(node, ctx);
    if (!runs.length) return null;
    return runsToParagraph(runs, ctx.page, ctx.doc, level);
}

function collectRows(node: StructNode): StructNode[] {
    const out: StructNode[] = [];
    for (const c of node.children || []) {
        const r = role(c);
        if (r === 'TR') out.push(c);
        else if (r === 'THead' || r === 'TBody' || r === 'TFoot') out.push(...collectRows(c));
    }
    return out;
}

/** Median of a numeric list (robust to alignment jitter); 0 for an empty list. */
function medianOf(values: number[]): number {
    if (!values.length) return 0;
    const s = [...values].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
}

/**
 * True when a cell is an empty placeholder: no text and no geometry. A tagged PDF pads the columns
 * (and rows) a merged cell covers with exactly these, keeping the grid rectangular. They are the
 * only cells span inference is ever allowed to absorb.
 */
function isEmptyCell(cell: OfficeContentNode): boolean {
    return !cell.bounds && !(cell.text || '').trim();
}

/**
 * Recovers horizontal cell merges (colSpan) that the tag tree padded with empty placeholder cells.
 * pdf.js does not expose the `/ColSpan` attribute, so a merge appears as a wide non-empty cell
 * followed by empty placeholders in the columns it visually covers. When a non-empty cell's box
 * extends past the start of an adjacent empty column, that column is absorbed: the placeholder is
 * dropped and `colSpan` is set. Because only *empty* neighbours a cell's own geometry covers are
 * ever merged, a regular grid (whose cells sit inside their own column) is never given a spurious
 * span. Positional `col` indices are left untouched so the generator rebuilds the grid from
 * `col` + `colSpan`. No-op without geometry (cells then have no bounds to reason from).
 */
function inferColSpans(rows: OfficeContentNode[]): void {
    // Representative left edge per grid column, from the non-empty cells that occupy it.
    const colLefts: number[][] = [];
    for (const row of rows) {
        const cells = row.children || [];
        for (let c = 0; c < cells.length; c++) {
            const b = cells[c].bounds;
            if (b) (colLefts[c] ||= []).push(b.x);
        }
    }
    const colLeft = (c: number): number => {
        const xs = colLefts[c];
        return xs && xs.length ? medianOf(xs) : NaN;
    };
    for (const row of rows) {
        const cells = row.children || [];
        const kept: OfficeContentNode[] = [];
        let c = 0;
        while (c < cells.length) {
            const cell = cells[c];
            if (isEmptyCell(cell)) { kept.push(cell); c++; continue; }
            const b = cell.bounds;
            let span = 1;
            if (b) {
                const right = b.x + b.width;
                // Absorb consecutive empty columns to the right whose start this cell's box passes.
                while (c + span < cells.length && isEmptyCell(cells[c + span])) {
                    const nextLeft = colLeft(c + span);
                    if (!Number.isFinite(nextLeft) || right <= nextLeft + 1) break;
                    span++;
                }
            }
            if (span > 1 && cell.metadata) (cell.metadata as CellMetadata).colSpan = span;
            kept.push(cell);
            c += span;
        }
        row.children = kept;
    }
}

/**
 * Recovers vertical cell merges (rowSpan) the tag tree padded with empty placeholder cells. pdf.js
 * does not expose the `/RowSpan` attribute, so a merge appears as a tall non-empty cell with empty
 * placeholders in the rows below it, at the same grid column. When a non-empty cell's box extends
 * well past the top of the next row's band and the cell directly below (same grid column) is an
 * empty placeholder, that placeholder is absorbed: it is dropped and `rowSpan` grows. Only detects
 * merges whose spanning cell has content tall enough to reach into the rows it covers - a vertically
 * merged cell with a single centred line is invisible to geometry and left as a full grid (correct,
 * just not marked as merged). Runs on the grid keyed by the positional `col`, which the tag tree
 * keeps rectangular, so it stays aligned with {@link inferColSpans}.
 */
function inferRowSpans(rows: OfficeContentNode[]): void {
    if (rows.length < 2) return;
    const byCol: Map<number, OfficeContentNode>[] = [];
    const rowTop: number[] = [];
    for (const row of rows) {
        const m = new Map<number, OfficeContentNode>();
        let top = Infinity;
        for (const cell of row.children || []) {
            const col = (cell.metadata as CellMetadata)?.col;
            if (typeof col === 'number' && !m.has(col)) m.set(col, cell);
            if (cell.bounds) top = Math.min(top, cell.bounds.y);
        }
        byCol.push(m);
        rowTop.push(Number.isFinite(top) ? top : NaN);
    }
    // Typical row pitch drives the coverage margin, so a cell must clearly enter the next band
    // (not merely touch its top, which every ordinary cell does) to count as spanning.
    const pitches: number[] = [];
    for (let r = 1; r < rowTop.length; r++) {
        if (Number.isFinite(rowTop[r]) && Number.isFinite(rowTop[r - 1])) pitches.push(rowTop[r] - rowTop[r - 1]);
    }
    const pitch = medianOf(pitches.filter(p => p > 0));
    if (!(pitch > 0)) return;
    const margin = 0.4 * pitch;

    const toRemove = new Set<OfficeContentNode>();
    for (let r = 0; r < rows.length; r++) {
        for (const [col, cell] of byCol[r]) {
            if (toRemove.has(cell) || isEmptyCell(cell) || !cell.bounds) continue;
            const cSpan = (cell.metadata as CellMetadata)?.colSpan || 1;
            const bottom = cell.bounds.y + cell.bounds.height;
            let span = 1;
            for (let k = 1; r + k < rows.length; k++) {
                const below = byCol[r + k].get(col);
                if (!below || !isEmptyCell(below)) break;
                if (!Number.isFinite(rowTop[r + k]) || bottom < rowTop[r + k] + margin) break;
                toRemove.add(below);
                // A 2-D merge (this cell also spans columns) leaves placeholders under every column
                // it covers in the rows below; absorb those too, or they render as a phantom column.
                for (let cc = col + 1; cc < col + cSpan; cc++) {
                    const extra = byCol[r + k].get(cc);
                    if (extra && isEmptyCell(extra)) toRemove.add(extra);
                }
                span++;
            }
            if (span > 1 && cell.metadata) (cell.metadata as CellMetadata).rowSpan = span;
        }
    }
    if (toRemove.size) {
        for (const row of rows) row.children = (row.children || []).filter(c => !toRemove.has(c));
    }
}

function buildTable(node: StructNode, ctx: WalkCtx): OfficeContentNode | null {
    const rows: OfficeContentNode[] = [];
    let rowIdx = 0;
    for (const tr of collectRows(node)) {
        const cells: OfficeContentNode[] = [];
        let colIdx = 0;
        for (const cellNode of (tr.children || [])) {
            const cr = role(cellNode);
            if (cr !== 'TH' && cr !== 'TD') continue;
            const cellChildren = walkChildren(cellNode, ctx, 0);
            const meta: CellMetadata = { row: rowIdx, col: colIdx };
            if (cr === 'TH') meta.style = 'header';
            const cell: OfficeContentNode = {
                type: 'cell',
                text: cellChildren.map(n => n.text || '').join(' ').trim(),
                children: cellChildren,
                metadata: meta,
            };
            const cb = unionAll(cellChildren.map(c => c.bounds));
            if (cb) cell.bounds = cb;
            cells.push(cell);
            colIdx++;
        }
        if (!cells.length) continue;
        const row: OfficeContentNode = { type: 'row', children: cells, text: cells.map(c => c.text || '').join(' ').trim() };
        const rb = unionAll(cells.map(c => c.bounds));
        if (rb) row.bounds = rb;
        rows.push(row);
        rowIdx++;
    }
    if (!rows.length) return null;
    // Recover merged cells the tags padded with empty placeholders (needs geometry). Column spans
    // run first, on the still-rectangular grid (they use positional indices); row spans run after,
    // keyed by the geometry-stable `col`, so they tolerate the placeholders columns already dropped.
    if (ctx.doc.cfg.includePositions) { inferColSpans(rows); inferRowSpans(rows); }
    const table: OfficeContentNode = { type: 'table', children: rows, text: rows.map(r => r.text || '').join('\n') };
    const tb = unionAll(rows.map(r => r.bounds));
    if (tb) table.bounds = tb;
    return table;
}

/** Classifies a list marker label into an ordered/unordered list type. */
function classifyListType(label: string): 'ordered' | 'unordered' {
    const t = label.trim();
    // Decimal (incl. multilevel "1.1."), single alpha "a.", or roman "iii." markers are ordered.
    if (/^\(?\d+(\.\d+)*[.)]?$/.test(t) || /^[a-zA-Z][.)]$/.test(t) || /^[ivxlcdmIVXLCDM]+[.)]$/.test(t)) return 'ordered';
    return 'unordered';
}

/** Collapses TOC dot-leaders ("Title ...... 3" -> "Title 3") in a node's text and text-run children. */
function stripDotLeaders(node: OfficeContentNode): void {
    const clean = (s: string | undefined) => (s || '').replace(/\s*\.{4,}\s*/g, ' ');
    if (node.text) node.text = clean(node.text).trim();
    for (const c of node.children || []) {
        if (c.type === 'text') c.text = clean(c.text);
        else stripDotLeaders(c);
    }
}

/** Converts a roman numeral to its integer value, or 0 if it is not a valid roman numeral. */
function romanToInt(s: string): number {
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

/**
 * Parses the number an ordered-list label represents: decimal ("3."), multi-level decimal ("1.2.",
 * uses the last component), roman ("iii."), or single alpha ("c." -> 3). Returns null when the label
 * carries no recognizable number.
 */
function parseListNumber(label: string): number | null {
    const t = label.trim();
    if (!t) return null;
    // All-decimal (possibly multi-level like "1.2."): use the last numeric component.
    if (/^[\d.()\[\]\s]+$/.test(t)) {
        const parts = t.match(/\d+/g);
        if (parts && parts.length) return parseInt(parts[parts.length - 1], 10);
    }
    const core = t.replace(/^[(\[]+/, '').replace(/[.)\]\s]+$/, '');
    if (/^[ivxlcdm]+$/i.test(core)) { const n = romanToInt(core); return n > 0 ? n : null; }
    if (/^[a-z]$/i.test(core)) return core.toLowerCase().charCodeAt(0) - 96; // a -> 1
    return null;
}

function buildList(node: StructNode, ctx: WalkCtx, indent: number): OfficeContentNode[] {
    const items: OfficeContentNode[] = [];
    const listId = `pdf-list-${++ctx.listCounter.n}`;
    let idx = 0;
    for (const li of (node.children || [])) {
        if (role(li) !== 'LI') {
            // Some producers nest content directly; recurse transparently.
            if (role(li) === 'L') items.push(...buildList(li, ctx, indent + 1));
            continue;
        }
        let label = '';
        const bodyNodes: OfficeContentNode[] = [];
        const nested: OfficeContentNode[] = [];
        for (const part of li.children || []) {
            const pr = role(part);
            if (pr === 'Lbl') {
                label = collectRuns(part, ctx).map(r => r.text).join('').trim();
            } else if (pr === 'LBody') {
                // Direct text runs of the body (excluding nested lists/tables).
                const directRuns = collectRuns(part, ctx, new Set(['L', 'Table']));
                if (directRuns.length) { const p = runsToParagraph(directRuns, ctx.page, ctx.doc, 0); if (p) bodyNodes.push(p); }
                for (const b of part.children || []) {
                    const br = role(b);
                    if (br === 'L') nested.push(...buildList(b, ctx, indent + 1));
                    else if (br === 'Table') { const t = buildTable(b, ctx); if (t) bodyNodes.push(t); }
                }
            } else if (pr === 'L') {
                nested.push(...buildList(part, ctx, indent + 1));
            }
        }
        const listType = classifyListType(label);
        // Number ordered items from the label the PDF actually rendered (decimal "3.", roman "iii.",
        // or alpha "c.") rather than the positional index, so a list interrupted by a paragraph and
        // split into two <L> elements resumes its numbering (3, 4) instead of restarting at 1. Falls
        // back to the position when the label carries no parseable number.
        let itemIndex = idx;
        if (listType === 'ordered') {
            const n = parseListNumber(label);
            if (n !== null && n >= 1) itemIndex = n - 1;
        }
        const meta: ListMetadata = {
            listType,
            indentation: indent,
            alignment: 'left',
            listId,
            itemIndex,
        };
        const item: OfficeContentNode = {
            type: 'list',
            text: bodyNodes.map(n => n.text || '').join(' ').trim(),
            children: bodyNodes,
            metadata: meta,
        };
        const b = unionAll(bodyNodes.map(n => n.bounds));
        if (b) item.bounds = b;
        items.push(item);
        items.push(...nested);
        idx++;
    }
    return items;
}

function buildNote(node: StructNode, ctx: WalkCtx): OfficeContentNode | null {
    const children = walkChildren(node, ctx, 0);
    if (!children.length) {
        const p = paragraphFrom(node, ctx, 0);
        if (!p) return null;
        children.push(p);
    }
    const meta: NoteMetadata = { noteType: 'footnote' };
    // Strip the leading marker glyph the PDF renders inside the note body ("1 In paged media…") so it
    // is not duplicated next to the generated citation ("[^1]: 1 In paged…"); keep it as the note id.
    const markerRe = /^\s*([0-9]+|[ivxlcdm]+|[a-z]|[*†‡§])[.)]?\s+/i;
    const firstText = children.map(n => n.text || '').join(' ').trim();
    const mk = firstText.match(markerRe);
    if (mk) {
        meta.noteId = mk[1];
        const strip = (n: OfficeContentNode): boolean => {
            if (n.type === 'text' && n.text) { const m = n.text.match(markerRe); if (m) { n.text = n.text.slice(m[0].length); return true; } return false; }
            if (n.text) { const m = n.text.match(markerRe); if (m) n.text = n.text.slice(m[0].length); }
            for (const c of n.children || []) if (strip(c)) return true;
            return false;
        };
        for (const child of children) if (strip(child)) break;
    }
    const note: OfficeContentNode = { type: 'note', text: children.map(n => n.text || '').join(' ').trim(), children, metadata: meta };
    const b = unionAll(children.map(n => n.bounds));
    if (b) note.bounds = b;
    return note;
}
