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
    const table: OfficeContentNode = { type: 'table', children: rows, text: rows.map(r => r.text || '').join('\n') };
    const tb = unionAll(rows.map(r => r.bounds));
    if (tb) table.bounds = tb;
    return table;
}

/** Classifies a list marker label into an ordered/unordered list type. */
function classifyListType(label: string): 'ordered' | 'unordered' {
    const t = label.trim();
    if (/^\(?\d+[.)]?$/.test(t) || /^[a-zA-Z][.)]$/.test(t) || /^[ivxlcdmIVXLCDM]+[.)]$/.test(t)) return 'ordered';
    return 'unordered';
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
        const meta: ListMetadata = {
            listType: classifyListType(label),
            indentation: indent,
            alignment: 'left',
            listId,
            itemIndex: idx,
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
    const note: OfficeContentNode = { type: 'note', text: children.map(n => n.text || '').join(' ').trim(), children, metadata: meta };
    const b = unionAll(children.map(n => n.bounds));
    if (b) note.bounds = b;
    return note;
}
