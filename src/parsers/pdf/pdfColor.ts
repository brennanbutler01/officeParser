/**
 * Per-run text fill color from a page's operator list (pdf.js exposes no color on `getTextContent`).
 *
 * pdf.js's worker normalizes every fill-color operator into a single `OPS.setFillRGBColor` whose one
 * argument is an already-computed lowercase `#rrggbb` string (gray, CMYK, indexed, ICC and separation
 * spaces are all pre-converted for us), and every text-show operator into `OPS.showText`. So we never
 * touch color spaces or glyphs: we track the current fill color and, by mirroring exactly the text
 * matrix that `getTextContent` uses, compute each show's start point in the same viewport space as a
 * text item's `transform`. That lets a run look up its color by position.
 *
 * `getTextContent` derives an item's `transform` from `ctm ∘ textMatrix`, where its `ctm` is the full
 * graphics-state matrix: `q`/`Q` save and restore it, every page `cm` multiplies into it, and a form
 * XObject's `/Matrix` composes on entry. We mirror all four here, so both share one frame and their
 * coordinates coincide. That matters because most producers open the page with a `cm` - Chrome/Skia
 * flips with `1 0 0 -1 0 H cm`, Ghostscript scales with `0.1 0 0 0.1 0 0 cm`, Cairo flips too - and
 * a cm-less frame would put every mark somewhere the runs never are. Pattern/shading/transparent
 * fills yield no color, so a run over them simply keeps its default (never a wrong color).
 *
 * @module parsers/pdf/pdfColor
 */

import { identityMatrix, mulMatrix as mul, toMatrix6 as toMat6 } from './geometry.js';

/** Applies the text-move translation `[x, y]` to a matrix, as pdf.js's `TextState` does. */
function translate(m: number[], x: number, y: number): number[] {
    return [m[0], m[1], m[2], m[3], m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

/** One text-show occurrence: its start point in viewport space plus the active fill color. */
export interface ColorMark {
    x: number;
    y: number;
    color: string;
}

/**
 * Walks an already-fetched operator list and returns, in paint order, the viewport-space start point
 * and fill color of every text-show. Mirrors `getTextContent`'s matrix bookkeeping (text matrix plus
 * the full CTM: `q`/`Q`, page `cm` and form-XObject matrices). Marks are skipped when the active fill
 * is a pattern, shading or transparent, or when the text is invisible (render modes 3/7);
 * stroke-render text (modes 1/5) uses the stroke color.
 *
 * @param viewportTransform - `viewport.transform` for the page (scale-1 layout viewport).
 */
export function collectColorMarks(ops: any, viewportTransform: number[], OPS: any): ColorMark[] {
    const marks: ColorMark[] = [];
    const fnArray: number[] | undefined = ops?.fnArray;
    const argsArray: any[] | undefined = ops?.argsArray;
    if (!fnArray || !argsArray) return marks;

    let ctm = identityMatrix();
    const ctmStack: number[][] = [];
    let tm = identityMatrix();
    let tlm = identityMatrix();
    let leading = 0, rise = 0, renderMode = 0;
    let fill: string | undefined;   // current fill color hex, or undefined for pattern/unknown
    let stroke: string | undefined;
    // The CTM and the fill/stroke color are both graphics state, so both are saved and restored
    // together across q/Q (and form XObjects). Without the color half, a color set inside a
    // `q ... Q` bracket (a border, a background, a logo) would leak onto all following text; without
    // the CTM half, every mark on a page that opens with a `cm` lands in the wrong place.
    const colorStack: { fill: string | undefined; stroke: string | undefined }[] = [];

    for (let i = 0; i < fnArray.length; i++) {
        const a = argsArray[i];
        switch (fnArray[i]) {
            // Graphics-state save/restore: CTM and color together, exactly as pdf.js's StateManager.
            case OPS.save: ctmStack.push(ctm); colorStack.push({ fill, stroke }); break;
            case OPS.restore: {
                if (ctmStack.length) ctm = ctmStack.pop()!;
                const s = colorStack.pop(); if (s) { fill = s.fill; stroke = s.stroke; }
                break;
            }

            // Page `cm`: composes into the CTM, as getTextContent's Util.transform(ctm, args) does.
            case OPS.transform: { const cm = toMat6(a); if (cm) ctm = mul(ctm, cm); break; }

            // A form XObject composes its /Matrix and localizes graphics state, colors included.
            case OPS.paintFormXObjectBegin: {
                ctmStack.push(ctm);
                colorStack.push({ fill, stroke });
                const fm = toMat6(a?.[0]);
                if (fm) ctm = mul(ctm, fm);
                break;
            }
            case OPS.paintFormXObjectEnd: {
                if (ctmStack.length) ctm = ctmStack.pop()!;
                const s = colorStack.pop(); if (s) { fill = s.fill; stroke = s.stroke; }
                break;
            }

            // Fill / stroke color: all spaces arrive pre-normalized to an *RGBColor hex string.
            case OPS.setFillRGBColor: fill = typeof a?.[0] === 'string' ? a[0] : undefined; break;
            case OPS.setFillColorN:
            case OPS.setFillTransparent: fill = undefined; break;
            case OPS.setStrokeRGBColor: stroke = typeof a?.[0] === 'string' ? a[0] : undefined; break;
            case OPS.setStrokeColorN:
            case OPS.setStrokeTransparent: stroke = undefined; break;

            // Text matrix (mirror of pdf.js TextState).
            case OPS.beginText: tm = identityMatrix(); tlm = identityMatrix(); break;
            case OPS.setTextMatrix: { const tmx = toMat6(a?.[0]); if (tmx) { tm = tmx; tlm = tmx.slice(); } break; }
            case OPS.moveText: tlm = translate(tlm, +a[0], +a[1]); tm = tlm.slice(); break;
            case OPS.setLeadingMoveText: leading = -(+a[1]); tlm = translate(tlm, +a[0], +a[1]); tm = tlm.slice(); break;
            case OPS.setLeading: leading = +a[0]; break;
            case OPS.nextLine: tlm = translate(tlm, 0, -leading); tm = tlm.slice(); break;
            case OPS.setTextRise: rise = +a[0]; break;
            case OPS.setTextRenderingMode: renderMode = +a[0]; break;

            case OPS.showText: {
                if (renderMode === 3 || renderMode === 7) break;               // invisible text
                const color = (renderMode === 1 || renderMode === 5) ? stroke : fill;
                if (!color) break;                                             // pattern / transparent
                // Start point = (ctm ∘ tm) applied to (0, rise), then into viewport space.
                const ux = tm[2] * rise + tm[4];
                const uy = tm[3] * rise + tm[5];
                const wx = ctm[0] * ux + ctm[2] * uy + ctm[4];
                const wy = ctm[1] * ux + ctm[3] * uy + ctm[5];
                const vx = viewportTransform[0] * wx + viewportTransform[2] * wy + viewportTransform[4];
                const vy = viewportTransform[1] * wx + viewportTransform[3] * wy + viewportTransform[5];
                marks.push({ x: vx, y: vy, color });
                break;
            }
        }
    }
    return marks;
}

/**
 * A run's color lookup. Given the run's start point `(vx, vy)`, its font size and its width, returns
 * the fill color, or undefined. See {@link makeColorLookup}.
 */
export type ColorLookup = (vx: number, vy: number, fontSize: number, width: number) => string | undefined;

/**
 * Builds a color lookup over a page's marks, bucketed by rounded viewport-y. A run is coloured only
 * when the marks lying on its baseline *within its own horizontal extent* `[vx, vx+width]` agree on a
 * single colour: that colour is returned. If no mark falls inside the run it keeps its default (so a
 * default-black run never borrows a neighbour's colour), and if the marks disagree - a colour change
 * mid-run that pdf.js merged into one text item - it returns undefined rather than confidently
 * picking the wrong one.
 */
/**
 * A page with more colour marks than this is not coloured at all (the lookup returns undefined for
 * every run). A real page has hundreds to low thousands of coloured pieces; this ceiling is far above
 * any genuine document and exists so a crafted content stream that emits millions of marks cannot make
 * the per-page build (a sort of every bucket) or the queries do unbounded work.
 */
const MAX_COLOR_MARKS = 300_000;

export function makeColorLookup(marks: ColorMark[]): ColorLookup {
    if (marks.length > MAX_COLOR_MARKS) return () => undefined;

    // Bucket by rounded viewport-y, then sort each bucket by x once and precompute, per bucket, a
    // prefix count of colour changes along x. A run's query then answers "do the marks in its
    // horizontal window agree, and on which colour?" with two binary searches and one prefix
    // subtraction - O(log n) per bucket - instead of scanning every mark. Without this, a page that
    // crowds N marks and N runs on one baseline is O(N^2): a CPU-hang DoS, because colour is default-on.
    interface Bucket { xs: number[]; colors: string[]; changePrefix: Int32Array; }
    const rows = new Map<number, ColorMark[]>();
    for (const mk of marks) {
        const k = Math.round(mk.y);
        const bucket = rows.get(k);
        if (bucket) bucket.push(mk); else rows.set(k, [mk]);
    }
    const buckets = new Map<number, Bucket>();
    for (const [k, arr] of rows) {
        arr.sort((a, b) => a.x - b.x);
        const xs = new Array<number>(arr.length);
        const colors = new Array<string>(arr.length);
        // changePrefix[i] = number of colour changes in colors[0..i]. colors[a..b] are all one colour
        // iff changePrefix[b] === changePrefix[a], so a range's homogeneity is an O(1) comparison.
        const changePrefix = new Int32Array(arr.length);
        for (let i = 0; i < arr.length; i++) {
            xs[i] = arr[i].x;
            colors[i] = arr[i].color;
            changePrefix[i] = i === 0 ? 0 : changePrefix[i - 1] + (colors[i] !== colors[i - 1] ? 1 : 0);
        }
        buckets.set(k, { xs, colors, changePrefix });
    }
    // First index with xs[i] >= target.
    const lowerBound = (xs: number[], target: number): number => {
        let lo = 0, hi = xs.length;
        while (lo < hi) { const mid = (lo + hi) >> 1; if (xs[mid] < target) lo = mid + 1; else hi = mid; }
        return lo;
    };
    // Last index with xs[i] <= target, or -1 when none.
    const upperIndex = (xs: number[], target: number): number => {
        let lo = 0, hi = xs.length;
        while (lo < hi) { const mid = (lo + hi) >> 1; if (xs[mid] <= target) lo = mid + 1; else hi = mid; }
        return lo - 1;
    };
    return (vx: number, vy: number, fontSize: number, width: number): string | undefined => {
        // Cap the vertical window independent of font size: a colour mark on the run's baseline sits
        // within about half the font size, but the loop steps one integer y per row, so an unclamped
        // `fontSize` (from a hostile text matrix) would iterate billions of empty rows. A colour on a
        // real line is always within 64pt; this bounds the loop to ~129 buckets. The horizontal pad is
        // clamped the same way so the window never spans thousands of points on a giant font.
        const yTol = Math.min(Math.max(2, fontSize * 0.5), 64);
        const pad = Math.min(Math.max(1, fontSize * 0.5), 64);
        const lo = vx - pad, hi = vx + width + pad;
        let found: string | undefined;
        for (let k = Math.round(vy - yTol); k <= Math.round(vy + yTol); k++) {
            const b = buckets.get(k);
            if (!b) continue;
            const a = lowerBound(b.xs, lo);
            const z = upperIndex(b.xs, hi);
            if (a > z) continue; // no mark in this bucket within the run's horizontal window
            // A colour change inside the run's window means pdf.js merged two colours into one item:
            // do not guess.
            if (b.changePrefix[z] !== b.changePrefix[a]) return undefined;
            const c = b.colors[a];
            if (found === undefined) found = c;
            else if (found !== c) return undefined; // two baselines within the window disagree
        }
        return found;
    };
}
