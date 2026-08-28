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
 * `getTextContent` derives an item's `transform` from `ctm ∘ textMatrix` and, in its main pass,
 * ignores the page `cm`/`q`/`Q` (only form-XObject `/Matrix` moves its ctm). We ignore them here too,
 * so both share the same cm-less frame and their coordinates coincide. Pattern/shading/transparent
 * fills yield no color, so a run over them simply keeps its default (never a wrong color).
 *
 * @module parsers/pdf/pdfColor
 */

const IDENTITY = [1, 0, 0, 1, 0, 0];

/** Composes two affine matrices, `pdfjs.Util.transform(a, b)`, inlined to keep pdf.js types out. */
function mul(a: number[], b: number[]): number[] {
    return [
        a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
        a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
        a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5],
    ];
}

/** Applies the text-move translation `[x, y]` to a matrix, as pdf.js's `TextState` does. */
function translate(m: number[], x: number, y: number): number[] {
    return [m[0], m[1], m[2], m[3], m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}

/**
 * Coerces a pdf.js matrix operand to a plain `number[6]`, or null if it is not a 6-element
 * array-like. pdf.js passes the matrix for `setTextMatrix`/`paintFormXObjectBegin` as a single
 * argument that is a `Float32Array` (so `Array.isArray` is false and its six values sit at `[0]`).
 */
function toMat6(m: any): number[] | null {
    if (!m || typeof m.length !== 'number' || m.length < 6) return null;
    return [+m[0], +m[1], +m[2], +m[3], +m[4], +m[5]];
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
 * form-XObject CTM; page `cm`/`q`/`Q` ignored). Marks are skipped when the active fill is a pattern,
 * shading or transparent, or when the text is invisible (render modes 3/7); stroke-render text
 * (modes 1/5) uses the stroke color.
 *
 * @param viewportTransform - `viewport.transform` for the page (scale-1 layout viewport).
 */
export function collectColorMarks(ops: any, viewportTransform: number[], OPS: any): ColorMark[] {
    const marks: ColorMark[] = [];
    const fnArray: number[] | undefined = ops?.fnArray;
    const argsArray: any[] | undefined = ops?.argsArray;
    if (!fnArray || !argsArray) return marks;

    let ctm = IDENTITY.slice();
    const ctmStack: number[][] = [];
    let tm = IDENTITY.slice();
    let tlm = IDENTITY.slice();
    let leading = 0, rise = 0, renderMode = 0;
    let fill: string | undefined;   // current fill color hex, or undefined for pattern/unknown
    let stroke: string | undefined;
    // The fill/stroke color IS part of the graphics state, so it must be saved/restored across q/Q
    // (and form XObjects), unlike the CTM which we deliberately ignore for q/Q to stay in
    // getTextContent's coordinate frame. Without this, a color set inside a `q ... Q` bracket (a
    // border, a background, a logo) would leak onto all following text.
    const colorStack: { fill: string | undefined; stroke: string | undefined }[] = [];

    for (let i = 0; i < fnArray.length; i++) {
        const a = argsArray[i];
        switch (fnArray[i]) {
            // Graphics-state save/restore: track color only (the CTM is left in getTextContent's frame).
            case OPS.save: colorStack.push({ fill, stroke }); break;
            case OPS.restore: { const s = colorStack.pop(); if (s) { fill = s.fill; stroke = s.stroke; } break; }

            // CTM: only form-XObject matrices move it (page cm/q/Q are ignored, as getTextContent does).
            // A form localizes graphics state, so its color is saved/restored around it too.
            case OPS.paintFormXObjectBegin: {
                ctmStack.push(ctm);
                colorStack.push({ fill, stroke });
                const fm = toMat6(a?.[0]);
                if (fm) ctm = mul(ctm, fm);
                break;
            }
            case OPS.paintFormXObjectEnd: {
                ctm = ctmStack.pop() ?? IDENTITY.slice();
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
            case OPS.beginText: tm = IDENTITY.slice(); tlm = IDENTITY.slice(); break;
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
export function makeColorLookup(marks: ColorMark[]): ColorLookup {
    const rows = new Map<number, ColorMark[]>();
    for (const mk of marks) {
        const k = Math.round(mk.y);
        const bucket = rows.get(k);
        if (bucket) bucket.push(mk); else rows.set(k, [mk]);
    }
    return (vx: number, vy: number, fontSize: number, width: number): string | undefined => {
        const yTol = Math.max(2, fontSize * 0.5);
        const pad = Math.max(1, fontSize * 0.5);
        const lo = vx - pad, hi = vx + width + pad;
        let found: string | undefined;
        for (let k = Math.round(vy - yTol); k <= Math.round(vy + yTol); k++) {
            const bucket = rows.get(k);
            if (!bucket) continue;
            for (const mk of bucket) {
                if (mk.x < lo || mk.x > hi) continue;
                if (found === undefined) found = mk.color;
                else if (found !== mk.color) return undefined; // colour change inside one run: don't guess
            }
        }
        return found;
    };
}
