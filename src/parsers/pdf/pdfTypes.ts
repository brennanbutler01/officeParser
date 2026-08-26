/**
 * Internal shared types for the PDF extraction pipeline.
 *
 * These never leave the parser; the public AST types live in `../../types.ts`. Keeping them here
 * lets `geometry.ts`, `textLayout.ts` and `structTree.ts` exchange plain data with no pdf.js
 * objects in their signatures, so each stage is unit-testable in isolation.
 *
 * @module parsers/pdf/pdfTypes
 */

import { NodeBounds, TextFormatting, TextMetadata } from '../../types.js';

/** A run's writing angle after snapping to the nearest quarter turn. */
export type RunAngle = 0 | 90 | 180 | 270;

/** Resolved font attributes for one pdf.js font id, looked up once per document. */
export interface ResolvedFont {
    /** Human-readable family name with any subset prefix stripped (e.g. "Ubuntu Mono"). */
    name?: string;
    bold: boolean;
    italic: boolean;
    /** Ascent as a fraction of font size (from textContent.styles), used for top-edge geometry. */
    ascent: number;
    /** Descent as a fraction of font size (negative). */
    descent: number;
    /** True for vertical (top-to-bottom) writing mode fonts. */
    vertical: boolean;
}

/**
 * One pdf.js text item, normalized into top-left-origin viewport space (PDF points, y-down) and
 * with its font already resolved. This is the atom every later stage consumes.
 */
export interface RawRun {
    text: string;
    /** Left edge in viewport space. */
    x: number;
    /** Top edge (baseline minus ascent*fontSize). */
    yTop: number;
    /** Baseline y. */
    yBaseline: number;
    width: number;
    height: number;
    fontSize: number;
    /** pdf.js internal font id (e.g. "g_d0_f1"), used only to group same-font runs. */
    fontKey: string;
    dir: 'ltr' | 'rtl' | 'ttb';
    /** pdf.js hasEOL: this run visually ends its line. */
    hasEOL: boolean;
    /** Snapped writing angle. Only 0 goes through the normal flow; others are appended after. */
    angle: RunAngle;
    /** Marked-content id for the tagged path, or null. */
    mcid: string | null;
    /** True when the run sits inside an Artifact scope (running header/footer, decoration). */
    inArtifact: boolean;
    formatting: TextFormatting;
    /** Hyperlink metadata resolved from annotations, if any. */
    link?: TextMetadata;
}

/** A contiguous span of text on a single visual line, with its own box and formatting. */
export interface TextFragment {
    text: string;
    bounds: NodeBounds;
    formatting: TextFormatting;
    link?: TextMetadata;
}

/** A visual line: fragments sharing a baseline, plus derived geometry used by later stages. */
export interface PdfLine {
    fragments: TextFragment[];
    x: number;
    yTop: number;
    width: number;
    height: number;
    baseline: number;
    /** Char-weighted dominant font size of the line. */
    fontSize: number;
    /** True when the majority of characters are bold. */
    bold: boolean;
    /** True when the line's visible text ends in a hyphen (candidate for de-hyphenation). */
    endsWithHyphen: boolean;
    dir: 'ltr' | 'rtl';
}

/** An image placed on a page, in viewport space. */
export interface PdfImage {
    name: string;
    bounds: NodeBounds;
    data: Uint8Array | Uint8ClampedArray;
    pixelWidth: number;
    pixelHeight: number;
    kind?: number;
}

/** Everything collected from one page in the single collection pass. */
export interface PageExtract {
    pageNumber: number;
    /** Post-rotation viewport width in points (what pdf.js renders). */
    width: number;
    /** Post-rotation viewport height in points. */
    height: number;
    /** Authored (rotation-0) viewport width; layout runs in this frame, bounds map back at emit. */
    authoredW: number;
    /** Authored (rotation-0) viewport height. */
    authoredH: number;
    /** Page /Rotate in degrees (0/90/180/270). */
    rotation: number;
    runs: RawRun[];
    images: PdfImage[];
    /** Raw pdf.js struct tree for the tagged path; null when untagged/unavailable. */
    structTree: unknown | null;
}

/** Resolved PDF-specific knobs, read from `config.pdfParserConfig` with numeric coercion done. */
export interface PdfLayoutConfig {
    useTags: boolean;
    detectColumns: boolean;
    mergeHyphenatedWords: boolean;
    lineToleranceFactor: number;
    spaceToleranceFactor: number;
    headingDetection: 'auto' | 'font-size' | 'off';
    disableTextNormalization: boolean;
    /** When false, geometry is not written onto emitted nodes. */
    includePositions: boolean;
}
