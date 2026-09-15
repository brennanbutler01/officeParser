/**
 * Small shared numeric helpers used across the PDF layout/geometry code, the OCR reconstructor, and
 * the text generator, so the same statistic is computed one way everywhere.
 *
 * @module numberUtils
 */

/**
 * Upper median of a numeric list: the value at `floor(n/2)` of the sorted list (0 for empty). Not
 * the mean of the two middle values - the geometry heuristics want an actual observed value (a real
 * line pitch, column width or character width), and the extra branch buys nothing here.
 */
export function median(values: number[]): number {
    if (!values.length) return 0;
    const s = [...values].sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
}

/**
 * Clamps a document-derived repeat count to a safe range before it reaches `String.prototype.repeat`,
 * so a hostile file cannot turn a tiny attribute (an ODF `text:c`, a list `ilvl`/indentation) into a
 * multi-gigabyte string, and a negative/NaN count cannot throw. Returns 0 for NaN/negative and caps at
 * `max`. No real document repeats a character thousands of times or nests thousands deep.
 */
export function clampRepeat(count: number, max = 10000): number {
    return Number.isFinite(count) && count > 0 ? Math.min(Math.floor(count), max) : 0;
}
