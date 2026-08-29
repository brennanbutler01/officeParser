/**
 * Browser-side stub for 'pdf-lib'.
 *
 * pdf-lib is an optional peer dependency used only by the native PDF engine
 * (pdfConfig.engine: 'native'). It is deliberately NOT bundled into the prebuilt
 * browser bundle, so callers who want the native engine in the browser install
 * pdf-lib and bundle officeParser themselves (the dynamic import then resolves to
 * the real package). In the prebuilt bundle, selecting the native engine throws a
 * clear, actionable error; the default 'html' engine (window.print) needs nothing.
 */

const unavailable = () => {
    throw new Error("officeparser: the native PDF engine requires 'pdf-lib', which is not included in the prebuilt browser bundle. Install pdf-lib and bundle officeParser yourself, or use pdfConfig.engine: 'html'.");
};

export const PDFDocument = { create: unavailable };
export const StandardFonts = {};
export const rgb = () => ({});
export const degrees = () => 0;

export default { PDFDocument, StandardFonts, rgb, degrees };
