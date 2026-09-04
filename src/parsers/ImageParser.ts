/**
 * Image parser: treats a raster image (a scan, a photo of a page, a screenshot) as a one-page
 * document whose only text layer is whatever OCR can recognize.
 *
 * The AST mirrors a scanned PDF page: a single `page` node whose children are `paragraph` nodes
 * rebuilt from the recognized text, one per blank-line-separated block. Lines inside a block keep
 * their own breaks and indentation, so a layout-reconstructed table or column (see
 * `ocrConfig.preserveLayout`) still reads spatially in plain-text output while HTML/DOCX flow it.
 * The image itself is returned in `attachments` (carrying the same `ocrText`) when
 * `extractAttachments` is set.
 *
 * Unlike PDF, where OCR runs over *extracted* page images and so needs `extractAttachments`, here
 * the image is the input, so OCR runs on the bytes directly; `extractAttachments` only controls
 * whether the image is also handed back. With `ocr` off there is nothing to extract: the page is
 * empty (or holds one `image` node pointing at the attachment when one was extracted) and an
 * `IMAGE_NO_TEXT_EXTRACTED` warning says why, mirroring `PDF_NO_TEXT_EXTRACTED` for scanned PDFs.
 *
 * @module parsers/ImageParser
 */

import { FullOfficeParserConfig, ImageMetadata, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeWarningType, PageMetadata, SupportedFileType } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { checkAbortSignal, logWarning } from '../utils/errorUtils.js';
import { createAttachment, getMimeFromExtension } from '../utils/imageUtils.js';
import { performOcr } from '../utils/ocrUtils.js';
import { sniffImageSize } from '../utils/officeGenUtils.js';

/**
 * Splits recognized text into paragraph nodes, one per blank-line-separated block. Within a block
 * every line is kept (with its leading indentation, which the OCR layout reconstruction uses to
 * encode columns); only trailing whitespace and surrounding blank lines are dropped.
 */
function paragraphsFromText(text: string): OfficeContentNode[] {
    const out: OfficeContentNode[] = [];
    for (const block of text.split(/\n[ \t]*\n+/)) {
        const lines = block.split('\n').map(l => l.trimEnd());
        while (lines.length && !lines[0].trim()) lines.shift();
        while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
        if (!lines.length) continue;
        const t = lines.join('\n');
        out.push({ type: 'paragraph', text: t, children: [{ type: 'text', text: t }] });
    }
    return out;
}

/**
 * Parses a raster image into a one-page AST of OCR-recognized text.
 *
 * @param buffer - The image bytes
 * @param fileType - The canonical image type (`png`, `jpg`, `gif`, `bmp`, `tiff`, `webp`)
 * @param config - Parser configuration; `ocr` enables recognition, `extractAttachments` returns the image
 * @returns A promise resolving to the parsed AST
 */
export const parseImage = async (buffer: Buffer, fileType: SupportedFileType, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    checkAbortSignal(config.abortSignal);
    const attachmentName = `image.${fileType}`;

    let ocrText = '';
    if (config.ocr) {
        try { ocrText = (await performOcr(buffer, { ...config.ocrConfig })).trim(); }
        catch (e) { logWarning(OfficeWarningType.OCR_FAILED, config, attachmentName, e); }
    } else {
        logWarning(OfficeWarningType.IMAGE_NO_TEXT_EXTRACTED, config);
    }

    const attachments: OfficeAttachment[] = [];
    if (config.extractAttachments) {
        const attachment = createAttachment(attachmentName, buffer);
        if (ocrText) attachment.ocrText = ocrText;
        attachments.push(attachment);
    }

    const children = paragraphsFromText(ocrText);
    // Nothing recognized (OCR off, failed, or a blank image): when the image was extracted, point
    // the page at it so the document is not silently empty and generators can still render it.
    if (!children.length && attachments.length) {
        const meta: ImageMetadata = { attachmentName };
        children.push({ type: 'image', text: '', metadata: meta });
    }

    const pageMeta: PageMetadata = { pageNumber: 1 };
    const page: OfficeContentNode = {
        type: 'page',
        text: children.map(c => c.text || '').filter(Boolean).join('\n'),
        children,
        metadata: pageMeta,
    };

    const nativeProperties: Record<string, any> = { mimeType: getMimeFromExtension(fileType) };
    const size = sniffImageSize(buffer);
    if (size) { nativeProperties.pixelWidth = size.w; nativeProperties.pixelHeight = size.h; }

    return createAST(fileType, { nativeProperties }, [page], attachments, config, undefined);
};
