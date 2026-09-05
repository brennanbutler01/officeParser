/**
 * OfficeTemplate: fill a document template's `{{placeholder}}` tags with data and get back new
 * document(s) - a lightweight, dependency-free mail-merge / document-generation entry point.
 *
 * This is deliberately separate from parsing and generation: it does not build an AST or convert
 * anything, it copies the template and substitutes only the placeholders, so every bit of the
 * original formatting, layout and structure is preserved. Give it one data object to get one
 * document back, or an array to get one document per entry (a batch mail-merge).
 *
 * Today it templates DOCX (`.docx`); the run-aware substitution generalizes to the other OOXML/ODF
 * text formats later. See {@link module:template/docxTemplate} for the substitution details.
 *
 * @example
 * ```typescript
 * import { OfficeTemplate } from 'officeparser';
 * import { writeFileSync } from 'fs';
 *
 * // One document.
 * const bytes = await OfficeTemplate.render('invoice-template.docx', {
 *   data: { name: 'Acme Corp', amount: '$1,250.00', due: '2026-10-01' },
 * });
 * writeFileSync('invoice-acme.docx', bytes);
 *
 * // A batch: one document per row.
 * const docs = await OfficeTemplate.render('invoice-template.docx', {
 *   data: [ { name: 'Acme Corp', amount: '$1,250.00' }, { name: 'Globex', amount: '$980.00' } ],
 * });
 * docs.forEach((d, i) => writeFileSync(`invoice-${i}.docx`, d));
 * ```
 *
 * @module OfficeTemplate
 */

import { isCfb, isEncryptedOoxml, decryptOoxml, WRONG_PASSWORD } from './crypto/index.js';
import { openDocx, renderDocxTemplate } from './template/docxTemplate.js';
import { BlobLike, OfficeErrorType, OfficeParserConfig, TemplateConfig, TemplateData } from './types.js';
import { assertNode } from './utils/envUtils.js';
import { getOfficeError } from './utils/errorUtils.js';
import { resolveZipInstant } from './utils/officeGenUtils.js';

/** Accepted template inputs, matching `parseOffice`'s first argument. */
export type TemplateInput = string | Buffer | ArrayBuffer | Uint8Array | BlobLike;

/** Reads the template input into a Node Buffer (path via fs in Node; bytes/blob anywhere). */
async function readInput(input: TemplateInput): Promise<Buffer> {
    if (typeof input === 'string') {
        assertNode('path-parsing');
        const fs = await import('fs');
        return fs.readFileSync(input);
    }
    if (input instanceof ArrayBuffer) return Buffer.from(input);
    if (input instanceof Uint8Array) return Buffer.from(input.buffer, input.byteOffset, input.byteLength);
    if (Buffer.isBuffer(input)) return input;
    if (input && typeof (input as BlobLike).arrayBuffer === 'function') return Buffer.from(await (input as BlobLike).arrayBuffer());
    throw getOfficeError(OfficeErrorType.INVALID_INPUT);
}

/**
 * Template renderer. See {@link module:OfficeTemplate} for an overview and {@link TemplateConfig} for
 * the options.
 */
export class OfficeTemplate {
    /** Renders a single document from a template and one data map. */
    static render(template: TemplateInput, config: TemplateConfig & { data: TemplateData }): Promise<Uint8Array>;
    /** Renders one document per entry from a template and an array of data maps (a batch mail-merge). */
    static render(template: TemplateInput, config: TemplateConfig & { data: TemplateData[] }): Promise<Uint8Array[]>;
    static render(template: TemplateInput, config: TemplateConfig): Promise<Uint8Array | Uint8Array[]>;
    static async render(template: TemplateInput, config: TemplateConfig): Promise<Uint8Array | Uint8Array[]> {
        if (!config || config.data == null) throw getOfficeError(OfficeErrorType.INVALID_INPUT);

        let bytes = await readInput(template);

        // Decrypt an encrypted template up front (reuses the parser's crypto), so rendering sees plaintext.
        if (isCfb(bytes) && isEncryptedOoxml(bytes)) {
            if (!config.password) throw getOfficeError(OfficeErrorType.PASSWORD_REQUIRED);
            try { bytes = Buffer.from(decryptOoxml(bytes, config.password)); }
            catch (e) {
                if (e === WRONG_PASSWORD) throw getOfficeError(OfficeErrorType.PASSWORD_INCORRECT);
                throw getOfficeError(OfficeErrorType.DOCUMENT_DECRYPTION_FAILED, undefined, e instanceof Error ? e.message : String(e));
            }
        }

        const entries = openDocx(bytes);
        if (!entries) {
            const hint = config.fileType && config.fileType !== 'docx' ? `'${config.fileType}' is not supported` : 'the input is not a .docx';
            throw getOfficeError(OfficeErrorType.TEMPLATE_UNSUPPORTED_FORMAT, undefined, hint);
        }

        const start = config.delimiters?.start || '{{';
        const end = config.delimiters?.end || '}}';
        const onMissing = config.onMissing || 'keep';
        // Pin the zip mtime so a given (template, data) renders to identical bytes every time.
        const { mtime } = resolveZipInstant('1980-01-01T00:00:00Z');
        const onFieldMissing = (key: string): never => { throw getOfficeError(OfficeErrorType.TEMPLATE_FIELD_MISSING, undefined, key); };

        const renderOne = (data: TemplateData): Uint8Array =>
            renderDocxTemplate(entries, data, { start, end, onMissing, mtime, onFieldMissing });

        return Array.isArray(config.data) ? config.data.map(renderOne) : renderOne(config.data);
    }
}

/** Standalone alias for {@link OfficeTemplate.render}. */
export const renderTemplate = OfficeTemplate.render;
