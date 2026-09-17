import * as assert from 'node:assert/strict';
import { getEventListeners } from 'node:events';
import { zipSync, strToU8 } from 'fflate';
import { OfficeParser } from '../../src/OfficeParser';
import { detectOfficeTypeFromZip, extractFiles } from '../../src/utils/zipUtils';

const quiet = { outputErrorToConsole: false };

// Stored entries make the compressed input span multiple feed chunks without
// relying on random data, expensive compression, or a machine-speed deadline.
const archive = Buffer.from(zipSync({
    'first.bin': new Uint8Array(128 * 1024),
    mimetype: strToU8('application/vnd.oasis.opendocument.text'),
    'content.xml': strToU8('<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>Example text</text:p></office:text></office:body></office:document-content>'),
}, { level: 0 }));

export async function archiveCancellationTests(): Promise<void> {
    console.log('- archive cancellation...');
    const preAborted = new AbortController();
    preAborted.abort();
    let inspected = 0;
    await assert.rejects(extractFiles(archive, () => { inspected++; return true; }, {}, {
        ...quiet, abortSignal: preAborted.signal,
    }), { name: 'AbortError' });
    assert.equal(inspected, 0, 'pre-aborted extraction must not inspect entries');

    await assertMidExtractionAbort(archive);
    // Exercise the inflater too, using repeatable high-entropy bytes so the
    // compressed archive still crosses the input-chunk boundary.
    let seed = 1;
    const payload = Uint8Array.from({ length: 128 * 1024 }, () => {
        seed ^= seed << 13;
        seed ^= seed >>> 17;
        seed ^= seed << 5;
        return seed & 255;
    });
    const deflated = Buffer.from(zipSync({
        'first.bin': payload,
        'later.txt': strToU8('Later entry'),
    }));
    assert.ok(deflated.length > 64 * 1024);
    await assertMidExtractionAbort(deflated);
    assert.deepEqual(
        await extractFiles(deflated, () => true, {}, { ...quiet, abortSignal: new AbortController().signal }),
        await extractFiles(deflated, () => true, {}, quiet),
    );

    const detection = new AbortController();
    const detectionTimer = setTimeout(() => detection.abort(), 0);
    try {
        await assert.rejects(detectOfficeTypeFromZip(archive, {}, detection.signal), { name: 'AbortError' });
    } finally {
        clearTimeout(detectionTimer);
    }

    // This also proves the entry point checks cancellation before file I/O.
    await assert.rejects(OfficeParser.parseOffice('missing-cancelled-file.odt', {
        ...quiet, abortSignal: preAborted.signal,
    }), { name: 'AbortError' });

    const live = new AbortController();
    const config = { ...quiet, abortSignal: live.signal };
    const actual = await extractFiles(archive, () => true, {}, config);
    assert.deepEqual(actual, await extractFiles(archive, () => true, {}, quiet));
    assert.equal(getEventListeners(live.signal, 'abort').length, 0, 'success must remove its abort listener');
    assert.equal(await detectOfficeTypeFromZip(archive, {}, live.signal), 'odt');
    const parsed = await OfficeParser.parseOffice(archive, { ...config, fileType: 'odt' });
    assert.match(String((await parsed.to('text')).value), /Example text/);
    assert.deepEqual(await extractFiles(archive, () => false, {}, config), []);
    await assert.rejects(extractFiles(Buffer.from('invalid'), () => true, {}, config));
    await assert.rejects(extractFiles(archive.subarray(0, archive.length - 22), () => true, {}, config));
    await assert.rejects(extractFiles(archive, () => true, { maxUncompressedBytes: 1 }, config));
    await assert.rejects(extractFiles(archive, () => true, { maxZipEntries: 1 }, config));
    assert.equal(getEventListeners(live.signal, 'abort').length, 0, 'failures must remove their abort listeners');
    // A late abort after success must not affect another extraction.
    live.abort();
    assert.equal((await extractFiles(archive, () => true, {}, quiet)).length, 3);
}

async function assertMidExtractionAbort(input: Buffer): Promise<void> {
    const controller = new AbortController();
    const visited: string[] = [];
    let abortTimer: ReturnType<typeof setTimeout> | undefined;
    try {
        await assert.rejects(extractFiles(input, name => {
            visited.push(name);
            if (name === 'first.bin') abortTimer = setTimeout(() => controller.abort(), 0);
            return true;
        }, {}, { ...quiet, abortSignal: controller.signal }), { name: 'AbortError' });
        assert.deepEqual(visited, ['first.bin'], 'cancellation must stop input before later entries');
        assert.equal(getEventListeners(controller.signal, 'abort').length, 0, 'abort listener must be removed');
        await new Promise(resolve => setTimeout(resolve, 10));
        assert.deepEqual(visited, ['first.bin'], 'no scheduled work may resume after rejection');
    } finally {
        clearTimeout(abortTimer);
    }

}
