/**
 * OCR (Optical Character Recognition) Utilities
 * 
 * This module provides functions for extracting text from images using Tesseract.js.
 * Used when `config.ocr` is enabled to extract text from embedded images in documents.
 * 
 * Includes a worker pool via OcrSchedulerManager to improve performance when 
 * processing multiple images.
 * 
 * @module ocrUtils
 */

import { OcrConfig } from '../types.js';
import { isBrowser } from './envUtils.js';
import { getAbortError } from './errorUtils.js';
import { median } from './numberUtils.js';

/**
 * Internal interface for tracking jobs in the scheduler queue.
 */
interface OcrJob {
    image: any;
    config: OcrConfig;
    resolve: (text: string) => void;
    reject: (err: any) => void;
    startTime: number;
    timeoutMs: number;
    isFinished?: boolean;
}

/**
 * Internal interface for tracking workers within the pool.
 */
interface ManagedWorker {
    worker: any;
    language: string;
    lastUsed: number;
    isBusy: boolean;
    activeJob?: OcrJob;
}

/**
 * Reconstructs the two-dimensional page layout of a Tesseract result from its per-word bounding
 * boxes, so recognized text reads spatially (columns line up, right-hand text stays right) instead
 * of as one flat reading-order string. Each visual line's words are placed at a character column
 * derived from their x position (using the page's median glyph width), the block is left-normalized
 * so there is no large leading indent, lines keep their top-to-bottom order, and a wide vertical gap
 * becomes a blank line. Falls back to the flat `page.text` when no word geometry is available.
 */
function layoutOcrText(page: any): string {
    const blocks = page?.blocks;
    const flat: string = typeof page?.text === 'string' ? page.text : '';
    if (!Array.isArray(blocks) || !blocks.length) return flat;

    interface OcrWord { text: string; x0: number; x1: number; }
    interface OcrLine { words: OcrWord[]; y0: number; }
    const lines: OcrLine[] = [];
    for (const block of blocks) {
        for (const para of block?.paragraphs || []) {
            for (const line of para?.lines || []) {
                const words: OcrWord[] = [];
                for (const w of line?.words || []) {
                    const text = (w?.text || '').trim();
                    if (text) words.push({ text, x0: w.bbox?.x0 ?? 0, x1: w.bbox?.x1 ?? 0 });
                }
                if (words.length) lines.push({ words, y0: line.bbox?.y0 ?? line?.words?.[0]?.bbox?.y0 ?? 0 });
            }
        }
    }
    if (!lines.length) return flat;

    // Character width unit: median per-glyph width across all words.
    const glyphWidths: number[] = [];
    for (const l of lines) for (const w of l.words) if (w.text.length) glyphWidths.push((w.x1 - w.x0) / w.text.length);
    const charWidth = median(glyphWidths.filter(v => v > 0)) || 8;

    // Left-normalize so the leftmost word sits at column 0 (no giant leading indent).
    let minX = Infinity;
    for (const l of lines) for (const w of l.words) minX = Math.min(minX, w.x0);
    if (!Number.isFinite(minX)) minX = 0;

    lines.sort((a, b) => a.y0 - b.y0);
    const pitches: number[] = [];
    for (let i = 1; i < lines.length; i++) pitches.push(lines[i].y0 - lines[i - 1].y0);
    const linePitch = median(pitches.filter(v => v > 0)) || charWidth * 2;

    const out: string[] = [];
    let prevY: number | null = null;
    for (const l of lines) {
        if (prevY !== null && l.y0 - prevY > 1.8 * linePitch) out.push(''); // blank line for a big vertical gap
        prevY = l.y0;
        let s = '';
        let col = 0;
        for (const w of l.words) {
            const target = Math.max(col, Math.round((w.x0 - minX) / charWidth));
            if (target > col) { s += ' '.repeat(target - col); col = target; }
            else if (s.length && !s.endsWith(' ')) { s += ' '; col += 1; } // always keep words apart
            s += w.text;
            col += w.text.length;
        }
        out.push(s.replace(/\s+$/, ''));
    }
    return out.join('\n');
}

/**
 * Wraps a promise in a timeout.
 *
 * @param promise - The promise to wrap
 * @param ms - Timeout duration in milliseconds
 * @param errMsg - Error message to throw if timeout occurs
 * @returns The wrapped promise
 */
function withTimeout<T>(promise: Promise<T>, ms: number, errMsg: string): Promise<T> {
    let id: any;
    const timeout = new Promise<never>((_, reject) => {
        id = setTimeout(() => {
            reject(new Error(errMsg));
        }, ms);
    });
    return Promise.race([promise, timeout]).then(
        (res) => { clearTimeout(id); return res; },
        (err) => { clearTimeout(id); throw err; }
    );
}

/**
 * Manages a pool of Tesseract workers with "Smart Affinity".
 * 
 * Instead of a simple scheduler, this manager allows workers to persist with 
 * a specific language affinity. If a new language is requested and the pool 
 * is at capacity, it re-initializes the Least Recently Used (LRU) idle worker 
 * rather than resetting the entire pool.
 * 
 * Implements lazy loading of tesseract.js to ensure no background processes
 * are spawned unless OCR is explicitly used.
 */
class OcrSchedulerManager {
    private static instance: OcrSchedulerManager;
    private pool: ManagedWorker[] = [];
    private queue: OcrJob[] = [];
    private readonly MAX_WORKERS: number = 4;
    private idleTimeout: number = 10000; // 10s default
    private timeoutId: NodeJS.Timeout | null = null;
    private isProcessing: boolean = false;

    private constructor() { }

    /**
     * Returns the singleton instance of the manager.
     */
    public static getInstance(): OcrSchedulerManager {
        if (!OcrSchedulerManager.instance) {
            OcrSchedulerManager.instance = new OcrSchedulerManager();
        }
        return OcrSchedulerManager.instance;
    }

    /**
     * Checks if the singleton instance has been initialized.
     */
    public static hasInstance(): boolean {
        return !!OcrSchedulerManager.instance;
    }

    /**
     * Resets the inactivity timer. If the timer reaches its duration, 
     * all workers are terminated automatically.
     */
    private resetIdleTimer(): void {
        if (this.timeoutId) {
            clearTimeout(this.timeoutId);
        }

        if (this.idleTimeout > 0) {
            this.timeoutId = setTimeout(async () => {
                await this.terminate();
            }, this.idleTimeout);
        }
    }

    /**
     * Performs OCR on an image using the smart worker pool.
     * 
     * @param image - Image data (Buffer, string path, or Blob)
     * @param config - OCR configuration (language, custom paths, timeouts, signal)
     * @returns Recognized text
     */
    public async recognize(image: any, config?: OcrConfig): Promise<string> {
        const signal = config?.abortSignal;
        if (signal?.aborted) {
            return Promise.reject(getAbortError());
        }

        return new Promise((resolve, reject) => {
            // Update idle timeout if provided.
            const effectiveAutoTerminate = config?.timeout?.autoTerminate;
            if (effectiveAutoTerminate !== undefined) {
                this.idleTimeout = effectiveAutoTerminate;
            }

            // Reset the inactivity timer every time a new job is requested
            this.resetIdleTimer();

            let abortListener: (() => void) | null = null;
            let finished = false;
            let job: OcrJob;

            const cleanResolve = (val: string) => {
                if (finished) return;
                finished = true;
                if (job) job.isFinished = true;
                if (abortListener && signal) {
                    signal.removeEventListener('abort', abortListener);
                }
                resolve(val);
            };

            const cleanReject = (err: any) => {
                if (finished) return;
                finished = true;
                if (job) job.isFinished = true;
                if (abortListener && signal) {
                    signal.removeEventListener('abort', abortListener);
                }
                reject(err);
            };

            // Priority: timeout.recognition (new) > 30 s default.
            const recogTimeout = config?.timeout?.recognition ?? 30000;

            // Create job
            job = {
                image,
                config: config || {},
                resolve: cleanResolve,
                reject: cleanReject,
                startTime: Date.now(),
                timeoutMs: recogTimeout
            };

            if (signal) {
                abortListener = () => {
                    if (finished) return;
                    
                    const err = getAbortError();
                    cleanReject(err);

                    // 1. Remove job from queue if it hasn't run yet
                    const idx = this.queue.indexOf(job);
                    if (idx !== -1) {
                        this.queue.splice(idx, 1);
                    }

                    // 2. Find if any worker is currently running this job and terminate/remove it
                    const workerIndex = this.pool.findIndex(mw => mw.activeJob === job);
                    if (workerIndex !== -1) {
                        const managedWorker = this.pool[workerIndex];
                        // Remove from pool immediately to prevent reuse
                        this.pool.splice(workerIndex, 1);
                        // Terminate the worker process
                        try {
                            managedWorker.worker.terminate();
                        } catch (e) {}
                        // Trigger queue processing for subsequent tasks
                        this.processQueue();
                    }
                };
                signal.addEventListener('abort', abortListener);
            }

            // Add job to queue and trigger processing
            this.queue.push(job);
            this.processQueue();
        });
    }

    /**
     * Attempts to process the next job in the queue using an available worker.
     * Designed to be race-free and support concurrent/parallel job execution.
     */
    private async processQueue(): Promise<void> {
        if (this.isProcessing) return;
        this.isProcessing = true;

        try {
            while (this.queue.length > 0) {
                const nextJob = this.queue[0];
                if (nextJob.isFinished) {
                    this.queue.shift();
                    continue;
                }
                const requestedLanguage = nextJob.config.language || 'eng';

                // 1. Find an idle worker with the EXACT language affinity
                let managed = this.pool.find(mw => !mw.isBusy && mw.language === requestedLanguage);

                // 2. If not found and we have room, create a new worker
                if (!managed && this.pool.length < this.MAX_WORKERS) {
                    const job = this.queue.shift();
                    if (!job) continue;

                    this.createAndRunWorker(job, requestedLanguage);
                    continue;
                }

                // 3. If still not found and we are at capacity, find the LRU idle worker and re-initialize it
                if (!managed) {
                    const idleWorkers = this.pool.filter(mw => !mw.isBusy);
                    if (idleWorkers.length > 0) {
                        const job = this.queue.shift();
                        if (!job) continue;

                        managed = idleWorkers.reduce((prev, curr) => (prev.lastUsed < curr.lastUsed ? prev : curr));
                        this.reinitializeAndRunWorker(managed, job, requestedLanguage);
                        continue;
                    }
                }

                // 4. If we have a worker ready, execute the job
                if (managed) {
                    const job = this.queue.shift();
                    if (!job) continue;

                    this.runWorker(managed, job);
                    continue;
                }

                // No workers can be allocated right now (all busy and pool at capacity). Break work loop.
                break;
            }
        } finally {
            this.isProcessing = false;
        }
    }

    /**
     * Helper to dynamically instantiate a Tesseract worker, register it to the pool, and run the job.
     */
    private async createAndRunWorker(job: OcrJob, requestedLanguage: string): Promise<void> {
        // Priority: timeout.workerLoad (new) > 60 s default.
        const loadTimeout = job.config.timeout?.workerLoad ?? 60000;
        let managed: ManagedWorker | null = null;

        try {
            const { createWorker } = await import('tesseract.js');
            const options: any = { logger: () => { } };
            if (job.config.workerPath) options.workerPath = job.config.workerPath;
            if (job.config.corePath) options.corePath = job.config.corePath;
            if (job.config.langPath) options.langPath = job.config.langPath;

            const workerPromise = createWorker(requestedLanguage, 1, options);

            // To prevent dangling worker threads on timeout or abort, we register a post-resolution hook
            // that terminates the worker if the promise finishes after the timeout has fired or the job is finished.
            let hasTimedOutOrAborted = false;
            workerPromise.then(
                async (worker) => {
                    if (hasTimedOutOrAborted || job.isFinished) {
                        try {
                            await worker.terminate();
                        } catch (e) {}
                    }
                },
                () => {}
            );

            const worker = loadTimeout > 0
                ? await withTimeout(workerPromise, loadTimeout, `OCR worker initialization timed out after ${loadTimeout}ms`).catch(err => {
                    hasTimedOutOrAborted = true;
                    throw err;
                })
                : await workerPromise;

            // If the job finished/aborted while loading, clean up the worker and skip execution.
            if (job.isFinished) {
                hasTimedOutOrAborted = true;
                try {
                    await worker.terminate();
                } catch (e) {}
                this.processQueue();
                return;
            }

            managed = {
                worker,
                language: requestedLanguage,
                lastUsed: Date.now(),
                isBusy: false
            };
            this.pool.push(managed);

            await this.runWorker(managed, job);
        } catch (err) {
            job.reject(err);
            this.processQueue();
        }
    }

    /**
     * Helper to reinitialize an existing idle worker with a different language affinity and run the job.
     */
    private async reinitializeAndRunWorker(managed: ManagedWorker, job: OcrJob, requestedLanguage: string): Promise<void> {
        // Priority: timeout.workerLoad (new) > 60 s default.
        const loadTimeout = job.config.timeout?.workerLoad ?? 60000;
        managed.isBusy = true;
        managed.lastUsed = Date.now();
        managed.activeJob = job;

        try {
            const reinitPromise = managed.worker.reinitialize(requestedLanguage);
            if (loadTimeout > 0) {
                await withTimeout(reinitPromise, loadTimeout, `OCR worker re-initialization timed out after ${loadTimeout}ms`);
            } else {
                await reinitPromise;
            }
            managed.language = requestedLanguage;

            // If the job finished/aborted while reinitializing, clean up and skip execution.
            if (job.isFinished) {
                const index = this.pool.indexOf(managed);
                if (index !== -1) {
                    this.pool.splice(index, 1);
                }
                try {
                    await managed.worker.terminate();
                } catch (e) {}
                this.processQueue();
                return;
            }

            await this.runWorker(managed, job);
        } catch (err: any) {
            // Re-initialization failed/timed out, remove worker from pool and terminate
            const index = this.pool.indexOf(managed);
            if (index !== -1) {
                this.pool.splice(index, 1);
            }
            try {
                await managed.worker.terminate();
            } catch (e) {}

            job.reject(err);
            this.processQueue();
        }
    }

    /**
     * Helper to execute OCR text recognition on the worker and return the results.
     */
    private async runWorker(managed: ManagedWorker, job: OcrJob): Promise<void> {
        // Priority: timeout.recognition (new) > 30 s default.
        const recogTimeout = job.config.timeout?.recognition ?? 30000;
        managed.isBusy = true;
        managed.lastUsed = Date.now();
        managed.activeJob = job;

        try {
            // When preserving layout (default), ask Tesseract for the block/word tree so we can
            // rebuild the 2-D page layout from per-word boxes; otherwise the flat text is enough.
            const wantLayout = job.config.preserveLayout !== false;
            const recognizePromise = wantLayout
                ? managed.worker.recognize(job.image, {}, { text: true, blocks: true })
                : managed.worker.recognize(job.image);
            const { data } = recogTimeout > 0
                ? await withTimeout(recognizePromise, recogTimeout, `OCR recognition timed out after ${recogTimeout}ms`)
                : await recognizePromise;

            job.resolve(wantLayout ? layoutOcrText(data) : (data.text || ''));
        } catch (err: any) {
            // If it timed out, terminate and remove worker to avoid reusing a stuck process
            if (err.message?.includes('timed out')) {
                const index = this.pool.indexOf(managed);
                if (index !== -1) {
                    this.pool.splice(index, 1);
                }
                try {
                    await managed.worker.terminate();
                } catch (e) {}
            }
            job.reject(err);
        } finally {
            if (this.pool.includes(managed)) {
                managed.isBusy = false;
                managed.activeJob = undefined;
                managed.lastUsed = Date.now();
            }
            this.processQueue();
        }
    }

    /**
     * Terminates all workers in the pool and resets the state.
     */
    public async terminate(): Promise<void> {
        if (this.timeoutId) {
            clearTimeout(this.timeoutId);
            this.timeoutId = null;
        }

        const workersToTerminate = this.pool.map(mw => mw.worker.terminate());
        await Promise.all(workersToTerminate);
        this.pool = [];
    }
}

/**
 * Performs Optical Character Recognition (OCR) on an image to extract text.
 * 
 * Uses Tesseract.js to recognize text in the provided image buffer.
 * This is useful for extracting text from screenshots, scanned documents,
 * charts with labels, or any image containing text.
 * 
 * This function uses a shared worker pool to minimize initialization overhead.
 * 
 * @param image - The image data as a Buffer, file path, or Blob
 * @param config - Optional configuration for language and custom worker paths
 * @returns A promise that resolves to the recognized text as a string
 * @throws {Error} If the image cannot be processed or Tesseract initialization fails
 * 
 * @example
 * ```typescript
 * // Extract text from an English image
 * const text = await performOcr(imageBuffer, { language: 'eng' });
 * ```
 * 
 * @see https://github.com/naptha/tesseract.js for supported languages and options
 */
export const performOcr = async (image: Buffer | string, config?: OcrConfig): Promise<string> => {
    // Prepare image data
    let inputImage: any = image;

    // In browser environment, convert Buffer to Blob for better compatibility
    // @ts-ignore
    if (isBrowser && typeof Blob !== 'undefined' && Buffer.isBuffer(image)) {
        inputImage = new Blob([image as any], { type: 'image/bmp' });
    }

    return await OcrSchedulerManager.getInstance().recognize(inputImage, config);
};

/**
 * Terminates all OCR workers and cleans up resources.
 * 
 * Should be called when the application is shutting down or OCR is no longer needed
 * to prevent memory leaks and dangling worker processes.
 */
export const terminateOcr = async (): Promise<void> => {
    if (OcrSchedulerManager.hasInstance()) {
        await OcrSchedulerManager.getInstance().terminate();
    }
};
