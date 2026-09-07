/**
 * Minimal reader for the Compound File Binary (CFB / OLE2) container.
 *
 * Encrypted OOXML documents (a password-protected .docx/.xlsx/.pptx) are not zip files: they are CFB
 * containers holding two streams, `EncryptionInfo` (the descriptor) and `EncryptedPackage` (the
 * encrypted zip). We only ever need to pull named streams out of such a container, so this is a
 * read-only, extract-by-name implementation of just enough of [MS-CFB], not a general CFB library.
 *
 * @module crypto/cfb
 */

/** The 8-byte signature every CFB file begins with (`D0 CF 11 E0 A1 B1 1A E1`). */
import { DecryptionError } from './wrongPassword.js';

export const CFB_SIGNATURE = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;

/** True when the buffer starts with the CFB signature. */
export function isCfb(buf: Uint8Array): boolean {
    if (buf.length < 8) return false;
    for (let i = 0; i < 8; i++) if (buf[i] !== CFB_SIGNATURE[i]) return false;
    return true;
}

interface DirEntry {
    name: string;
    type: number;        // 0 empty, 1 storage, 2 stream, 5 root
    startSector: number;
    size: number;
}

/**
 * Parses a CFB container and extracts named streams. Constructed from the raw file bytes; use
 * {@link CfbContainer.read} to pull a stream out by name.
 */
export class CfbContainer {
    private view: DataView;
    private buf: Uint8Array;
    private sectorSize: number;
    private miniSectorSize: number;
    private miniCutoff: number;
    private fat: number[] = [];
    private miniFat: number[] = [];
    private dir: DirEntry[] = [];
    private miniStream: Uint8Array = new Uint8Array(0);
    /** A file of N bytes has at most ceil(N/sectorSize) sectors; every chain/table is bounded by this. */
    private maxSectors: number;

    constructor(buf: Uint8Array) {
        if (!isCfb(buf)) throw new DecryptionError('not a CFB container');
        this.buf = buf;
        this.view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
        const majorVersion = this.view.getUint16(26, true);
        this.sectorSize = 1 << this.view.getUint16(30, true);   // sector shift (9 -> 512, 12 -> 4096)
        this.miniSectorSize = 1 << this.view.getUint16(32, true); // mini sector shift (6 -> 64)
        this.miniCutoff = this.view.getUint32(56, true);
        // v3 files use 512-byte sectors and ignore anything past the 512-byte header; v4 use 4096.
        if (majorVersion !== 3 && majorVersion !== 4) throw new DecryptionError(`unsupported CFB major version ${majorVersion}`);
        if (this.sectorSize < 512 || this.sectorSize > 1 << 20) throw new DecryptionError('unsupported CFB sector size');
        this.maxSectors = Math.ceil(buf.length / this.sectorSize) + 1;

        this.readFat();
        this.readDirectory();
        this.readMiniStream();
    }

    /** Offset of the start of a numbered sector (sector 0 begins right after the header). */
    private sectorOffset(sector: number): number {
        return (sector + 1) * this.sectorSize;
    }

    private u32(offset: number): number {
        return this.view.getUint32(offset, true);
    }

    /** Reads the DIFAT (the list of FAT sectors) and then the FAT (the sector-chain table) itself. */
    private readFat(): void {
        const difat: number[] = [];
        // First 109 FAT sector locations live in the header at offset 76.
        for (let i = 0; i < 109; i++) {
            const s = this.u32(76 + i * 4);
            if (s === FREESECT || s === ENDOFCHAIN) break;
            difat.push(s);
        }
        // Remaining FAT sector locations, if any, are chained through DIFAT sectors. A malformed
        // (self-referential or over-long) DIFAT chain is bounded two ways: a visited-set rejects
        // cycles, and the count cannot exceed the file's sector count. Both keep a tiny hostile file
        // from growing `difat`/`fat` without bound.
        let difatSector = this.u32(68);
        const perDifat = this.sectorSize / 4 - 1;
        const seenDifat = new Set<number>();
        while (difatSector !== ENDOFCHAIN && difatSector !== FREESECT) {
            if (seenDifat.has(difatSector) || seenDifat.size > this.maxSectors || difat.length > this.maxSectors) break;
            seenDifat.add(difatSector);
            const base = this.sectorOffset(difatSector);
            for (let i = 0; i < perDifat; i++) {
                const s = this.u32(base + i * 4);
                if (s !== FREESECT && s !== ENDOFCHAIN) difat.push(s);
            }
            difatSector = this.u32(base + perDifat * 4);
        }
        for (const fatSector of difat) {
            if (this.fat.length > this.maxSectors * (this.sectorSize / 4)) break;
            const base = this.sectorOffset(fatSector);
            for (let i = 0; i < this.sectorSize / 4; i++) this.fat.push(this.u32(base + i * 4));
        }
    }

    /** Walks a FAT sector chain and concatenates the sector bytes. Rejects cycles via a visited set. */
    private readChain(start: number): Uint8Array {
        const parts: Uint8Array[] = [];
        let sector = start;
        const seen = new Set<number>();
        while (sector !== ENDOFCHAIN && sector !== FREESECT) {
            if (seen.has(sector) || seen.size > this.maxSectors) break;
            seen.add(sector);
            const off = this.sectorOffset(sector);
            parts.push(this.buf.subarray(off, off + this.sectorSize));
            sector = this.fat[sector];
            if (sector === undefined) break;
        }
        return concat(parts);
    }

    private readDirectory(): void {
        const dirBytes = this.readChain(this.u32(48)); // first directory sector
        const view = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
        for (let off = 0; off + 128 <= dirBytes.length; off += 128) {
            // The name field is a fixed 64-byte UTF-16 buffer; clamp the declared length to it so a
            // bogus value cannot read into the next entry.
            const nameLen = Math.min(view.getUint16(off + 64, true), 64);
            const type = view.getUint8(off + 66);
            if (type === 0) continue; // empty slot
            let name = '';
            for (let i = 0; i + 1 < nameLen; i += 2) {
                const code = view.getUint16(off + i, true);
                if (code === 0) break;
                name += String.fromCharCode(code);
            }
            const startSector = view.getUint32(off + 116, true);
            // Size is a 64-bit field; a stream over 4 GiB is not something we will ever decrypt.
            const size = view.getUint32(off + 120, true);
            this.dir.push({ name, type, startSector, size });
        }
    }

    /** The mini-stream (holder of all sub-cutoff streams) is the root entry's own stream. */
    private readMiniStream(): void {
        const root = this.dir.find(e => e.type === 5);
        if (!root) return;
        this.miniStream = this.readChain(root.startSector).subarray(0, root.size);
        // The mini FAT chains mini-sectors within the mini-stream, exactly as the FAT chains sectors.
        const miniFatBytes = this.readChain(this.u32(60)); // first mini FAT sector
        const view = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength);
        for (let i = 0; i + 4 <= miniFatBytes.length; i += 4) this.miniFat.push(view.getUint32(i, true));
    }

    private readMiniChain(start: number, size: number): Uint8Array {
        const parts: Uint8Array[] = [];
        let sector = start;
        const seen = new Set<number>();
        const maxMini = this.miniStream.length / this.miniSectorSize + 1;
        while (sector !== ENDOFCHAIN && sector !== FREESECT) {
            if (seen.has(sector) || seen.size > maxMini) break;
            seen.add(sector);
            const off = sector * this.miniSectorSize;
            parts.push(this.miniStream.subarray(off, off + this.miniSectorSize));
            sector = this.miniFat[sector];
            if (sector === undefined) break;
        }
        return concat(parts).subarray(0, size);
    }

    /** Whether a stream of the given name exists. */
    has(name: string): boolean {
        return this.dir.some(e => e.type === 2 && e.name === name);
    }

    /** Returns the bytes of the named stream, or `undefined` when it is absent. */
    read(name: string): Uint8Array | undefined {
        const entry = this.dir.find(e => e.type === 2 && e.name === name);
        if (!entry) return undefined;
        if (entry.size < this.miniCutoff) return this.readMiniChain(entry.startSector, entry.size);
        return this.readChain(entry.startSector).subarray(0, entry.size);
    }
}

function concat(parts: Uint8Array[]): Uint8Array {
    let total = 0;
    for (const p of parts) total += p.length;
    const out = new Uint8Array(total);
    let off = 0;
    for (const p of parts) { out.set(p, off); off += p.length; }
    return out;
}
