import CFB from 'cfb';
import { normalizeLines, splitLines, UnsupportedFileError, type SlideText } from './slides.js';

// Reader for the legacy binary PowerPoint 97-2003 format (.ppt / .pps), per [MS-PPT].
// The file is an OLE compound file; slide text lives in records inside the
// "PowerPoint Document" stream.

const RT = {
    Document: 0x03e8,
    Slide: 0x03ee,
    SlideShowSlideInfoAtom: 0x03f9,
    SlidePersistAtom: 0x03f3,
    UserEditAtom: 0x0ff5,
    CurrentUserAtom: 0x0ff6,
    PersistDirectoryAtom: 0x1772,
    SlideListWithText: 0x0ff0,
    TextHeaderAtom: 0x0f9f,
    TextCharsAtom: 0x0fa0,
    TextBytesAtom: 0x0fa8,
} as const;

// TextHeaderAtom.textType values that are not shown on the slide itself.
const TEXT_TYPE_NOTES = 2;
const TEXT_TYPE_NOT_USED = 3;
const SLIDE_LIST_INSTANCE_SLIDES = 0;
const HIDDEN_SLIDE_FLAG = 0x0004;
const MAX_DEPTH = 32;

interface RecordHeader {
    version: number;
    instance: number;
    type: number;
    length: number;
    /** Offset of the record body (just after the 8-byte header). */
    bodyOffset: number;
    /** Offset just past the end of the record. */
    end: number;
}

/** Extracts the text of every slide, in presentation order, from a .ppt/.pps file. */
export function extractPpt(data: Uint8Array): SlideText[] {
    let container: CFB.CFB$Container;
    try {
        container = CFB.read(data, { type: 'buffer' });
    } catch (err) {
        throw new UnsupportedFileError(`File is not a valid PowerPoint 97-2003 (.ppt) file: ${(err as Error).message}`);
    }

    if (CFB.find(container, 'EncryptedPackage')) {
        throw new UnsupportedFileError('PowerPoint file is password protected.');
    }
    const docEntry = CFB.find(container, 'PowerPoint Document');
    if (!docEntry?.content || docEntry.content.length === 0) {
        throw new UnsupportedFileError('File is an Office document but not a PowerPoint presentation.');
    }
    const doc = Buffer.from(docEntry.content as Uint8Array);
    const currentUser = CFB.find(container, 'Current User')?.content;

    const persist = currentUser ? readPersistDirectory(doc, Buffer.from(currentUser as Uint8Array)) : undefined;
    const documentOffset = persist?.documentOffset ?? findLastTopLevel(doc, RT.Document);
    if (documentOffset === undefined) {
        throw new UnsupportedFileError('PowerPoint file has no document record.');
    }

    const documentRecord = readHeader(doc, documentOffset);
    if (!documentRecord || documentRecord.type !== RT.Document) {
        throw new UnsupportedFileError('PowerPoint file has a corrupt document record.');
    }

    const slideList = childRecords(doc, documentRecord).find(
        (r) => r.type === RT.SlideListWithText && r.instance === SLIDE_LIST_INSTANCE_SLIDES,
    );
    if (!slideList) {
        throw new UnsupportedFileError('PowerPoint file does not contain any slides.');
    }

    // SlideListWithText holds, for each slide in order, a SlidePersistAtom followed by
    // the text of that slide's placeholders (title, body, ...).
    const slides: SlideText[] = [];
    let current: { persistId: number; lines: string[] } | undefined;
    let textType = -1;
    const finish = () => {
        if (current) {
            slides.push(buildSlide(doc, persist?.offsets, current.persistId, current.lines));
        }
    };
    for (const record of childRecords(doc, slideList)) {
        switch (record.type) {
            case RT.SlidePersistAtom:
                finish();
                current = { persistId: record.length >= 4 ? doc.readUInt32LE(record.bodyOffset) : 0, lines: [] };
                textType = -1;
                break;
            case RT.TextHeaderAtom:
                textType = record.length >= 4 ? doc.readUInt32LE(record.bodyOffset) : -1;
                break;
            case RT.TextCharsAtom:
            case RT.TextBytesAtom:
                if (current && textType !== TEXT_TYPE_NOTES && textType !== TEXT_TYPE_NOT_USED) {
                    current.lines.push(...splitLines(readText(doc, record)));
                }
                break;
            default:
                break;
        }
    }
    finish();
    return slides;
}

/** Adds text stored directly in the slide's shapes (text boxes that are not placeholders) and the hidden flag. */
function buildSlide(doc: Buffer, offsets: Map<number, number> | undefined, persistId: number, placeholderLines: string[]): SlideText {
    const lines = [...placeholderLines];
    let hidden = false;
    const slideOffset = offsets?.get(persistId);
    const slideRecord = slideOffset !== undefined ? readHeader(doc, slideOffset) : undefined;
    if (slideRecord && slideRecord.type === RT.Slide) {
        for (const record of childRecords(doc, slideRecord)) {
            if (record.type === RT.SlideShowSlideInfoAtom && record.length >= 12) {
                const flags = doc.readUInt16LE(record.bodyOffset + 10);
                hidden = (flags & HIDDEN_SLIDE_FLAG) !== 0;
            }
        }
        collectShapeText(doc, slideRecord, lines, 0);
    }
    return { lines: normalizeLines(lines), hidden };
}

/** Finds TextCharsAtom/TextBytesAtom records nested in the slide's drawing (OfficeArtClientTextbox). */
function collectShapeText(doc: Buffer, parent: RecordHeader, lines: string[], depth: number): void {
    if (depth > MAX_DEPTH) {
        return;
    }
    let textType = -1;
    for (const record of childRecords(doc, parent)) {
        if (record.type === RT.TextHeaderAtom && record.length >= 4) {
            textType = doc.readUInt32LE(record.bodyOffset);
        } else if (record.type === RT.TextCharsAtom || record.type === RT.TextBytesAtom) {
            if (textType !== TEXT_TYPE_NOTES && textType !== TEXT_TYPE_NOT_USED) {
                lines.push(...splitLines(readText(doc, record)));
            }
        } else if (record.version === 0xf) {
            collectShapeText(doc, record, lines, depth + 1);
        }
    }
}

function readText(doc: Buffer, record: RecordHeader): string {
    const body = doc.subarray(record.bodyOffset, record.end);
    // TextCharsAtom is UTF-16LE; TextBytesAtom stores the low byte of each UTF-16 code unit.
    return record.type === RT.TextCharsAtom ? body.toString('utf16le') : body.toString('latin1');
}

/**
 * Follows Current User -> UserEditAtom chain -> PersistDirectoryAtoms to map persist ids
 * to stream offsets. Newer edits (incremental "fast saves") override older ones.
 */
function readPersistDirectory(doc: Buffer, currentUser: Buffer): { documentOffset: number; offsets: Map<number, number> } | undefined {
    const userHeader = readHeader(currentUser, 0);
    if (!userHeader || userHeader.type !== RT.CurrentUserAtom || userHeader.length < 12) {
        return undefined;
    }
    let editOffset: number | undefined = currentUser.readUInt32LE(userHeader.bodyOffset + 8);
    const offsets = new Map<number, number>();
    let documentPersistId: number | undefined;
    const visited = new Set<number>();

    while (editOffset !== undefined && !visited.has(editOffset)) {
        visited.add(editOffset);
        const edit = readHeader(doc, editOffset);
        if (!edit || edit.type !== RT.UserEditAtom || edit.length < 20) {
            return undefined;
        }
        const offsetLastEdit = doc.readUInt32LE(edit.bodyOffset + 8);
        const offsetPersistDirectory = doc.readUInt32LE(edit.bodyOffset + 12);
        documentPersistId ??= doc.readUInt32LE(edit.bodyOffset + 16);

        const dir = readHeader(doc, offsetPersistDirectory);
        if (!dir || dir.type !== RT.PersistDirectoryAtom) {
            return undefined;
        }
        let pos = dir.bodyOffset;
        while (pos + 4 <= dir.end) {
            const entry = doc.readUInt32LE(pos);
            const startId = entry & 0xfffff;
            const count = entry >>> 20;
            pos += 4;
            for (let i = 0; i < count && pos + 4 <= dir.end; i++, pos += 4) {
                if (!offsets.has(startId + i)) {
                    offsets.set(startId + i, doc.readUInt32LE(pos));
                }
            }
        }
        editOffset = offsetLastEdit > 0 ? offsetLastEdit : undefined;
    }

    const documentOffset = documentPersistId !== undefined ? offsets.get(documentPersistId) : undefined;
    return documentOffset === undefined ? undefined : { documentOffset, offsets };
}

/** Fallback when the persist directory is unusable: the last top-level record of the given type. */
function findLastTopLevel(doc: Buffer, type: number): number | undefined {
    let found: number | undefined;
    let offset = 0;
    while (offset + 8 <= doc.length) {
        const header = readHeader(doc, offset);
        if (!header) {
            break;
        }
        if (header.type === type) {
            found = offset;
        }
        offset = header.end;
    }
    return found;
}

function readHeader(buf: Buffer, offset: number): RecordHeader | undefined {
    if (offset < 0 || offset + 8 > buf.length) {
        return undefined;
    }
    const verInst = buf.readUInt16LE(offset);
    const length = buf.readUInt32LE(offset + 4);
    const end = offset + 8 + length;
    if (end > buf.length) {
        return undefined;
    }
    return {
        version: verInst & 0xf,
        instance: verInst >>> 4,
        type: buf.readUInt16LE(offset + 2),
        length,
        bodyOffset: offset + 8,
        end,
    };
}

function childRecords(buf: Buffer, parent: RecordHeader): RecordHeader[] {
    const records: RecordHeader[] = [];
    let offset = parent.bodyOffset;
    while (offset + 8 <= parent.end) {
        const header = readHeader(buf, offset);
        if (!header || header.end > parent.end) {
            break;
        }
        records.push(header);
        offset = header.end;
    }
    return records;
}
