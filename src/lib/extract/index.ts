import { extractPpt } from './ppt.js';
import { extractPptx } from './pptx.js';
import { UnsupportedFileError, type SlideText } from './slides.js';

export { UnsupportedFileError, type SlideText } from './slides.js';

const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** File name extensions treated as PowerPoint presentations. */
export const POWERPOINT_EXTENSIONS = ['.pptx', '.pptm', '.ppsx', '.ppsm', '.ppt', '.pps'];

export function isPowerPointFileName(fileName: string): boolean {
    const lower = fileName.toLowerCase();
    // Office lock/owner files ("~$Song.pptx") are temporary and not real presentations.
    return !lower.startsWith('~$') && POWERPOINT_EXTENSIONS.some((ext) => lower.endsWith(ext));
}

/**
 * Extracts slide text from a PowerPoint file. The format is detected from the
 * file contents, so a .ppt saved with a .pptx name (or vice versa) still works.
 */
export async function extractSlides(data: Uint8Array): Promise<SlideText[]> {
    try {
        if (startsWith(data, ZIP_MAGIC)) {
            return await extractPptx(data);
        }
        if (startsWith(data, OLE_MAGIC)) {
            return extractPpt(Buffer.from(data.buffer, data.byteOffset, data.byteLength));
        }
    } catch (err) {
        // Extraction only reads the bytes in memory, so any other failure means the file
        // itself is damaged; retrying would fail the same way.
        if (err instanceof UnsupportedFileError) {
            throw err;
        }
        throw new UnsupportedFileError(`PowerPoint file is damaged and cannot be read: ${(err as Error).message}`);
    }
    throw new UnsupportedFileError('File is not a PowerPoint presentation (.pptx or .ppt).');
}

function startsWith(data: Uint8Array, magic: number[]): boolean {
    return data.length >= magic.length && magic.every((byte, i) => data[i] === byte);
}
