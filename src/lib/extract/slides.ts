/** Text content of one PowerPoint slide, as lyric lines in display order. */
export interface SlideText {
    lines: string[];
    /** Slide is hidden in PowerPoint (not shown during a slide show). */
    hidden: boolean;
}

/** Thrown when a file cannot be read as a PowerPoint presentation. Not retryable. */
export class UnsupportedFileError extends Error {
    override name = 'UnsupportedFileError';
}

// Zero-width and BOM characters that PowerPoint sometimes leaves in text and
// that would otherwise show up as invisible garbage in ProPresenter.
const INVISIBLE_CHARS = /[\u200B-\u200D\u2060\uFEFF]/g;
// C0 control characters other than tab; line breaks are handled before this runs.
const CONTROL_CHARS = /[\u0000-\u0008\u000A-\u001F\u007F]/g;

/** Cleans one line of text: tabs become spaces, invisible/control characters are removed, ends are trimmed. */
export function cleanLine(line: string): string {
    return line
        .replace(/\t/g, ' ')
        .replace(INVISIBLE_CHARS, '')
        .replace(CONTROL_CHARS, '')
        .trim();
}

/**
 * Normalizes a slide's lines: cleans each line, drops blank lines at the start
 * and end, and collapses runs of blank lines in the middle to a single blank line.
 */
export function normalizeLines(rawLines: string[]): string[] {
    const lines: string[] = [];
    for (const raw of rawLines) {
        const line = cleanLine(raw);
        if (line === '' && (lines.length === 0 || lines[lines.length - 1] === '')) {
            continue;
        }
        lines.push(line);
    }
    while (lines.length > 0 && lines[lines.length - 1] === '') {
        lines.pop();
    }
    return lines;
}

/** Splits text that uses PowerPoint line separators (CR, LF, CRLF, vertical tab, U+2028/2029) into lines. */
export function splitLines(text: string): string[] {
    return text.split(/\r\n|[\r\n\u000B\u2028\u2029]/);
}
