// Builds the RTF that ProPresenter 7 stores for a slide's text box. The header and
// per-paragraph formatting mirror what ProPresenter 7.16 itself writes.

export interface RtfStyle {
    /** PostScript font name, e.g. "TimesNewRomanPSMT" or "ArialMT". */
    fontName: string;
    /** Font size in points. */
    fontSize: number;
}

/** Escapes text for RTF. Anything outside printable ASCII becomes a \uN escape, so no code page is involved. */
export function escapeRtfText(text: string): string {
    let out = '';
    for (let i = 0; i < text.length; i++) {
        const code = text.charCodeAt(i);
        if (code === 0x5c || code === 0x7b || code === 0x7d) {
            // \ { }
            out += '\\' + text[i];
        } else if (code === 0x09) {
            out += '\\tab ';
        } else if (code >= 0x20 && code < 0x7f) {
            out += text[i];
        } else if (code >= 0x80) {
            // RTF \uN takes a signed 16-bit value; characters outside the BMP are written
            // as their two UTF-16 surrogates. "?" is the fallback for readers without Unicode (\uc1).
            out += `\\u${code > 0x7fff ? code - 0x10000 : code} ?`;
        }
        // Remaining control characters are dropped.
    }
    return out;
}

function escapeFontName(name: string): string {
    return name.replace(/[\\{};]/g, '');
}

/** Returns the RTF document for a list of lines, one paragraph per line, centered and white. */
export function buildSlideRtf(lines: string[], style: RtfStyle): string {
    const halfPoints = Math.round(style.fontSize * 2);
    const header =
        '{\\rtf0\\ansi\\ansicpg1252' +
        `{\\fonttbl\\f0\\fnil ${escapeFontName(style.fontName)};}` +
        '{\\colortbl;\\red255\\green255\\blue255;\\red255\\green255\\blue255;}' +
        '{\\*\\expandedcolortbl;\\csgenericrgb\\c100000\\c100000\\c100000\\c100000;\\csgenericrgb\\c100000\\c100000\\c100000\\c0;}' +
        '{\\*\\listtable}{\\*\\listoverridetable}\\uc1\\paperw38400\\margl0\\margr0\\margt0\\margb0';
    const paragraph =
        '\\pard\\li0\\fi0\\ri0\\qc\\sb0\\sa0\\sl240\\slmult1\\slleading0\\f0\\b0\\i0\\ul0\\strike0' +
        `\\fs${halfPoints}\\expnd0\\expndtw0\\cf1\\strokewidth0\\strokec1\\nosupersub\\ulc0\\highlight2\\cb2`;
    const body = (lines.length > 0 ? lines : [''])
        .map((line) => {
            const text = escapeRtfText(line);
            // A space ends the preceding control word; it is not needed when the text starts with one.
            return paragraph + (text.startsWith('\\') ? '' : ' ') + text;
        })
        .join('\\par');
    return `${header}${body}}`;
}
