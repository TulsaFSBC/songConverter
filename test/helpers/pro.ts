import { getPresentationType } from '../../src/lib/pro/proto.js';

type Json = null | boolean | number | string | Json[] | { [key: string]: Json };

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Decodes a .pro file to a plain object (enums as names, bytes as base64). */
export function decodePresentation(data: Uint8Array): Record<string, Json> {
    const Presentation = getPresentationType();
    return Presentation.toObject(Presentation.decode(data), { enums: String, longs: Number, bytes: String }) as Record<string, Json>;
}

/** The RTF of each slide's text box. */
export function slideRtfs(presentation: Record<string, Json>): string[] {
    const cues = (presentation.cues ?? []) as Record<string, Json>[];
    return cues.map((cue) => {
        const action = (cue.actions as Record<string, Json>[])[0] as Record<string, Json>;
        const slide = (action.slide as Record<string, Json>).presentation as Record<string, Json>;
        const element = ((slide.base_slide as Record<string, Json>).elements as Record<string, Json>[])[0] as Record<string, Json>;
        const text = (element.element as Record<string, Json>).text as Record<string, Json>;
        return Buffer.from(text.rtf_data as string, 'base64').toString('latin1');
    });
}

/**
 * Minimal RTF reader for the RTF this app (and ProPresenter) writes: returns the
 * paragraphs' text, decoding \uN escapes, \tab and escaped symbols, and skipping
 * {\fonttbl}/{\colortbl}/{\*...} groups.
 */
export function rtfToLines(rtf: string): string[] {
    const lines: string[] = [];
    let current = '';
    let depth = 0;
    let skipDepth = -1;
    let pendingSurrogate: number | undefined;
    let i = 0;
    const emit = (code: number) => {
        if (code >= 0xd800 && code <= 0xdbff) {
            pendingSurrogate = code;
            return;
        }
        if (pendingSurrogate !== undefined) {
            current += String.fromCharCode(pendingSurrogate, code);
            pendingSurrogate = undefined;
            return;
        }
        current += String.fromCharCode(code);
    };
    while (i < rtf.length) {
        const ch = rtf[i] as string;
        if (ch === '{') {
            depth++;
            if (skipDepth < 0 && (rtf.startsWith('{\\fonttbl', i) || rtf.startsWith('{\\colortbl', i) || rtf.startsWith('{\\*', i))) {
                skipDepth = depth;
            }
            i++;
        } else if (ch === '}') {
            if (depth === skipDepth) {
                skipDepth = -1;
            }
            depth--;
            i++;
        } else if (ch === '\\') {
            const next = rtf[i + 1] ?? '';
            if (next === '\\' || next === '{' || next === '}') {
                if (skipDepth < 0) {
                    current += next;
                }
                i += 2;
                continue;
            }
            const match = /^\\([a-zA-Z]+)(-?\d+)? ?/.exec(rtf.slice(i));
            if (!match) {
                i += 2;
                continue;
            }
            i += match[0].length;
            if (skipDepth >= 0) {
                continue;
            }
            const [, word, param] = match;
            if (word === 'par') {
                lines.push(current);
                current = '';
            } else if (word === 'tab') {
                current += '\t';
            } else if (word === 'u' && param !== undefined) {
                const value = Number(param);
                emit(value < 0 ? value + 0x10000 : value);
                // Skip the one fallback character (\uc1).
                if (rtf[i] === '?') {
                    i++;
                }
            }
        } else {
            if (skipDepth < 0 && depth > 0) {
                current += ch;
            }
            i++;
        }
    }
    lines.push(current);
    return lines;
}

/**
 * Removes everything that legitimately differs between two presentations with the same
 * layout (ids, timestamps, names, text, number of slides), leaving only the structure.
 */
export function structureOf(value: Json, key = ''): Json {
    if (Array.isArray(value)) {
        if (key === 'cues') {
            return value.length > 0 ? [structureOf(value[0] as Json)] : [];
        }
        if (key === 'cue_identifiers') {
            return '<cue ids>';
        }
        if (key === 'custom_attributes') {
            return value.length > 0 ? [structureOf(value[0] as Json)] : [];
        }
        return value.map((item) => structureOf(item));
    }
    if (value !== null && typeof value === 'object') {
        const out: Record<string, Json> = {};
        for (const k of Object.keys(value).sort()) {
            if (k === 'range') {
                out[k] = '<range>';
            } else if (k === 'rtf_data' && key === 'text') {
                out[k] = '<slide rtf>';
            } else if (k === 'seconds') {
                out[k] = 0;
            } else if (k === 'name' && key === '') {
                out[k] = '<presentation name>';
            } else {
                out[k] = structureOf(value[k] as Json, k);
            }
        }
        return out;
    }
    if (typeof value === 'string' && UUID.test(value)) {
        return '<uuid>';
    }
    return value;
}
