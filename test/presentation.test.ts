import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { describe, it } from 'node:test';
import { DEFAULT_CONVERSION_OPTIONS } from '../src/lib/convert.js';
import { buildPresentation, paragraphRanges } from '../src/lib/pro/presentation.js';
import { buildSlideRtf, escapeRtfText } from '../src/lib/pro/rtf.js';
import { decodePresentation, rtfToLines, slideRtfs, structureOf } from './helpers/pro.js';

const style = { fontName: 'TimesNewRomanPSMT', fontSize: 90 };
const PARAGRAPH =
    '\\pard\\li0\\fi0\\ri0\\qc\\sb0\\sa0\\sl240\\slmult1\\slleading0\\f0\\b0\\i0\\ul0\\strike0\\fs180\\expnd0\\expndtw0\\cf1\\strokewidth0\\strokec1\\nosupersub\\ulc0\\highlight2\\cb2';

describe('RTF', () => {
    it('escapes RTF special characters and writes non-ASCII as \\u escapes', () => {
        assert.equal(escapeRtfText('a\\b{c}d'), 'a\\\\b\\{c\\}d');
        assert.equal(escapeRtfText('Го'), '\\u1043 ?\\u1086 ?');
        assert.equal(escapeRtfText('don’t —'), 'don\\u8217 ?t \\u8212 ?');
        assert.equal(escapeRtfText('café'), 'caf\\u233 ?');
        // Above U+7FFF the value is written as a signed 16-bit number; emoji become two surrogates.
        assert.equal(escapeRtfText('！'), '\\u-255 ?');
        assert.equal(escapeRtfText('\u{1F64F}'), '\\u-10179 ?\\u-8625 ?');
        assert.equal(escapeRtfText('tab\there'), 'tab\\tab here');
        assert.equal(escapeRtfText('bell\u0007'), 'bell');
    });

    it('matches the paragraph layout ProPresenter 7.16 writes', () => {
        const rtf = buildSlideRtf(['//', 'Ты', '****'], style);
        assert.ok(rtf.startsWith('{\\rtf0\\ansi\\ansicpg1252{\\fonttbl\\f0\\fnil TimesNewRomanPSMT;}'));
        assert.ok(
            rtf.endsWith(`${PARAGRAPH} //\\par${PARAGRAPH}\\u1058 ?\\u1099 ?\\par${PARAGRAPH} ****}`),
            rtf,
        );
    });

    it('round-trips every kind of text', () => {
        const lines = ['Plain', 'Quotes "x" & <y> {z} \\ $1 $&', 'Всё — éè \u{1F64F}', '', '12 numbers first', 'tab\tinside'];
        assert.deepEqual(rtfToLines(buildSlideRtf(lines, style)), lines);
    });

    it('uses the configured font and size', () => {
        const rtf = buildSlideRtf(['x'], { fontName: 'ArialMT', fontSize: 72 });
        assert.match(rtf, /\\fnil ArialMT;/);
        assert.match(rtf, /\\fs144\\/);
    });
});

describe('ProPresenter presentation', () => {
    const slides = [['First line', 'Second line ’'], ['Господи'], [], ['Last']];
    const data = buildPresentation('My Song', slides, DEFAULT_CONVERSION_OPTIONS);
    const decoded = decodePresentation(data);

    it('has the same structure as a presentation saved by ProPresenter 7.16', () => {
        const expected = JSON.parse(readFileSync(new URL('../../test/fixtures/propresenter-7.16-structure.json', import.meta.url), 'utf8'));
        assert.deepEqual(structureOf(decoded), expected);
    });

    it('stores the name, one cue per slide, and the text of each slide', () => {
        assert.equal(decoded.name, 'My Song');
        assert.equal((decoded.cues as unknown[]).length, slides.length);
        assert.deepEqual(
            slideRtfs(decoded).map((rtf) => rtfToLines(rtf)),
            slides.map((lines) => (lines.length > 0 ? lines : [''])),
        );
    });

    it('lists every slide in the group, in order, with matching ids', () => {
        const group = (decoded.cue_groups as { cue_identifiers: { string: string }[] }[])[0];
        const cueIds = (decoded.cues as { uuid: { string: string } }[]).map((cue) => cue.uuid.string);
        assert.deepEqual(
            group?.cue_identifiers.map((id) => id.string),
            cueIds,
        );
        assert.equal(new Set(cueIds).size, cueIds.length);
    });

    it('records paragraph ranges in UTF-16 units like ProPresenter', () => {
        assert.deepEqual(paragraphRanges(['ab', '', '\u{1F64F}x']), [
            { start: 0, end: 2 },
            { start: 3, end: 3 },
            { start: 4, end: 7 },
        ]);
    });

    it('produces different ids for every presentation', () => {
        const other = decodePresentation(buildPresentation('My Song', slides, DEFAULT_CONVERSION_OPTIONS));
        assert.notEqual((other.uuid as { string: string }).string, (decoded.uuid as { string: string }).string);
    });
});
