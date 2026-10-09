import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { extractSlides, isPowerPointFileName, UnsupportedFileError } from '../src/lib/extract/index.js';
import { normalizeLines } from '../src/lib/extract/slides.js';
import { buildPpt } from './helpers/ppt.js';
import { buildPptx } from './helpers/pptx.js';

describe('pptx extraction', () => {
    it('keeps soft line breaks (<a:br/>) as separate lines and joins runs within a line', async () => {
        const pptx = await buildPptx([
            {
                shapes: [
                    {
                        placeholder: 'ctrTitle',
                        paragraphs: [['I don’t want anything but You', 'You can have all |this world', 'Give me |Jesus']],
                    },
                ],
            },
        ]);
        const [slide] = await extractSlides(pptx);
        assert.deepEqual(slide?.lines, ['I don’t want anything but You', 'You can have all this world', 'Give me Jesus']);
    });

    it('uses presentation order, not file-name order', async () => {
        const slides = Array.from({ length: 11 }, (_, i) => ({ shapes: [{ paragraphs: [[`Slide file ${i + 1}`]] }] }));
        // Show file 11 first, then files 2, 10, 1 ...
        const order = [10, 1, 9, 0, 2, 3, 4, 5, 6, 7, 8];
        const result = await extractSlides(await buildPptx(slides, order));
        assert.deepEqual(
            result.map((s) => s.lines[0]),
            order.map((i) => `Slide file ${i + 1}`),
        );
    });

    it('treats each paragraph as a line and trims blank lines at the edges', async () => {
        const pptx = await buildPptx([{ shapes: [{ paragraphs: [[''], ['Verse line 1'], [''], [''], ['Verse line 2', '']] }] }]);
        const [slide] = await extractSlides(pptx);
        assert.deepEqual(slide?.lines, ['Verse line 1', '', 'Verse line 2']);
    });

    it('decodes XML entities and keeps characters that broke the old converter', async () => {
        const text = 'Quotes "x" & <tags> {braces} back\\slash $& — é Ж \u{1F64F}';
        const pptx = await buildPptx([{ shapes: [{ paragraphs: [[text]] }] }]);
        const [slide] = await extractSlides(pptx);
        assert.deepEqual(slide?.lines, [text]);
    });

    it('does not split a slide on "---" inside the lyrics', async () => {
        const pptx = await buildPptx([{ shapes: [{ paragraphs: [['Line one', '---', 'Line two']] }] }]);
        const slides = await extractSlides(pptx);
        assert.equal(slides.length, 1);
        assert.deepEqual(slides[0]?.lines, ['Line one', '---', 'Line two']);
    });

    it('ignores slide number, date and footer placeholders', async () => {
        const pptx = await buildPptx([
            {
                shapes: [
                    { placeholder: 'title', paragraphs: [['Lyrics']] },
                    { placeholder: 'sldNum', paragraphs: [['7']] },
                    { placeholder: 'dt', paragraphs: [['1/1/2026']] },
                    { placeholder: 'ftr', paragraphs: [['Church name']] },
                ],
            },
        ]);
        const [slide] = await extractSlides(pptx);
        assert.deepEqual(slide?.lines, ['Lyrics']);
    });

    it('reads text from several shapes, groups and tables in document order', async () => {
        const group =
            '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
            '<p:sp><p:nvSpPr><p:cNvPr id="21" name="In group"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>' +
            '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Grouped line</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>';
        const table =
            '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="30" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tr h="1">' +
            '<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell A</a:t></a:r></a:p></a:txBody></a:tc>' +
            '<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell B</a:t></a:r></a:p></a:txBody></a:tc>' +
            '</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
        const pptx = await buildPptx([{ shapes: [{ paragraphs: [['First shape']] }, { paragraphs: [['Second shape']] }], extraXml: group + table }]);
        const [slide] = await extractSlides(pptx);
        assert.deepEqual(slide?.lines, ['First shape', 'Second shape', 'Grouped line', 'Cell A', 'Cell B']);
    });

    it('flags hidden slides', async () => {
        const pptx = await buildPptx([{ shapes: [{ paragraphs: [['Shown']] }] }, { shapes: [{ paragraphs: [['Hidden']] }], hidden: true }]);
        const slides = await extractSlides(pptx);
        assert.deepEqual(
            slides.map((s) => s.hidden),
            [false, true],
        );
    });

    it('keeps empty slides so slide numbering matches PowerPoint', async () => {
        const pptx = await buildPptx([{ shapes: [{ paragraphs: [['A']] }] }, { shapes: [] }, { shapes: [{ paragraphs: [['B']] }] }]);
        const slides = await extractSlides(pptx);
        assert.deepEqual(
            slides.map((s) => s.lines),
            [['A'], [], ['B']],
        );
    });
});

describe('ppt (PowerPoint 97-2003) extraction', () => {
    it('reads placeholder text in slide order with paragraph and line breaks', async () => {
        const ppt = buildPpt([
            { placeholderText: 'Song title' },
            { placeholderText: 'Вот line 1 \rLine 2\u000bLine 3\r' },
            { placeholderText: 'Plain 8-bit text', asBytes: true },
        ]);
        const slides = await extractSlides(ppt);
        assert.deepEqual(
            slides.map((s) => s.lines),
            [['Song title'], ['Вот line 1', 'Line 2', 'Line 3'], ['Plain 8-bit text']],
        );
    });

    it('reads text boxes stored in the slide drawing and skips notes', async () => {
        const ppt = buildPpt([{ placeholderText: 'Title', textBoxText: 'Text box line', notesText: 'Speaker notes' }]);
        const [slide] = await extractSlides(ppt);
        assert.deepEqual(slide?.lines, ['Title', 'Text box line']);
    });

    it('flags hidden slides', async () => {
        const ppt = buildPpt([{ placeholderText: 'Shown' }, { placeholderText: 'Hidden', hidden: true }]);
        const slides = await extractSlides(ppt);
        assert.deepEqual(
            slides.map((s) => s.hidden),
            [false, true],
        );
    });

    it('follows the newest edit in fast-saved files', async () => {
        const ppt = buildPpt([{ placeholderText: 'Old text' }], [{ placeholderText: 'New text' }, { placeholderText: 'Added slide' }]);
        const slides = await extractSlides(ppt);
        assert.deepEqual(
            slides.map((s) => s.lines),
            [['New text'], ['Added slide']],
        );
    });
});

describe('format detection', () => {
    it('rejects files that are not PowerPoint presentations', async () => {
        await assert.rejects(extractSlides(new TextEncoder().encode('just some text')), UnsupportedFileError);
        await assert.rejects(extractSlides(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3])), UnsupportedFileError);
        await assert.rejects(extractSlides(new Uint8Array(0)), UnsupportedFileError);
    });

    it('recognizes PowerPoint file names', () => {
        for (const name of ['Song.pptx', 'SONG.PPTX', 'Song.ppt', 'song.pptm', 'show.ppsx', 'old.pps']) {
            assert.equal(isPowerPointFileName(name), true, name);
        }
        for (const name of ['Song.pro', 'Song.docx', 'pptx', '~$Song.pptx', 'Song.pptx.pro']) {
            assert.equal(isPowerPointFileName(name), false, name);
        }
    });
});

describe('normalizeLines', () => {
    it('removes invisible characters and turns tabs into spaces', () => {
        assert.deepEqual(normalizeLines(['﻿Hello​\tworld  ', '  ']), ['Hello world']);
    });
});
