import { randomUUID } from 'node:crypto';
import { getPresentationType } from './proto.js';
import { buildSlideRtf, type RtfStyle } from './rtf.js';

// Builds a ProPresenter 7 presentation (rv.data.Presentation) for a list of lyric slides.
// The message layout mirrors a presentation saved by ProPresenter 7.16.3 for Windows:
// one full-screen, centered, white text box per slide, all slides in a single group.

export interface PresentationOptions extends RtfStyle {
    /** Font family name stored alongside the PostScript name. */
    fontFamily: string;
    /** Name of the single group that holds every slide. */
    groupName: string;
    /** Creation time; defaults to now. */
    now?: Date;
    /** UUID generator; injectable so tests can produce deterministic output. */
    uuid?: () => string;
}

const SLIDE_WIDTH = 1920;
const SLIDE_HEIGHT = 1080;
const ZERO_UUID = '00000000-0000-0000-0000-000000000000';
const WHITE = { red: 1, green: 1, blue: 1, alpha: 1 };
const BLACK = { alpha: 1 };

/** Returns the encoded .pro file for the given slides. */
export function buildPresentation(name: string, slides: string[][], options: PresentationOptions): Uint8Array {
    const Presentation = getPresentationType();
    const object = buildPresentationObject(name, slides, options);
    const message = Presentation.fromObject(object);
    const problem = Presentation.verify(message);
    if (problem) {
        throw new Error(`Generated ProPresenter presentation is invalid: ${problem}`);
    }
    return Presentation.encode(message).finish();
}

/** Plain-object form of the presentation (field names as in the .proto files). */
export function buildPresentationObject(name: string, slides: string[][], options: PresentationOptions): Record<string, unknown> {
    const uuid = options.uuid ?? randomUUID;
    const seconds = Math.floor((options.now ?? new Date()).getTime() / 1000);
    const cueIds = slides.map(() => uuid());

    return {
        application_info: {
            platform: 'PLATFORM_WINDOWS',
            platform_version: { major_version: 10, patch_version: 4294967295, build: '19045' },
            application: 'APPLICATION_PROPRESENTER',
            application_version: { major_version: 7, minor_version: 16, patch_version: 3, build: '118489864' },
        },
        uuid: { string: uuid() },
        name,
        last_date_used: { seconds },
        last_modified_date: { seconds },
        background: { color: BLACK },
        chord_chart: {},
        // ProPresenter accepts (and keeps) a selected arrangement that is not in the arrangement list.
        selected_arrangement: { string: uuid() },
        cue_groups: [
            {
                group: { uuid: { string: uuid() }, name: options.groupName },
                cue_identifiers: cueIds.map((id) => ({ string: id })),
            },
        ],
        cues: slides.map((lines, i) => buildCue(cueIds[i] ?? uuid(), lines, options, uuid)),
        ccli: {},
        timeline: { duration: 300 },
    };
}

function buildCue(cueId: string, lines: string[], options: PresentationOptions, uuid: () => string): Record<string, unknown> {
    return {
        uuid: { string: cueId },
        completion_target_uuid: { string: ZERO_UUID },
        completion_action_type: 'COMPLETION_ACTION_TYPE_LAST',
        completion_action_uuid: { string: ZERO_UUID },
        trigger_time: {},
        actions: [
            {
                uuid: { string: uuid() },
                isEnabled: true,
                type: 'ACTION_TYPE_PRESENTATION_SLIDE',
                slide: {
                    presentation: {
                        base_slide: {
                            elements: [buildTextElement(lines, options, uuid)],
                            background_color: BLACK,
                            size: { width: SLIDE_WIDTH, height: SLIDE_HEIGHT },
                            uuid: { string: uuid() },
                        },
                        notes: buildNotes(),
                        chord_chart: {},
                    },
                },
            },
        ],
        isEnabled: true,
    };
}

function buildTextElement(lines: string[], options: PresentationOptions, uuid: () => string): Record<string, unknown> {
    const shadow = { angle: 315, offset: 5, radius: 5, color: BLACK, opacity: 0.75 };
    return {
        element: {
            uuid: { string: uuid() },
            bounds: { origin: {}, size: { width: SLIDE_WIDTH, height: SLIDE_HEIGHT } },
            opacity: 1,
            path: {
                closed: true,
                points: [
                    { point: {}, q0: {}, q1: {} },
                    { point: { x: 1 }, q0: { x: 1 }, q1: { x: 1 } },
                    { point: { x: 1, y: 1 }, q0: { x: 1, y: 1 }, q1: { x: 1, y: 1 } },
                    { point: { y: 1 }, q0: { y: 1 }, q1: { y: 1 } },
                ],
                shape: { type: 'TYPE_RECTANGLE' },
            },
            // Fill, stroke and shadow are present but not enabled, as in ProPresenter's default text box.
            fill: { color: { red: 0.11764705926179886, green: 0.5647059082984924, blue: 1, alpha: 1 } },
            stroke: { width: 3, color: WHITE },
            shadow,
            feather: { radius: 0.05 },
            text: {
                attributes: {
                    font: { name: options.fontName, size: options.fontSize, family: options.fontFamily, face: 'Regular' },
                    text_solid_fill: WHITE,
                    paragraph_style: { alignment: 'ALIGNMENT_CENTER', line_height_multiple: 1, text_list: {} },
                    stroke_color: WHITE,
                    custom_attributes: paragraphRanges(lines).map((range) => ({
                        range,
                        original_font_size: options.fontSize,
                    })),
                },
                shadow,
                rtf_data: Buffer.from(buildSlideRtf(lines, options), 'latin1'),
                vertical_alignment: 'VERTICAL_ALIGNMENT_MIDDLE',
                scale_behavior: 'SCALE_BEHAVIOR_SCALE_FONT_DOWN',
                margins: {},
                is_superscript_standardized: true,
                transformDelimiter: '  \u2022  ',
                chord_pro: { color: BLACK },
            },
            text_line_mask: {},
        },
        info: 3,
        text_scroller: { scroll_rate: 0.5, should_repeat: true, repeat_distance: 0.052083333333333336 },
    };
}

/**
 * Character ranges of each paragraph in the slide's plain text (lines joined by a newline),
 * in UTF-16 code units, as ProPresenter records them for "scale font down".
 */
export function paragraphRanges(lines: string[]): { start: number; end: number }[] {
    const ranges: { start: number; end: number }[] = [];
    let start = 0;
    for (const line of lines) {
        ranges.push({ start, end: start + line.length });
        start += line.length + 1;
    }
    return ranges;
}

function buildNotes(): Record<string, unknown> {
    const rtf =
        '{\\rtf0\\ansi\\ansicpg1252{\\fonttbl\\f0\\fnil ArialMT;}' +
        '{\\colortbl;\\red0\\green0\\blue0;\\red255\\green255\\blue255;\\red255\\green255\\blue255;}' +
        '{\\*\\expandedcolortbl;\\csgenericrgb\\c0\\c0\\c0\\c100000;\\csgenericrgb\\c100000\\c100000\\c100000\\c100000;\\csgenericrgb\\c100000\\c100000\\c100000\\c0;}' +
        '{\\*\\listtable}{\\*\\listoverridetable}\\uc1\\paperw12240\\margl0\\margr0\\margt0\\margb0' +
        '\\pard\\li0\\fi0\\ri0\\ql\\sb0\\sa0\\sl240\\slmult1\\slleading0\\f0\\b0\\i0\\ul0\\strike0\\fs100\\expnd0\\expndtw0\\cf1\\strokewidth0\\strokec2\\nosupersub\\ulc0\\highlight3\\cb3}';
    return {
        rtf_data: Buffer.from(rtf, 'latin1'),
        attributes: {
            font: { name: 'ArialMT', size: 50, family: 'Arial', face: 'Regular' },
            text_solid_fill: BLACK,
            paragraph_style: { line_height_multiple: 1, text_list: {} },
            stroke_color: WHITE,
        },
    };
}
