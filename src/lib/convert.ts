import { extractSlides, POWERPOINT_EXTENSIONS, UnsupportedFileError } from './extract/index.js';
import { buildPresentation, type PresentationOptions } from './pro/presentation.js';

export interface ConversionOptions extends Omit<PresentationOptions, 'now' | 'uuid'> {
    /** Include slides that are hidden in PowerPoint. */
    includeHiddenSlides: boolean;
}

export interface ConversionResult {
    /** Name of the .pro file, e.g. "Amazing Grace.pro". */
    fileName: string;
    /** Presentation name shown in ProPresenter. */
    presentationName: string;
    content: Uint8Array;
    slideCount: number;
    /** The lyric lines of each slide, as written to the .pro file. */
    slides: string[][];
}

export const DEFAULT_CONVERSION_OPTIONS: ConversionOptions = {
    fontName: 'TimesNewRomanPSMT',
    fontFamily: 'TimesNewRomanPSMT',
    fontSize: 90,
    groupName: 'Group',
    includeHiddenSlides: false,
};

/** "Amazing Grace.PPTX" -> "Amazing Grace" */
export function presentationNameFromFileName(fileName: string): string {
    const lower = fileName.toLowerCase();
    const ext = POWERPOINT_EXTENSIONS.find((e) => lower.endsWith(e));
    const base = ext ? fileName.slice(0, -ext.length) : fileName.replace(/\.[^.]*$/, '');
    return base.trim() || 'Untitled';
}

/** Converts a PowerPoint file (.pptx/.ppt) into a ProPresenter 7 .pro file. */
export async function convertPowerPoint(
    data: Uint8Array,
    sourceFileName: string,
    options: ConversionOptions = DEFAULT_CONVERSION_OPTIONS,
): Promise<ConversionResult> {
    const extracted = await extractSlides(data);
    const slides = extracted.filter((slide) => options.includeHiddenSlides || !slide.hidden).map((slide) => slide.lines);

    if (!slides.some((lines) => lines.length > 0)) {
        throw new UnsupportedFileError('PowerPoint file does not contain any slide text.');
    }

    const presentationName = presentationNameFromFileName(sourceFileName);
    return {
        fileName: `${presentationName}.pro`,
        presentationName,
        content: buildPresentation(presentationName, slides, options),
        slideCount: slides.length,
        slides,
    };
}
