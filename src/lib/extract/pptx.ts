import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { normalizeLines, splitLines, UnsupportedFileError, type SlideText } from './slides.js';

// With preserveOrder, every element is an object with one tag-name key holding
// its children, plus an optional ':@' key holding its attributes.
type XmlNode = Record<string, unknown>;

const xmlParser = new XMLParser({
    preserveOrder: true,
    ignoreAttributes: false,
    attributeNamePrefix: '',
    // Keep spaces at the edges of text runs: "You can have all " + "this world".
    trimValues: false,
    parseTagValue: false,
    parseAttributeValue: false,
    processEntities: true,
    htmlEntities: true,
});

// Placeholders that hold slide furniture rather than lyrics.
const IGNORED_PLACEHOLDERS = new Set(['sldNum', 'dt', 'ftr', 'hdr']);

// Element holding each shape kind's non-visual properties (p:cNvPr, which carries hidden="1").
const NON_VISUAL: Record<string, string> = { 'p:sp': 'p:nvSpPr', 'p:grpSp': 'p:nvGrpSpPr', 'p:graphicFrame': 'p:nvGraphicFramePr' };

const REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';

/** Extracts the text of every slide, in presentation order, from a .pptx/.pptm/.ppsx file. */
export async function extractPptx(data: Uint8Array): Promise<SlideText[]> {
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(data);
    } catch (err) {
        throw new UnsupportedFileError(`File is not a valid PowerPoint (.pptx) package: ${(err as Error).message}`);
    }

    const slidePaths = await getSlidePaths(zip);
    if (slidePaths.length === 0) {
        throw new UnsupportedFileError('PowerPoint file does not contain any slides.');
    }

    const slides: SlideText[] = [];
    for (const slidePath of slidePaths) {
        const file = zip.file(slidePath);
        if (!file) {
            continue;
        }
        const xml = await file.async('string');
        slides.push(parseSlide(xml));
    }
    return slides;
}

/**
 * Returns slide part paths in the order the presentation shows them
 * (p:sldIdLst in presentation.xml), not the order of file names in the zip:
 * slide10.xml can come before slide2.xml, and reordered slides keep their file names.
 */
async function getSlidePaths(zip: JSZip): Promise<string[]> {
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    const relsXml = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string');

    if (presentationXml && relsXml) {
        const targets = new Map<string, string>();
        for (const rel of select(parseXml(relsXml), ['Relationships', 'Relationship'])) {
            const attrs = attributes(rel);
            if (attrs.Type === REL_SLIDE && attrs.Id && attrs.Target) {
                targets.set(attrs.Id, resolvePartPath('ppt/', attrs.Target));
            }
        }
        const ordered: string[] = [];
        for (const slideId of select(parseXml(presentationXml), ['p:presentation', 'p:sldIdLst', 'p:sldId'])) {
            const target = targets.get(attributes(slideId)['r:id'] ?? '');
            if (target) {
                ordered.push(target);
            }
        }
        if (ordered.length > 0) {
            return ordered;
        }
    }

    // Fallback for packages without a usable slide list: numeric file-name order.
    return Object.keys(zip.files)
        .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
        .sort((a, b) => slideNumber(a) - slideNumber(b));
}

function slideNumber(path: string): number {
    return Number(/slide(\d+)\.xml$/.exec(path)?.[1] ?? 0);
}

/** Resolves a relationship target (relative to the source part's folder, or absolute) to a zip path. */
function resolvePartPath(baseDir: string, target: string): string {
    const parts = target.startsWith('/') ? [] : baseDir.split('/').filter(Boolean);
    for (const segment of target.split('/')) {
        if (segment === '..') {
            parts.pop();
        } else if (segment !== '.' && segment !== '') {
            parts.push(segment);
        }
    }
    return parts.join('/');
}

function parseSlide(xml: string): SlideText {
    const doc = parseXml(xml);
    const slide = select(doc, ['p:sld'])[0];
    const show = slide ? attributes(slide).show : undefined;
    const hidden = show === '0' || show === 'false';

    const lines: string[] = [];
    for (const spTree of select(doc, ['p:sld', 'p:cSld', 'p:spTree'])) {
        collectShapeTree(children(spTree, 'p:spTree'), lines);
    }
    return { lines: normalizeLines(lines), hidden };
}

/** Walks shapes in document order, collecting text from text shapes, groups and tables. */
function collectShapeTree(nodes: XmlNode[], lines: string[]): void {
    for (const node of nodes) {
        const tag = tagName(node);
        if (isHiddenShape(node, tag)) {
            continue; // Hidden with the Selection Pane: not shown in PowerPoint, so not projected either.
        }
        switch (tag) {
            case 'p:sp':
                if (!isIgnoredPlaceholder(node)) {
                    for (const txBody of findAll(children(node, tag), 'p:txBody')) {
                        collectTextBody(children(txBody, 'p:txBody'), lines);
                    }
                }
                break;
            case 'p:grpSp':
                collectShapeTree(children(node, tag), lines);
                break;
            case 'p:graphicFrame':
                // Tables: a:graphic/a:graphicData/a:tbl/a:tr/a:tc/a:txBody
                for (const txBody of findDeep(children(node, tag), 'a:txBody')) {
                    collectTextBody(children(txBody, 'a:txBody'), lines);
                }
                break;
            case 'mc:AlternateContent': {
                // Use the fallback (or first choice) representation only, so content isn't duplicated.
                const branch = findAll(children(node, tag), 'mc:Fallback')[0] ?? findAll(children(node, tag), 'mc:Choice')[0];
                if (branch) {
                    collectShapeTree(children(branch, tagName(branch)), lines);
                }
                break;
            }
            default:
                break;
        }
    }
}

function isHiddenShape(node: XmlNode, tag: string): boolean {
    const nonVisual = NON_VISUAL[tag];
    const cNvPr = nonVisual ? select(children(node, tag), [nonVisual, 'p:cNvPr'])[0] : undefined;
    const hidden = cNvPr ? attributes(cNvPr).hidden : undefined;
    return hidden === '1' || hidden === 'true';
}

function isIgnoredPlaceholder(shape: XmlNode): boolean {
    const ph = select(children(shape, 'p:sp'), ['p:nvSpPr', 'p:nvPr', 'p:ph'])[0];
    return ph !== undefined && IGNORED_PLACEHOLDERS.has(attributes(ph).type ?? '');
}

/** Each a:p paragraph starts a new line; each a:br inside a paragraph also starts a new line. */
function collectTextBody(nodes: XmlNode[], lines: string[]): void {
    for (const paragraph of findAll(nodes, 'a:p')) {
        let current = '';
        for (const part of children(paragraph, 'a:p')) {
            const tag = tagName(part);
            if (tag === 'a:r' || tag === 'a:fld') {
                for (const t of findAll(children(part, tag), 'a:t')) {
                    current += textOf(t);
                }
            } else if (tag === 'a:br') {
                lines.push(...splitLines(current));
                current = '';
            }
        }
        // A text run may itself contain line separators (e.g. a vertical tab typed in PowerPoint).
        lines.push(...splitLines(current));
    }
}

function parseXml(xml: string): XmlNode[] {
    return xmlParser.parse(xml) as XmlNode[];
}

function tagName(node: XmlNode): string {
    for (const key of Object.keys(node)) {
        if (key !== ':@') {
            return key;
        }
    }
    return '';
}

function children(node: XmlNode, tag: string): XmlNode[] {
    const value = node[tag];
    return Array.isArray(value) ? (value as XmlNode[]) : [];
}

function attributes(node: XmlNode): Record<string, string> {
    return (node[':@'] as Record<string, string> | undefined) ?? {};
}

/** Direct children with the given tag. */
function findAll(nodes: XmlNode[], tag: string): XmlNode[] {
    return nodes.filter((node) => tagName(node) === tag);
}

/** Follows a path of tag names down from a list of sibling nodes, e.g. ['p:sld', 'p:cSld', 'p:spTree']. */
function select(nodes: XmlNode[], path: string[]): XmlNode[] {
    let current = nodes;
    for (const [index, tag] of path.entries()) {
        const matches = findAll(current, tag);
        if (index === path.length - 1) {
            return matches;
        }
        current = matches.flatMap((match) => children(match, tag));
    }
    return current;
}

/** All descendants with the given tag, in document order, without descending into matches. */
function findDeep(nodes: XmlNode[], tag: string): XmlNode[] {
    const result: XmlNode[] = [];
    for (const node of nodes) {
        const name = tagName(node);
        if (name === tag) {
            result.push(node);
        } else if (name !== '#text') {
            result.push(...findDeep(children(node, name), tag));
        }
    }
    return result;
}

function textOf(node: XmlNode): string {
    let text = '';
    for (const child of children(node, tagName(node))) {
        const value = child['#text'];
        if (typeof value === 'string' || typeof value === 'number') {
            text += String(value);
        }
    }
    return text;
}
