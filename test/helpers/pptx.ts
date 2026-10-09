import JSZip from 'jszip';

// Builds small but realistic .pptx files for tests.

export interface TestShape {
    /** Paragraphs; each paragraph is a list of lines joined by <a:br/>. Each line may be split into runs with '|'. */
    paragraphs: string[][];
    placeholder?: string;
}

export interface TestSlide {
    shapes: TestShape[];
    hidden?: boolean;
    /** Raw XML inserted into the shape tree (for groups, tables, ...). */
    extraXml?: string;
}

const NS =
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

export function escapeXml(text: string): string {
    return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function runXml(text: string): string {
    return `<a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeXml(text)}</a:t></a:r>`;
}

export function shapeXml(shape: TestShape, id: number): string {
    const ph = shape.placeholder ? `<p:ph type="${shape.placeholder}"/>` : '';
    const paragraphs = shape.paragraphs
        .map((lines) => {
            const content = lines.map((line) => line.split('|').map(runXml).join('')).join('<a:br><a:rPr lang="en-US"/></a:br>');
            return `<a:p><a:pPr algn="ctr"/>${content}<a:endParaRPr lang="en-US" dirty="0"/></a:p>`;
        })
        .join('');
    return (
        `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
        `<p:spPr/><p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`
    );
}

function slideXml(slide: TestSlide): string {
    const show = slide.hidden ? ' show="0"' : '';
    const shapes = slide.shapes.map((shape, i) => shapeXml(shape, i + 2)).join('');
    return (
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<p:sld ${NS}${show}><p:cSld><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
        `${shapes}${slide.extraXml ?? ''}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
    );
}

/**
 * Builds a .pptx. `order` gives the presentation order as indexes into `slides`
 * (file slideN.xml holds slides[N-1]), so tests can check that file names don't decide order.
 */
export async function buildPptx(slides: TestSlide[], order?: number[]): Promise<Uint8Array> {
    const zip = new JSZip();
    const showOrder = order ?? slides.map((_, i) => i);
    zip.file(
        '[Content_Types].xml',
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Default Extension="xml" ContentType="application/xml"/></Types>',
    );
    const rels = slides
        .map((_, i) => `<Relationship Id="rId${i + 10}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`)
        .join('');
    zip.file(
        'ppt/_rels/presentation.xml.rels',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>${rels}</Relationships>`,
    );
    const ids = showOrder.map((index, i) => `<p:sldId id="${256 + i}" r:id="rId${index + 10}"/>`).join('');
    zip.file(
        'ppt/presentation.xml',
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation ${NS}><p:sldMasterIdLst/><p:sldIdLst>${ids}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
    );
    slides.forEach((slide, i) => zip.file(`ppt/slides/slide${i + 1}.xml`, slideXml(slide)));
    return zip.generateAsync({ type: 'uint8array' });
}
