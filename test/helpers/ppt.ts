import CFB from 'cfb';

// Builds minimal PowerPoint 97-2003 (.ppt) files for tests, following [MS-PPT].

export interface TestPptSlide {
    /** Placeholder text, stored in the document's SlideListWithText. Uses \r between paragraphs and \v for line breaks. */
    placeholderText?: string;
    /** Store placeholder text as 8-bit TextBytesAtom instead of UTF-16 TextCharsAtom. */
    asBytes?: boolean;
    /** Text box text stored inside the slide's own drawing. */
    textBoxText?: string;
    /** Notes-type text that must be ignored. */
    notesText?: string;
    /** Text in a footer placeholder shape (slide furniture) that must be ignored. */
    footerText?: string;
    hidden?: boolean;
}

function record(type: number, body: Buffer, instance = 0, version = 0): Buffer {
    const header = Buffer.alloc(8);
    header.writeUInt16LE((version & 0xf) | (instance << 4), 0);
    header.writeUInt16LE(type, 2);
    header.writeUInt32LE(body.length, 4);
    return Buffer.concat([header, body]);
}

function container(type: number, children: Buffer[], instance = 0): Buffer {
    return record(type, Buffer.concat(children), instance, 0xf);
}

function u32(...values: number[]): Buffer {
    const buf = Buffer.alloc(values.length * 4);
    values.forEach((v, i) => buf.writeUInt32LE(v >>> 0, i * 4));
    return buf;
}

function textRecords(text: string, textType: number, asBytes = false): Buffer[] {
    return [
        record(0x0f9f, u32(textType)),
        asBytes ? record(0x0fa8, Buffer.from(text, 'latin1')) : record(0x0fa0, Buffer.from(text, 'utf16le')),
    ];
}

function slideContainer(slide: TestPptSlide): Buffer {
    const info = Buffer.alloc(16);
    info.writeUInt16LE(slide.hidden ? 0x0004 : 0, 10);
    const textBox = slide.textBoxText
        ? [container(0xf004, [container(0xf00d, textRecords(slide.textBoxText, 4))])]
        : [];
    if (slide.footerText) {
        // OfficeArtClientData with a PlaceholderAtom whose placementId (byte 4) is 0x09 = footer.
        const placeholder = Buffer.alloc(8);
        placeholder.writeUInt8(0x09, 4);
        textBox.push(container(0xf004, [container(0xf011, [record(0x0bc3, placeholder)]), container(0xf00d, textRecords(slide.footerText, 4))]));
    }
    const drawing = container(0x040c, [container(0xf002, [container(0xf003, textBox)])]);
    return container(0x03ee, [record(0x03ef, Buffer.alloc(24), 0, 2), record(0x03f9, info), drawing]);
}

function documentContainer(slides: TestPptSlide[], firstSlidePersistId: number): Buffer {
    const entries: Buffer[] = [];
    slides.forEach((slide, i) => {
        entries.push(record(0x03f3, u32(firstSlidePersistId + i, 0, 0, 256 + i, 0)));
        if (slide.placeholderText !== undefined) {
            entries.push(...textRecords(slide.placeholderText, 1, slide.asBytes));
        }
        if (slide.notesText !== undefined) {
            entries.push(...textRecords(slide.notesText, 2));
        }
    });
    const masterList = container(0x0ff0, [record(0x03f3, u32(99, 0, 0, 0x80000001, 0))], 1);
    return container(0x03e8, [record(0x03e9, Buffer.alloc(40), 0, 1), masterList, container(0x0ff0, entries, 0)]);
}

function persistDirectory(firstId: number, offsets: number[]): Buffer {
    return record(0x1772, u32(firstId | (offsets.length << 20), ...offsets));
}

function userEdit(offsetLastEdit: number, offsetPersistDirectory: number, docPersistId: number, seed: number): Buffer {
    const body = Buffer.alloc(28);
    body.writeUInt32LE(256, 0);
    body.writeUInt16LE(0x0000, 4);
    body.writeUInt8(0, 6);
    body.writeUInt8(3, 7);
    body.writeUInt32LE(offsetLastEdit, 8);
    body.writeUInt32LE(offsetPersistDirectory, 12);
    body.writeUInt32LE(docPersistId, 16);
    body.writeUInt32LE(seed, 20);
    body.writeUInt16LE(1, 24);
    return record(0x0ff5, body);
}

function currentUser(offsetToCurrentEdit: number): Buffer {
    const body = Buffer.alloc(20);
    body.writeUInt32LE(0x14, 0);
    body.writeUInt32LE(0xe391c05f, 4);
    body.writeUInt32LE(offsetToCurrentEdit, 8);
    body.writeUInt16LE(0, 12);
    body.writeUInt16LE(0x03f4, 14);
    body.writeUInt8(3, 16);
    return record(0x0ff6, body);
}

/**
 * Builds a .ppt file. With `editedSlides`, a second "fast save" edit is appended that
 * replaces the document (and slides), so readers must follow the newest edit.
 */
export function buildPpt(slides: TestPptSlide[], editedSlides?: TestPptSlide[], { encrypted = false } = {}): Uint8Array {
    const parts: Buffer[] = [];
    let offset = 0;
    const append = (buf: Buffer): number => {
        const at = offset;
        parts.push(buf);
        offset += buf.length;
        return at;
    };

    // Edit 1: document = persist id 1, slides = persist ids 2..n+1.
    const docOffset = append(documentContainer(slides, 2));
    const slideOffsets = slides.map((slide) => append(slideContainer(slide)));
    const dir1 = append(persistDirectory(1, [docOffset, ...slideOffsets]));
    let currentEdit = append(userEdit(0, dir1, 1, slides.length + 1));

    if (editedSlides) {
        // Edit 2: new document and slides with fresh persist ids, older edit chained via offsetLastEdit.
        const firstId = slides.length + 2;
        const doc2 = append(documentContainer(editedSlides, firstId + 1));
        const slide2Offsets = editedSlides.map((slide) => append(slideContainer(slide)));
        const dir2 = append(persistDirectory(firstId, [doc2, ...slide2Offsets]));
        currentEdit = append(userEdit(currentEdit, dir2, firstId, firstId + editedSlides.length));
    }

    const cfb = CFB.utils.cfb_new();
    CFB.utils.cfb_add(cfb, '/PowerPoint Document', Buffer.concat(parts));
    const user = currentUser(currentEdit);
    if (encrypted) {
        user.writeUInt32LE(0xf3d1c4df, 12); // headerToken of an encrypted document
    }
    CFB.utils.cfb_add(cfb, '/Current User', user);
    return new Uint8Array(CFB.write(cfb, { type: 'buffer' }) as Buffer);
}
