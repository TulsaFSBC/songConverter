import type { DriveApi, DriveItem, Logger } from '../../src/lib/graph.js';
import { GraphError } from '../../src/lib/graph.js';

interface StoredFile {
    item: DriveItem;
    folder: string;
    content: Uint8Array;
}

/** In-memory stand-in for a SharePoint document library, with hooks for simulating timing. */
export class FakeDrive implements DriveApi {
    readonly files = new Map<string, StoredFile>();
    readonly uploads: { folder: string; name: string; content: Uint8Array }[] = [];
    /** Called after a download completes, before the caller continues (e.g. to simulate an edit). */
    onDownload?: (id: string, call: number) => Promise<void> | void;
    /** Delay applied to each download, per call number, to force interleavings. */
    downloadDelay?: (id: string, call: number) => number;
    private nextId = 1;
    private clock = Date.parse('2026-01-01T00:00:00Z');
    private downloads = 0;

    constructor(readonly driveId = 'drive-1') {}

    tick(ms = 1000): string {
        this.clock += ms;
        return new Date(this.clock).toISOString();
    }

    now(): number {
        return this.clock;
    }

    add(folder: string, name: string, content: Uint8Array): DriveItem {
        const id = `item-${this.nextId++}`;
        const item: DriveItem = {
            id,
            name,
            cTag: `"c:{${id}},1"`,
            eTag: `"{${id}},1"`,
            size: content.length,
            lastModifiedDateTime: this.tick(),
            file: { mimeType: 'application/octet-stream' },
            parentReference: { driveId: this.driveId, path: `/drives/${this.driveId}/root:${folder ? '/' + encodeURI(folder) : ''}` },
        };
        this.files.set(id, { item, folder, content });
        return item;
    }

    /** Simulates a user saving new content. */
    edit(id: string, content: Uint8Array): void {
        const file = this.files.get(id);
        if (!file) {
            throw new Error(`no file ${id}`);
        }
        const version = Number(/,(\d+)"$/.exec(file.item.cTag ?? '')?.[1] ?? 1) + 1;
        file.content = content;
        file.item = { ...file.item, cTag: `"c:{${id}},${version}"`, eTag: `"{${id}},${version}"`, size: content.length, lastModifiedDateTime: this.tick() };
    }

    async getItem(driveId: string, itemId: string): Promise<DriveItem> {
        const file = this.files.get(itemId);
        if (driveId !== this.driveId || !file) {
            throw new GraphError('Microsoft Graph returned 404 itemNotFound: The resource could not be found.', 404, 'itemNotFound');
        }
        return structuredClone(file.item);
    }

    async download(driveId: string, itemId: string): Promise<Uint8Array> {
        const file = this.files.get(itemId);
        if (driveId !== this.driveId || !file) {
            throw new GraphError('Microsoft Graph returned 404 itemNotFound', 404, 'itemNotFound');
        }
        const call = ++this.downloads;
        const content = file.content;
        const wait = this.downloadDelay?.(itemId, call) ?? 0;
        if (wait > 0) {
            await new Promise((resolve) => setTimeout(resolve, wait));
        }
        await this.onDownload?.(itemId, call);
        return content;
    }

    async upload(driveId: string, folder: string, fileName: string, content: Uint8Array): Promise<DriveItem> {
        if (driveId !== this.driveId) {
            throw new GraphError('wrong drive', 404);
        }
        this.uploads.push({ folder, name: fileName, content });
        const existing = [...this.files.values()].find((f) => f.folder.toLowerCase() === folder.toLowerCase() && f.item.name.toLowerCase() === fileName.toLowerCase());
        if (existing) {
            existing.content = content;
            existing.item = { ...existing.item, size: content.length, lastModifiedDateTime: this.tick() };
            return structuredClone(existing.item);
        }
        return structuredClone(this.add(folder, fileName, content));
    }

    async listChildren(driveId: string, folder: string): Promise<DriveItem[]> {
        if (driveId !== this.driveId) {
            throw new GraphError('wrong drive', 404);
        }
        const lower = folder.toLowerCase();
        const items: DriveItem[] = [];
        const subfolders = new Set<string>();
        for (const file of this.files.values()) {
            const f = file.folder.toLowerCase();
            if (f === lower) {
                items.push(structuredClone(file.item));
            } else if (lower === '' || f.startsWith(lower + '/')) {
                const rest = file.folder.slice(folder ? folder.length + 1 : 0);
                const child = rest.split('/')[0];
                if (child) {
                    subfolders.add(child);
                }
            }
        }
        for (const name of subfolders) {
            items.push({ id: `folder:${folder}/${name}`, name, folder: { childCount: 1 } });
        }
        return items;
    }

    /** The current content of the file at folder/name, if any. */
    contentOf(folder: string, name: string): Uint8Array | undefined {
        return [...this.files.values()].find((f) => f.folder === folder && f.item.name === name)?.content;
    }
}

export const silentLogger: Logger = { log() {}, warn() {}, error() {} };
