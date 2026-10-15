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
    /** Folders that exist even when empty. */
    readonly folders = new Set<string>();
    /** Every folder passed to listChildren, in order. */
    readonly listed: string[] = [];
    readonly uploads: { folder: string; name: string; content: Uint8Array }[] = [];
    /** Called after a download completes, before the caller continues (e.g. to simulate an edit). */
    onDownload?: (id: string, call: number) => Promise<void> | void;
    /** Delay applied to each download, per call number, to force interleavings. */
    downloadDelay?: (id: string, call: number) => number;
    /** Return an error to make a download fail (e.g. to simulate SharePoint being unavailable). */
    downloadError?: (id: string) => Error | undefined;
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

    addFolder(path: string): this {
        this.folders.add(path.toLowerCase());
        return this;
    }

    /** Moves a file to another folder; like SharePoint, its id and dates do not change. */
    move(id: string, folder: string): void {
        const file = this.files.get(id);
        if (!file) {
            throw new Error(`no file ${id}`);
        }
        file.folder = folder;
        file.item.parentReference = { driveId: this.driveId, path: `/drives/${this.driveId}/root:${folder ? '/' + encodeURI(folder) : ''}` };
    }

    /** Simulates a metadata-only change (e.g. editing a column): modified time changes, content tag does not. */
    touch(id: string): void {
        const file = this.files.get(id);
        if (!file) {
            throw new Error(`no file ${id}`);
        }
        file.item = { ...file.item, lastModifiedDateTime: this.tick() };
    }

    delete(id: string): void {
        this.files.delete(id);
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
        const failure = this.downloadError?.(itemId);
        if (failure) {
            throw failure;
        }
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

    async listChildren(driveId: string, folder: string, { missingOk = true }: { missingOk?: boolean } = {}): Promise<DriveItem[]> {
        if (driveId !== this.driveId) {
            throw new GraphError('wrong drive', 404);
        }
        this.listed.push(folder);
        const lower = folder.toLowerCase();
        const paths = [...[...this.files.values()].map((f) => f.folder.toLowerCase()), ...this.folders];
        const exists = lower === '' || paths.some((p) => p === lower || p.startsWith(lower + '/'));
        if (!exists && !missingOk) {
            throw new GraphError('Microsoft Graph returned 404 itemNotFound', 404, 'itemNotFound');
        }
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
