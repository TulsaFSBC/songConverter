import type { OutputConfig, SyncConfig } from './config.js';
import { presentationNameFromFileName, proFileName, type ConversionOptions } from './convert.js';
import { isPowerPointFileName, UnsupportedFileError } from './extract/index.js';
import { ConfigError } from './config.js';
import { GraphError, type DriveApi, type DriveItem, type Logger } from './graph.js';
import { convertDriveItem, isInFolder } from './job.js';
import type { StateStore } from './state.js';

export interface SyncSummary {
    /** True when this run only recorded the files already in the folder (first run). */
    initialized: boolean;
    scanned: number;
    /** Files that appeared since the last run. */
    added: number;
    /** Files whose content changed since they were last handled. */
    changed: number;
    converted: number;
    failed: number;
    /** Added/changed files left for a later run (still being edited, or over SYNC_MAX_FILES_PER_RUN). */
    deferred: number;
}

interface SourceFile {
    item: DriveItem;
    folder: string;
    /** Content version; changes only when the file's content changes. */
    tag: string;
}

/** A file that keeps failing for temporary reasons is retried for this long, then skipped until it changes. */
const RETRY_WINDOW_MS = 24 * 60 * 60 * 1000;

function contentTag(item: DriveItem): string {
    return item.cTag ?? item.eTag ?? item.lastModifiedDateTime ?? '';
}

function timestamp(value: string | undefined): number {
    const time = value ? Date.parse(value) : NaN;
    return Number.isFinite(time) ? time : 0;
}

/** Identifies the watched location, so changing the SYNC_* folder settings starts a fresh baseline. */
function sourceKey(config: SyncConfig): string {
    return `${config.driveId}|${config.sourceFolder.toLowerCase()}|${config.includeSubfolders ? 'subfolders' : 'folder'}`;
}

/** Lists PowerPoint files under the source folder (and subfolders if enabled), skipping the output folder. */
async function listSourceFiles(graph: DriveApi, config: SyncConfig, outputFolder: string, sameDrive: boolean): Promise<SourceFile[]> {
    const files: SourceFile[] = [];
    const folders = [config.sourceFolder];
    const seen = new Set<string>();
    while (folders.length > 0) {
        const folder = folders.shift() ?? '';
        if (seen.has(folder.toLowerCase())) {
            continue;
        }
        seen.add(folder.toLowerCase());
        let children: DriveItem[];
        try {
            // The configured folder must exist; subfolders can disappear between listings.
            children = await graph.listChildren(config.driveId, folder, { missingOk: folder !== config.sourceFolder });
        } catch (err) {
            if (err instanceof GraphError && err.status === 404) {
                throw new ConfigError(
                    `SYNC_SOURCE_FOLDER "${config.sourceFolder}" was not found in the document library SYNC_DRIVE_ID. ` +
                        'Use the folder path inside the library, without the library name (e.g. "Songs", not "Shared Documents/Songs").',
                );
            }
            throw err;
        }
        for (const item of children) {
            const path = folder ? `${folder}/${item.name}` : item.name;
            if (item.folder) {
                if (config.includeSubfolders && !(sameDrive && isInFolder(path, outputFolder))) {
                    folders.push(path);
                }
            } else if (item.file && isPowerPointFileName(item.name)) {
                files.push({ item, folder, tag: contentTag(item) });
            }
        }
    }
    return files;
}

/**
 * Converts PowerPoint files that were added to the watched folder, or whose content
 * changed, since the previous run. The first run only records the files already there
 * (unless SYNC_CONVERT_EXISTING is on), so existing songs are not converted in bulk.
 *
 * Files are tracked by id and content version rather than by date, so a file moved or
 * copied into the folder (which keeps its old dates) still counts as added. Files that
 * are still being edited, over the per-run limit, or failed for a temporary reason are
 * not marked as handled and are picked up by the next run.
 */
export async function syncLibrary(
    graph: DriveApi,
    store: StateStore,
    config: SyncConfig,
    output: OutputConfig,
    options: ConversionOptions,
    logger: Logger,
    now = Date.now(),
): Promise<SyncSummary> {
    const outputDriveId = output.driveId ?? config.driveId;
    const sameDrive = outputDriveId === config.driveId;
    if (sameDrive && isInFolder(config.sourceFolder, output.folder)) {
        throw new Error('SYNC_SOURCE_FOLDER must not be the OUTPUT_FOLDER or inside it.');
    }

    const source = sourceKey(config);
    const previous = await store.load();
    const sources = await listSourceFiles(graph, config, output.folder, sameDrive);
    const summary: SyncSummary = { initialized: false, scanned: sources.length, added: 0, changed: 0, converted: 0, failed: 0, deferred: 0 };

    let handled: Record<string, string>;
    let failing: Record<string, string> = {};
    if (previous && previous.source === source) {
        handled = previous.files;
        failing = previous.failing ?? {};
    } else {
        // First run (or the watched folder changed): record what is already there.
        summary.initialized = true;
        handled = {};
        const isUpToDate = config.convertExisting ? await upToDateCheck(graph, outputDriveId, output.folder) : () => true;
        for (const file of sources) {
            if (isUpToDate(file)) {
                handled[file.item.id] = file.tag;
            }
        }
    }

    const pending: SourceFile[] = [];
    for (const file of sources) {
        const known = handled[file.item.id];
        if (known === file.tag) {
            continue;
        }
        if (known === undefined) {
            summary.added++;
        } else {
            summary.changed++;
        }
        if (now - timestamp(file.item.lastModifiedDateTime) < config.minAgeMs) {
            summary.deferred++; // Probably still being edited; check again next run.
            continue;
        }
        pending.push(file);
    }
    // Oldest changes first, so a burst of edits cannot starve older files.
    pending.sort((a, b) => timestamp(a.item.lastModifiedDateTime) - timestamp(b.item.lastModifiedDateTime));
    const batch = pending.slice(0, config.maxFilesPerRun);
    summary.deferred += pending.length - batch.length;

    const results = new Map<string, string>();
    const stillFailing: Record<string, string> = {};
    for (const file of batch) {
        const label = file.folder ? `${file.folder}/${file.item.name}` : file.item.name;
        try {
            const result = await convertDriveItem(graph, config.driveId, file.item.id, output, options, logger);
            if (result.status === 'converted') {
                summary.converted++;
                results.set(file.item.id, result.contentTag ?? file.tag);
            } else {
                logger.log(`Skipped "${label}": ${result.reason}`);
                results.set(file.item.id, file.tag);
            }
        } catch (err) {
            summary.failed++;
            const message = err instanceof Error ? err.message : String(err);
            if (err instanceof UnsupportedFileError) {
                // Retrying will not help until someone changes the file.
                logger.error(`Cannot convert "${label}": ${message}`);
                results.set(file.item.id, file.tag);
            } else {
                const firstFailed = failing[file.item.id] ?? new Date(now).toISOString();
                if (now - timestamp(firstFailed) >= RETRY_WINDOW_MS) {
                    logger.error(`Failed to convert "${label}" since ${firstFailed}; skipping it until it changes: ${message}`);
                    results.set(file.item.id, file.tag);
                } else {
                    logger.error(`Failed to convert "${label}"; will retry next run: ${message}`);
                    stillFailing[file.item.id] = firstFailed;
                }
            }
        }
    }

    // Rebuild the state from the current listing, which also forgets deleted files.
    const files: Record<string, string> = {};
    const failingNext: Record<string, string> = {};
    for (const file of sources) {
        const id = file.item.id;
        const tag = results.get(id) ?? handled[id];
        if (tag !== undefined) {
            files[id] = tag;
        }
        // Keep failure times for files that failed again or were not attempted this run.
        const failedAt = stillFailing[id] ?? (results.has(id) ? undefined : failing[id]);
        if (failedAt !== undefined) {
            failingNext[id] = failedAt;
        }
    }
    await store.save({ version: 1, source, files, failing: failingNext, updated: new Date(now).toISOString() });
    return summary;
}

/** For SYNC_CONVERT_EXISTING: "already has a .pro written after the PowerPoint last changed". */
async function upToDateCheck(graph: DriveApi, driveId: string, folder: string): Promise<(file: SourceFile) => boolean> {
    const outputs = new Map<string, DriveItem>();
    for (const item of await graph.listChildren(driveId, folder)) {
        if (item.file) {
            outputs.set(item.name.toLowerCase(), item);
        }
    }
    return (file) => {
        const existing = outputs.get(proFileName(presentationNameFromFileName(file.item.name)).toLowerCase());
        // Empty .pro files (left by the old converter) are never up to date.
        return existing !== undefined && (existing.size ?? 1) > 0 && timestamp(existing.lastModifiedDateTime) >= timestamp(file.item.lastModifiedDateTime);
    };
}
