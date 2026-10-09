import type { OutputConfig, SyncConfig } from './config.js';
import type { ConversionOptions } from './convert.js';
import { isPowerPointFileName, UnsupportedFileError } from './extract/index.js';
import { presentationNameFromFileName } from './convert.js';
import type { DriveApi, DriveItem, Logger } from './graph.js';
import { convertDriveItem, isInFolder } from './job.js';

export interface SyncSummary {
    scanned: number;
    pending: number;
    converted: number;
    failed: number;
    /** Files left for the next run because of SYNC_MAX_FILES_PER_RUN. */
    deferred: number;
}

// Remembers files that cannot be converted (by id + content version) so a broken file is
// not retried every run until it changes. Lives as long as the Function App instance.
// Temporary failures (network, throttling) are not remembered and retry on the next run.
const failedVersions = new Map<string, number>();
const FAILED_RETRY_MS = 6 * 60 * 60 * 1000;

interface SourceFile {
    item: DriveItem;
    folder: string;
}

/** Recursively lists PowerPoint files under the source folder, skipping the output folder. */
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
        for (const item of await graph.listChildren(config.driveId, folder)) {
            const path = folder ? `${folder}/${item.name}` : item.name;
            if (item.folder) {
                if (config.includeSubfolders && !(sameDrive && isInFolder(path, outputFolder))) {
                    folders.push(path);
                }
            } else if (item.file && isPowerPointFileName(item.name)) {
                files.push({ item, folder });
            }
        }
    }
    return files;
}

function timestamp(value: string | undefined): number {
    const time = value ? Date.parse(value) : NaN;
    return Number.isFinite(time) ? time : 0;
}

/**
 * Finds PowerPoint files that have no .pro file yet, or were changed after their .pro file
 * was written, and converts them. Safe to run repeatedly: anything that fails or is
 * deferred is picked up again on the next run.
 */
export async function syncLibrary(
    graph: DriveApi,
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

    for (const [key, failedAt] of failedVersions) {
        if (now - failedAt >= FAILED_RETRY_MS) {
            failedVersions.delete(key);
        }
    }

    const sources = await listSourceFiles(graph, config, output.folder, sameDrive);
    const outputs = new Map<string, DriveItem>();
    for (const item of await graph.listChildren(outputDriveId, output.folder)) {
        if (item.file) {
            outputs.set(item.name.toLowerCase(), item);
        }
    }

    const pending = sources
        .filter(({ item }) => {
            const modified = timestamp(item.lastModifiedDateTime);
            if (now - modified < config.minAgeMs) {
                return false; // Probably still being edited; check again next run.
            }
            const failedAt = failedVersions.get(`${item.id}:${item.cTag ?? item.eTag}`);
            if (failedAt !== undefined && now - failedAt < FAILED_RETRY_MS) {
                return false;
            }
            const existing = outputs.get(`${presentationNameFromFileName(item.name)}.pro`.toLowerCase());
            return !existing || timestamp(existing.lastModifiedDateTime) < modified;
        })
        // Oldest changes first, so a burst of edits cannot starve older files.
        .sort((a, b) => timestamp(a.item.lastModifiedDateTime) - timestamp(b.item.lastModifiedDateTime));

    const batch = pending.slice(0, config.maxFilesPerRun);
    const summary: SyncSummary = { scanned: sources.length, pending: pending.length, converted: 0, failed: 0, deferred: pending.length - batch.length };

    for (const { item, folder } of batch) {
        const label = folder ? `${folder}/${item.name}` : item.name;
        try {
            const result = await convertDriveItem(graph, config.driveId, item.id, output, options, logger);
            if (result.status === 'converted') {
                summary.converted++;
            } else {
                logger.log(`Skipped "${label}": ${result.reason}`);
            }
        } catch (err) {
            summary.failed++;
            if (err instanceof UnsupportedFileError) {
                failedVersions.set(`${item.id}:${item.cTag ?? item.eTag}`, now);
            }
            logger.error(`Failed to convert "${label}": ${err instanceof Error ? err.message : String(err)}`);
        }
    }
    return summary;
}
