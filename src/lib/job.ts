import type { OutputConfig } from './config.js';
import { convertPowerPoint, type ConversionOptions } from './convert.js';
import { isPowerPointFileName, UnsupportedFileError } from './extract/index.js';
import type { DriveApi, DriveItem, Logger } from './graph.js';

export type JobResult =
    | {
          status: 'converted';
          source: string;
          output: { name: string; id: string; webUrl?: string; folder: string };
          slideCount: number;
      }
    | { status: 'skipped'; source: string; reason: string };

/** How many times to redo a conversion when the source file changes while it is being converted. */
const MAX_PASSES = 3;

/** Returns the path of an item's parent folder within its drive ("" for the root), or undefined if unknown. */
export function parentFolderPath(item: DriveItem): string | undefined {
    const path = item.parentReference?.path;
    if (path === undefined) {
        return undefined;
    }
    const index = path.indexOf('root:');
    if (index < 0) {
        return undefined;
    }
    const rest = path.slice(index + 'root:'.length);
    try {
        return decodeURIComponent(rest).replace(/^\/+|\/+$/g, '');
    } catch {
        return rest.replace(/^\/+|\/+$/g, '');
    }
}

/** True when folder is the output folder or inside it (case-insensitive, like SharePoint). */
export function isInFolder(folder: string, outputFolder: string): boolean {
    const a = folder.toLowerCase();
    const b = outputFolder.toLowerCase();
    return a === b || a.startsWith(b + '/');
}

/**
 * Converts one PowerPoint file in a SharePoint/OneDrive document library and uploads the
 * resulting .pro file to the output folder. Everything happens in memory, so any number
 * of conversions can run at the same time without interfering with each other.
 */
export async function convertDriveItem(
    graph: DriveApi,
    driveId: string,
    itemId: string,
    output: OutputConfig,
    options: ConversionOptions,
    logger: Logger,
): Promise<JobResult> {
    let item = await graph.getItem(driveId, itemId);
    const source = item.name;

    if (item.deleted || item.folder || !item.file) {
        return { status: 'skipped', source, reason: 'Item is not a file.' };
    }
    if (!isPowerPointFileName(item.name)) {
        return { status: 'skipped', source, reason: 'File is not a PowerPoint presentation.' };
    }
    const outputDriveId = output.driveId ?? driveId;
    const folder = parentFolderPath(item);
    if (outputDriveId === driveId && folder !== undefined && isInFolder(folder, output.folder)) {
        return { status: 'skipped', source, reason: 'File is in the output folder.' };
    }
    if (item.size !== undefined && item.size > output.maxFileBytes) {
        throw new UnsupportedFileError(
            `File is ${Math.round(item.size / 1024 / 1024)} MB, larger than the ${Math.round(output.maxFileBytes / 1024 / 1024)} MB limit (MAX_FILE_SIZE_MB).`,
        );
    }

    for (let pass = 1; ; pass++) {
        const contentTag = item.cTag ?? item.eTag;
        logger.log(`Converting "${item.name}" (${item.size ?? '?'} bytes, pass ${pass}).`);

        const data = await graph.download(driveId, item.id);
        const result = await convertPowerPoint(data, item.name, options);
        const uploaded = await graph.upload(outputDriveId, output.folder, result.fileName, result.content);
        logger.log(`Uploaded "${output.folder}/${uploaded.name}" with ${result.slideCount} slides.`);

        // If the PowerPoint changed while it was being converted (autosave, or another
        // conversion of an older version finishing late), convert it again so the .pro
        // file always ends up matching the latest version.
        const latest = await graph.getItem(driveId, item.id);
        const latestTag = latest.cTag ?? latest.eTag;
        if (latest.deleted || !contentTag || latestTag === contentTag || pass >= MAX_PASSES) {
            return {
                status: 'converted',
                source,
                output: { name: uploaded.name, id: uploaded.id, webUrl: uploaded.webUrl, folder: output.folder },
                slideCount: result.slideCount,
            };
        }
        logger.log(`"${item.name}" changed during conversion; converting the new version.`);
        item = latest;
    }
}
