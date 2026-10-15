import { DEFAULT_CONVERSION_OPTIONS, type ConversionOptions } from './convert.js';

// All settings come from Function App application settings (environment variables).
// See README.md for the full list.

export class ConfigError extends Error {
    override name = 'ConfigError';
}

function readString(name: string): string | undefined {
    const value = process.env[name]?.trim();
    return value ? value : undefined;
}

function readRequired(name: string): string {
    const value = readString(name);
    if (!value) {
        throw new ConfigError(`Application setting ${name} is required but not set.`);
    }
    return value;
}

function readNumber(name: string, fallback: number, min: number, max: number): number {
    const raw = readString(name);
    if (raw === undefined) {
        return fallback;
    }
    const value = Number(raw);
    if (!Number.isFinite(value) || value < min || value > max) {
        throw new ConfigError(`Application setting ${name} must be a number between ${min} and ${max} (got "${raw}").`);
    }
    return value;
}

function readBoolean(name: string, fallback: boolean): boolean {
    const raw = readString(name)?.toLowerCase();
    if (raw === undefined) {
        return fallback;
    }
    if (['true', '1', 'yes'].includes(raw)) {
        return true;
    }
    if (['false', '0', 'no'].includes(raw)) {
        return false;
    }
    throw new ConfigError(`Application setting ${name} must be true or false (got "${raw}").`);
}

/** Normalizes a folder path within a document library: "/Songs/ProPresenter/" -> "Songs/ProPresenter". */
export function normalizeFolderPath(folder: string): string {
    return folder
        .split(/[\\/]+/)
        .map((segment) => segment.trim())
        .filter(Boolean)
        .join('/');
}

export interface GraphAuthConfig {
    tenantId?: string;
    clientId?: string;
    /** When absent, the Function App's managed identity is used. */
    clientSecret?: string;
}

export function loadGraphAuthConfig(): GraphAuthConfig {
    const clientSecret = readString('CLIENT_SECRET');
    if (clientSecret) {
        return { tenantId: readRequired('TENANT_ID'), clientId: readRequired('CLIENT_ID'), clientSecret };
    }
    // Managed identity: CLIENT_ID optionally selects a user-assigned identity.
    return { clientId: readString('MANAGED_IDENTITY_CLIENT_ID') };
}

export interface OutputConfig {
    /** The one folder (at the top of the document library, no subfolders) that receives every .pro file. */
    folder: string;
    /**
     * Document library (Graph drive id) holding that folder: OUTPUT_DRIVE_ID, else the synced
     * library, so all .pro files land in one place. Only when neither is configured does it
     * fall back to the source file's library.
     */
    driveId?: string;
    /** Files larger than this are not converted. */
    maxFileBytes: number;
}

/** @param syncDriveId Graph drive id of the synced library, if library sync is configured. */
export function loadOutputConfig(syncDriveId?: string): OutputConfig {
    const raw = readString('OUTPUT_FOLDER') ?? 'proPresenter Files';
    const folder = normalizeFolderPath(raw);
    if (!folder) {
        throw new ConfigError('Application setting OUTPUT_FOLDER must name a folder, not the library root.');
    }
    if (folder.includes('/')) {
        throw new ConfigError(
            `Application setting OUTPUT_FOLDER must be a single folder at the top of the document library (e.g. "proPresenter Files"), not a path ("${raw}").`,
        );
    }
    const outputDriveId = readString('OUTPUT_DRIVE_ID');
    if (outputDriveId && asGuid(outputDriveId)) {
        throw new ConfigError(
            `Application setting OUTPUT_DRIVE_ID "${outputDriveId}" is a SharePoint library id, not a Microsoft Graph drive id (which starts with "b!"). ` +
                'Leave OUTPUT_DRIVE_ID empty to use the synced library.',
        );
    }
    return {
        folder,
        driveId: outputDriveId ?? syncDriveId,
        maxFileBytes: readNumber('MAX_FILE_SIZE_MB', 250, 1, 2048) * 1024 * 1024,
    };
}

/** Returns the GUID in "243caec2-…", "{243CAEC2-…}" or "%7B243CAEC2-…%7D" form, lower-cased, or undefined. */
export function asGuid(value: string): string | undefined {
    let text = value.trim();
    try {
        text = decodeURIComponent(text);
    } catch {
        // Not URL-encoded.
    }
    const match = /^\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?$/i.exec(text);
    return match?.[1]?.toLowerCase();
}

/**
 * The document library to sync. Either a Microsoft Graph drive id ("b!…"), or a site URL plus
 * the library's name or id as SharePoint shows them, which is looked up with Graph.
 */
export type LibraryRef = { driveId: string } | { siteUrl: string; library: string };

/** Library sync is enabled by setting SYNC_SITE_URL (with SYNC_LIBRARY) or SYNC_DRIVE_ID. */
export function isSyncEnabled(): boolean {
    return readString('SYNC_DRIVE_ID') !== undefined || readString('SYNC_SITE_URL') !== undefined;
}

export function loadSyncLibrary(): LibraryRef {
    const driveId = readString('SYNC_DRIVE_ID');
    const siteUrl = readString('SYNC_SITE_URL');
    const library = readString('SYNC_LIBRARY');
    if (driveId && !asGuid(driveId)) {
        return { driveId };
    }
    // A GUID in SYNC_DRIVE_ID is the library id from SharePoint's library settings page, not a Graph drive id.
    const libraryRef = library ?? driveId ?? 'Documents';
    if (!siteUrl) {
        throw new ConfigError(
            driveId
                ? `SYNC_DRIVE_ID "${driveId}" is a SharePoint library id, not a Microsoft Graph drive id (which starts with "b!"). ` +
                      'Also set SYNC_SITE_URL to the site address (e.g. https://contoso.sharepoint.com/sites/Worship) and the library will be looked up.'
                : 'Library sync needs SYNC_SITE_URL (the SharePoint site address) or SYNC_DRIVE_ID.',
        );
    }
    return { siteUrl, library: libraryRef };
}

export interface SyncConfig {
    driveId: string;
    sourceFolder: string;
    includeSubfolders: boolean;
    maxFilesPerRun: number;
    /** Files modified more recently than this are left for a later run (they may still be being edited). */
    minAgeMs: number;
    /** On the first run, also convert existing files that have no up-to-date .pro (instead of only recording them). */
    convertExisting: boolean;
}

export function syncSchedule(): string {
    return readString('SYNC_SCHEDULE') ?? '0 */5 * * * *';
}

/** @param driveId Graph drive id of the synced library (resolved from loadSyncLibrary()). */
export function loadSyncConfig(driveId: string): SyncConfig {
    return {
        driveId,
        sourceFolder: normalizeFolderPath(readString('SYNC_SOURCE_FOLDER') ?? ''),
        includeSubfolders: readBoolean('SYNC_INCLUDE_SUBFOLDERS', false),
        maxFilesPerRun: readNumber('SYNC_MAX_FILES_PER_RUN', 25, 1, 500),
        minAgeMs: readNumber('SYNC_MIN_AGE_SECONDS', 60, 0, 86400) * 1000,
        convertExisting: readBoolean('SYNC_CONVERT_EXISTING', false),
    };
}

export function loadConversionOptions(): ConversionOptions {
    const fontName = readString('PRO_FONT_NAME') ?? DEFAULT_CONVERSION_OPTIONS.fontName;
    return {
        fontName,
        fontFamily: readString('PRO_FONT_FAMILY') ?? (readString('PRO_FONT_NAME') ? fontName : DEFAULT_CONVERSION_OPTIONS.fontFamily),
        fontSize: readNumber('PRO_FONT_SIZE', DEFAULT_CONVERSION_OPTIONS.fontSize, 4, 500),
        groupName: readString('PRO_GROUP_NAME') ?? DEFAULT_CONVERSION_OPTIONS.groupName,
        includeHiddenSlides: readBoolean('INCLUDE_HIDDEN_SLIDES', DEFAULT_CONVERSION_OPTIONS.includeHiddenSlides),
    };
}
