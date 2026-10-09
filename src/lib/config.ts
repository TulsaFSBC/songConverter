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
    /** Folder (path inside the document library) that receives the .pro files. */
    folder: string;
    /** Document library for the .pro files; defaults to the source file's library. */
    driveId?: string;
    /** Files larger than this are not converted. */
    maxFileBytes: number;
}

export function loadOutputConfig(): OutputConfig {
    const folder = normalizeFolderPath(readString('OUTPUT_FOLDER') ?? 'proPresenter Files');
    if (!folder) {
        throw new ConfigError('Application setting OUTPUT_FOLDER must name a folder, not the library root.');
    }
    return {
        folder,
        driveId: readString('OUTPUT_DRIVE_ID'),
        maxFileBytes: readNumber('MAX_FILE_SIZE_MB', 250, 1, 2048) * 1024 * 1024,
    };
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

/** Library sync is enabled by setting SYNC_DRIVE_ID. */
export function isSyncEnabled(): boolean {
    return readString('SYNC_DRIVE_ID') !== undefined;
}

export function syncSchedule(): string {
    return readString('SYNC_SCHEDULE') ?? '0 */5 * * * *';
}

export function loadSyncConfig(): SyncConfig {
    return {
        driveId: readRequired('SYNC_DRIVE_ID'),
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
