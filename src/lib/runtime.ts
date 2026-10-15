import { isSyncEnabled, loadConversionOptions, loadGraphAuthConfig, loadOutputConfig, loadSyncLibrary, type OutputConfig } from './config.js';
import type { ConversionOptions } from './convert.js';
import { GraphClient } from './graph.js';
import { BlobStateStore, type StateStore } from './state.js';

// Shared, lazily created objects. Creating them on first use (rather than at import)
// means a missing setting is reported in the function's logs and response instead of
// stopping the whole Function App from starting.

let graph: GraphClient | undefined;
let stateStore: StateStore | undefined;
const libraryDrives = new Map<string, string>();

export function getGraphClient(): GraphClient {
    graph ??= new GraphClient(loadGraphAuthConfig());
    return graph;
}

export function getStateStore(): StateStore {
    stateStore ??= BlobStateStore.fromEnvironment();
    return stateStore;
}

/** Graph drive id of the synced library, looked up from SYNC_SITE_URL/SYNC_LIBRARY when needed; undefined if sync is off. */
export async function getSyncDriveId(): Promise<string | undefined> {
    if (!isSyncEnabled()) {
        return undefined;
    }
    const ref = loadSyncLibrary();
    if ('driveId' in ref) {
        return ref.driveId;
    }
    const key = `${ref.siteUrl}|${ref.library}`;
    let driveId = libraryDrives.get(key);
    if (!driveId) {
        driveId = await getGraphClient().findLibraryDrive(ref.siteUrl, ref.library);
        libraryDrives.set(key, driveId);
    }
    return driveId;
}

export async function getJobSettings(): Promise<{ output: OutputConfig; options: ConversionOptions }> {
    // The synced library is only needed as the default output library.
    const outputDriveSet = Boolean(process.env.OUTPUT_DRIVE_ID?.trim());
    return { output: loadOutputConfig(outputDriveSet ? undefined : await getSyncDriveId()), options: loadConversionOptions() };
}
