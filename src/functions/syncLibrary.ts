import { app, type InvocationContext, type Timer } from '@azure/functions';
import { isSyncEnabled, loadSyncConfig, syncSchedule } from '../lib/config.js';
import { getGraphClient, getJobSettings } from '../lib/runtime.js';
import { syncLibrary } from '../lib/sync.js';

// Optional scheduled sync: every few minutes, converts PowerPoint files in a SharePoint
// folder that have no up-to-date .pro file. Turned on by setting SYNC_DRIVE_ID, and
// replaces the need for a Power Automate flow.

export async function syncLibraryTimer(_timer: Timer, context: InvocationContext): Promise<void> {
    const { output, options } = getJobSettings();
    const config = loadSyncConfig();
    const summary = await syncLibrary(getGraphClient(), config, output, options, context);
    const message =
        `Sync finished: ${summary.scanned} PowerPoint files found, ${summary.pending} needed converting, ` +
        `${summary.converted} converted, ${summary.failed} failed, ${summary.deferred} left for the next run.`;
    if (summary.failed > 0) {
        context.warn(message);
    } else {
        context.log(message);
    }
}

if (isSyncEnabled()) {
    app.timer('syncLibrary', {
        schedule: syncSchedule(),
        runOnStartup: false,
        useMonitor: true,
        handler: syncLibraryTimer,
    });
}
