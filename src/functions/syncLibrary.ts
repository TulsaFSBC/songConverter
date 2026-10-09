import { app, type InvocationContext, type Timer } from '@azure/functions';
import { isSyncEnabled, loadSyncConfig, syncSchedule } from '../lib/config.js';
import { getGraphClient, getJobSettings, getStateStore } from '../lib/runtime.js';
import { syncLibrary } from '../lib/sync.js';

// Optional scheduled sync: every few minutes, converts PowerPoint files added to (or
// changed in) a SharePoint folder since the previous run. Turned on by setting
// SYNC_DRIVE_ID, and replaces the need for a Power Automate flow.

export async function syncLibraryTimer(_timer: Timer, context: InvocationContext): Promise<void> {
    const { output, options } = getJobSettings();
    const config = loadSyncConfig();
    const summary = await syncLibrary(getGraphClient(), getStateStore(), config, output, options, context);
    const message = summary.initialized && !config.convertExisting
        ? `Sync started: recorded the ${summary.scanned} PowerPoint files already in the folder. ` +
          'Files added or changed from now on will be converted.'
        : `Sync finished: ${summary.scanned} PowerPoint files, ${summary.added} added and ${summary.changed} changed since the last run; ` +
          `${summary.converted} converted, ${summary.failed} failed, ${summary.deferred} waiting for the next run.`;
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
