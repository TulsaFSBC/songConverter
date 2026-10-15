import { app, type InvocationContext, type Timer } from '@azure/functions';
import { isSyncEnabled, loadSyncConfig, syncSchedule } from '../lib/config.js';
import { getGraphClient, getJobSettings, getStateStore } from '../lib/runtime.js';
import { syncLibrary } from '../lib/sync.js';

// Scheduled sync: every few minutes, converts PowerPoint files added to (or changed in)
// a SharePoint folder since the previous run, replacing the need for a Power Automate flow.
// The function is always registered so it shows up in the Azure portal; it does nothing
// until SYNC_DRIVE_ID is set. To stop it completely, disable the function in the portal.

export async function syncLibraryTimer(_timer: Timer, context: InvocationContext): Promise<void> {
    if (!isSyncEnabled()) {
        context.log(
            'Library sync is off: set the SYNC_DRIVE_ID application setting to turn it on, ' +
                'or disable the syncLibrary function if you only use Power Automate.',
        );
        return;
    }
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

app.timer('syncLibrary', {
    schedule: syncSchedule(),
    runOnStartup: false,
    useMonitor: true,
    handler: syncLibraryTimer,
});
