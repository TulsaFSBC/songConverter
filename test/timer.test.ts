import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import functions from '@azure/functions';
import { syncLibraryTimer } from '../src/functions/syncLibrary.js';

// @azure/functions is CommonJS; its classes are only reachable through the default export from ESM.
const { InvocationContext } = functions;

describe('syncLibrary timer function', () => {
    it('is always registered, and does nothing but log while SYNC_DRIVE_ID is not set', async () => {
        delete process.env.SYNC_DRIVE_ID;
        const messages: string[] = [];
        const context = new InvocationContext({
            functionName: 'syncLibrary',
            invocationId: 'test',
            logHandler: (_level, ...args) => messages.push(args.join(' ')),
        });
        await syncLibraryTimer({ isPastDue: false, scheduleStatus: undefined } as never, context);
        assert.equal(messages.length, 1);
        assert.match(messages[0] ?? '', /Library sync is off/);
    });
});
