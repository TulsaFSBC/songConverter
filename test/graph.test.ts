import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import type { TokenCredential } from '@azure/identity';
import { ConfigError } from '../src/lib/config.js';
import { GraphClient } from '../src/lib/graph.js';

const credential: TokenCredential = { getToken: async () => ({ token: 'test-token', expiresOnTimestamp: Date.now() + 3_600_000 }) };
const realFetch = globalThis.fetch;

/** Answers Graph requests from a table of URL substrings to JSON bodies (404 when nothing matches). */
function fakeGraph(routes: Record<string, unknown>): string[] {
    const requested: string[] = [];
    globalThis.fetch = (async (input: string | URL | Request) => {
        const url = String(input);
        requested.push(url);
        const match = Object.keys(routes).find((key) => url.includes(key));
        if (match === undefined) {
            return new Response(JSON.stringify({ error: { code: 'itemNotFound', message: 'Not found' } }), { status: 404 });
        }
        return new Response(JSON.stringify(routes[match]), { status: 200, headers: { 'Content-Type': 'application/json' } });
    }) as typeof fetch;
    return requested;
}

describe('finding a document library', () => {
    afterEach(() => {
        globalThis.fetch = realFetch;
    });

    const site = { '/sites/contoso.sharepoint.com:/sites/Worship?': { id: 'contoso.sharepoint.com,site-guid,web-guid' } };
    const drives = {
        '/drives?$select': {
            value: [
                { id: 'b!documents', name: 'Documents', webUrl: 'https://contoso.sharepoint.com/sites/Worship/Shared%20Documents' },
                { id: 'b!songs', name: 'Song Library', webUrl: 'https://contoso.sharepoint.com/sites/Worship/SongLibrary' },
            ],
        },
    };

    it('by the library id from the library settings page', async () => {
        const requested = fakeGraph({ ...site, '/lists/243caec2-9363-411a-869a-ae66e37ba368/drive': { id: 'b!from-list-id' } });
        const graph = new GraphClient({}, credential);
        assert.equal(await graph.findLibraryDrive('https://contoso.sharepoint.com/sites/Worship', '%7B243CAEC2-9363-411A-869A-AE66E37BA368%7D'), 'b!from-list-id');
        assert.ok(requested.some((url) => url.includes('/lists/243caec2-9363-411a-869a-ae66e37ba368/drive')));
    });

    it('by its name, or the name used in its web address', async () => {
        fakeGraph({ ...site, ...drives });
        const graph = new GraphClient({}, credential);
        const url = 'https://contoso.sharepoint.com/sites/Worship';
        assert.equal(await graph.findLibraryDrive(url, 'Documents'), 'b!documents');
        assert.equal(await graph.findLibraryDrive(url, 'shared documents'), 'b!documents');
        assert.equal(await graph.findLibraryDrive(url, 'Song Library'), 'b!songs');
        assert.equal(await graph.findLibraryDrive(url, 'SongLibrary'), 'b!songs');
    });

    it('explains what is wrong when the site or library does not exist', async () => {
        fakeGraph({ ...site, ...drives });
        const graph = new GraphClient({}, credential);
        await assert.rejects(graph.findLibraryDrive('https://contoso.sharepoint.com/sites/Worship', 'Hymns'), (err: Error) =>
            err instanceof ConfigError && /Libraries found: "Documents", "Song Library"/.test(err.message),
        );
        await assert.rejects(graph.findLibraryDrive('https://contoso.sharepoint.com/sites/Missing', 'Documents'), (err: Error) =>
            err instanceof ConfigError && /site .* was not found/.test(err.message),
        );
        await assert.rejects(
            graph.findLibraryDrive('https://contoso.sharepoint.com/sites/Worship', '00000000-0000-0000-0000-000000000001'),
            (err: Error) => err instanceof ConfigError && /No document library with id/.test(err.message),
        );
    });
});
