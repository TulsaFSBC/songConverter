import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import functions from '@azure/functions';
import { fileIsModified, parseRequest } from '../src/functions/fileIsModified.js';
import { GraphError, sitePathFromUrl } from '../src/lib/graph.js';

// @azure/functions is CommonJS; its classes are only reachable through the default export from ESM.
const { HttpRequest, InvocationContext } = functions;

function post(body: string): InstanceType<typeof HttpRequest> {
    return new HttpRequest({ method: 'POST', url: 'http://localhost/api/fileIsModified', body: { string: body } });
}

const context = () => new InvocationContext({ functionName: 'fileIsModified', invocationId: 'test', logHandler: () => {} });

describe('fileIsModified HTTP function', () => {
    it('returns 400 for a body that is not JSON', async () => {
        const response = await fileIsModified(post('not json'), context());
        assert.equal(response.status, 400);
    });

    it('returns 400 when driveId or driveItemId is missing', async () => {
        for (const body of ['{}', '{"driveId":"d"}', '{"driveItemId":"i"}', '[]', '"text"', '{"driveId":"","driveItemId":"i"}']) {
            const response = await fileIsModified(post(body), context());
            assert.equal(response.status, 400, body);
        }
    });

    it('returns 500 with a clear message when the app is not configured', async () => {
        delete process.env.CLIENT_SECRET;
        process.env.OUTPUT_FOLDER = '/';
        try {
            const response = await fileIsModified(post('{"driveId":"d","driveItemId":"i"}'), context());
            assert.equal(response.status, 500);
            assert.match(JSON.stringify(response.jsonBody), /OUTPUT_FOLDER/);
        } finally {
            delete process.env.OUTPUT_FOLDER;
        }
    });
});

describe('request formats', () => {
    it('accepts drive item ids or SharePoint list item ids', () => {
        assert.deepEqual(parseRequest({ driveId: ' b!x ', driveItemId: '01ABC' }), { driveId: 'b!x', driveItemId: '01ABC' });
        assert.deepEqual(parseRequest({ siteUrl: 'https://contoso.sharepoint.com/sites/Worship', listId: '{6f3c}', listItemId: 12 }), {
            siteUrl: 'https://contoso.sharepoint.com/sites/Worship',
            listId: '{6f3c}',
            listItemId: '12',
        });
        assert.equal(typeof parseRequest({ siteUrl: 'https://contoso.sharepoint.com', listItemId: 12 }), 'string');
        assert.equal(typeof parseRequest({ driveId: 'b!x', driveItemId: 5.5 }), 'string');
    });

    it('turns site URLs into Graph site paths', () => {
        assert.equal(sitePathFromUrl('https://contoso.sharepoint.com/sites/Worship/'), 'contoso.sharepoint.com:/sites/Worship');
        assert.equal(sitePathFromUrl('https://contoso.sharepoint.com/sites/Worship%20Team'), 'contoso.sharepoint.com:/sites/Worship%20Team');
        assert.equal(sitePathFromUrl('https://contoso.sharepoint.com/sites/Worship Team'), 'contoso.sharepoint.com:/sites/Worship%20Team');
        assert.equal(sitePathFromUrl('https://contoso.sharepoint.com'), 'contoso.sharepoint.com');
        assert.equal(sitePathFromUrl('https://contoso.sharepoint.com/sites/Worship/SitePages/Home.aspx'), 'contoso.sharepoint.com:/sites/Worship');
        assert.equal(
            sitePathFromUrl('https://contoso.sharepoint.com/sites/Worship/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FWorship'),
            'contoso.sharepoint.com:/sites/Worship',
        );
        assert.equal(sitePathFromUrl(' https://contoso.sharepoint.com/teams/Music/_layouts/15/viewlsts.aspx '), 'contoso.sharepoint.com:/teams/Music');
        assert.equal(sitePathFromUrl('https://contoso.sharepoint.com/SitePages/Home.aspx'), 'contoso.sharepoint.com');
        assert.throws(() => sitePathFromUrl('not a url'), GraphError);
    });
});
