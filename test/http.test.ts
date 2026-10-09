import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import functions from '@azure/functions';
import { fileIsModified } from '../src/functions/fileIsModified.js';

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
