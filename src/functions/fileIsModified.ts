import { app, type HttpRequest, type HttpResponseInit, type InvocationContext } from '@azure/functions';
import { ConfigError } from '../lib/config.js';
import { UnsupportedFileError } from '../lib/extract/index.js';
import { GraphError } from '../lib/graph.js';
import { convertDriveItem } from '../lib/job.js';
import { getGraphClient, getJobSettings } from '../lib/runtime.js';

// HTTP endpoint for the Power Automate flow: POST /api/fileIsModified?code=<function key>
// Body: { "driveId": "<document library id>", "driveItemId": "<file id>" }

interface ConvertRequest {
    driveId: string;
    driveItemId: string;
}

function parseRequest(body: unknown): ConvertRequest | string {
    if (typeof body !== 'object' || body === null) {
        return 'Request body must be a JSON object with "driveId" and "driveItemId".';
    }
    const { driveId, driveItemId, itemId } = body as Record<string, unknown>;
    const id = driveItemId ?? itemId;
    if (typeof driveId !== 'string' || !driveId.trim()) {
        return 'Request body is missing "driveId".';
    }
    if (typeof id !== 'string' || !id.trim()) {
        return 'Request body is missing "driveItemId".';
    }
    return { driveId: driveId.trim(), driveItemId: id.trim() };
}

export async function fileIsModified(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    let body: unknown;
    try {
        body = await request.json();
    } catch {
        return { status: 400, jsonBody: { status: 'error', error: 'Request body is not valid JSON.' } };
    }
    const parsed = parseRequest(body);
    if (typeof parsed === 'string') {
        return { status: 400, jsonBody: { status: 'error', error: parsed } };
    }
    context.log(`Conversion requested for drive ${parsed.driveId}, item ${parsed.driveItemId}.`);

    try {
        const { output, options } = getJobSettings();
        const result = await convertDriveItem(getGraphClient(), parsed.driveId, parsed.driveItemId, output, options, context);
        if (result.status === 'skipped') {
            context.log(`Skipped "${result.source}": ${result.reason}`);
        }
        return { status: 200, jsonBody: result };
    } catch (err) {
        return errorResponse(err, context);
    }
}

/**
 * Maps failures to status codes Power Automate understands: 4xx for problems a retry
 * cannot fix, 5xx (which Power Automate retries) for temporary ones.
 */
function errorResponse(err: unknown, context: InvocationContext): HttpResponseInit {
    const message = err instanceof Error ? err.message : String(err);
    if (err instanceof UnsupportedFileError) {
        context.warn(`Could not convert file: ${message}`);
        return { status: 422, jsonBody: { status: 'error', error: message } };
    }
    if (err instanceof GraphError && err.status === 404) {
        context.warn(message);
        return { status: 404, jsonBody: { status: 'error', error: 'File not found. It may have been deleted or moved.' } };
    }
    if (err instanceof ConfigError) {
        context.error(message);
        return { status: 500, jsonBody: { status: 'error', error: message } };
    }
    if (err instanceof GraphError && (err.status === 401 || err.status === 403)) {
        context.error(`${message} Check the app registration's permissions on the SharePoint site.`);
        return { status: 500, jsonBody: { status: 'error', error: 'The Function App is not allowed to access this file. See the logs for details.' } };
    }
    context.error('Conversion failed:', err);
    const status = err instanceof GraphError && err.status >= 500 ? 503 : 500;
    return { status, jsonBody: { status: 'error', error: message } };
}

app.http('fileIsModified', {
    methods: ['POST'],
    authLevel: 'function',
    handler: fileIsModified,
});
