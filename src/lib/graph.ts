import { ClientSecretCredential, ManagedIdentityCredential, type TokenCredential } from '@azure/identity';
import type { GraphAuthConfig } from './config.js';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MAX_ATTEMPTS = 4;
const REQUEST_TIMEOUT_MS = 60_000;
const MAX_RETRY_DELAY_MS = 30_000;

export interface DriveItem {
    id: string;
    name: string;
    /** Changes only when the file content changes (not on rename/metadata edits). */
    cTag?: string;
    eTag?: string;
    size?: number;
    lastModifiedDateTime?: string;
    webUrl?: string;
    file?: { mimeType?: string };
    folder?: { childCount?: number };
    deleted?: unknown;
    parentReference?: { driveId?: string; id?: string; path?: string };
}

/** A Graph request that failed; `status` is the HTTP status (0 for network errors). */
export class GraphError extends Error {
    override name = 'GraphError';
    constructor(
        message: string,
        readonly status: number,
        readonly code?: string,
    ) {
        super(message);
    }
}

export interface Logger {
    log(...args: unknown[]): void;
    warn(...args: unknown[]): void;
    error(...args: unknown[]): void;
}

const ITEM_SELECT = 'id,name,cTag,eTag,size,lastModifiedDateTime,webUrl,file,folder,deleted,parentReference';

/** Encodes a folder/file path for Graph path-based addressing (root:/path:), one segment at a time. */
export function encodeDrivePath(path: string): string {
    return path
        .split('/')
        .filter(Boolean)
        .map((segment) => encodeURIComponent(segment).replace(/'/g, '%27'))
        .join('/');
}

/** The file operations the conversion jobs need (implemented by GraphClient; faked in tests). */
export interface DriveApi {
    getItem(driveId: string, itemId: string): Promise<DriveItem>;
    download(driveId: string, itemId: string): Promise<Uint8Array>;
    upload(driveId: string, folder: string, fileName: string, content: Uint8Array): Promise<DriveItem>;
    listChildren(driveId: string, folder: string): Promise<DriveItem[]>;
}

/** Minimal Microsoft Graph client for OneDrive/SharePoint files, with retries for throttling and transient errors. */
export class GraphClient implements DriveApi {
    private readonly credential: TokenCredential;

    constructor(auth: GraphAuthConfig) {
        if (auth.clientSecret && auth.tenantId && auth.clientId) {
            this.credential = new ClientSecretCredential(auth.tenantId, auth.clientId, auth.clientSecret);
        } else if (auth.clientId) {
            this.credential = new ManagedIdentityCredential({ clientId: auth.clientId });
        } else {
            this.credential = new ManagedIdentityCredential();
        }
    }

    async getItem(driveId: string, itemId: string): Promise<DriveItem> {
        return this.json<DriveItem>(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}?$select=${ITEM_SELECT}`);
    }

    /** Downloads a file's content. Graph redirects to a pre-authenticated download URL. */
    async download(driveId: string, itemId: string): Promise<Uint8Array> {
        const response = await this.request(`/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`);
        return new Uint8Array(await response.arrayBuffer());
    }

    /**
     * Uploads (creating or replacing) a file at folder/fileName in the drive.
     * Missing folders in the path are created by Graph.
     */
    async upload(driveId: string, folder: string, fileName: string, content: Uint8Array): Promise<DriveItem> {
        const path = encodeDrivePath(`${folder}/${fileName}`);
        const response = await this.request(`/drives/${encodeURIComponent(driveId)}/root:/${path}:/content`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/octet-stream' },
            body: content,
        });
        return (await response.json()) as DriveItem;
    }

    /** Lists the children of a folder (path within the drive; "" for the root). Returns [] if the folder does not exist. */
    async listChildren(driveId: string, folder: string): Promise<DriveItem[]> {
        const base = `/drives/${encodeURIComponent(driveId)}`;
        const location = folder ? `${base}/root:/${encodeDrivePath(folder)}:/children` : `${base}/root/children`;
        let url: string | undefined = `${location}?$select=${ITEM_SELECT}&$top=200`;
        const items: DriveItem[] = [];
        try {
            while (url) {
                const page: { value: DriveItem[]; '@odata.nextLink'?: string } = await this.json(url);
                items.push(...page.value);
                url = page['@odata.nextLink'];
            }
        } catch (err) {
            if (err instanceof GraphError && err.status === 404) {
                return [];
            }
            throw err;
        }
        return items;
    }

    private async json<T>(pathOrUrl: string): Promise<T> {
        const response = await this.request(pathOrUrl);
        return (await response.json()) as T;
    }

    private async request(pathOrUrl: string, init: RequestInit = {}): Promise<Response> {
        const url = pathOrUrl.startsWith('https://') ? pathOrUrl : `${GRAPH_BASE}${pathOrUrl}`;
        for (let attempt = 1; ; attempt++) {
            let response: Response;
            try {
                const token = await this.credential.getToken(GRAPH_SCOPE);
                if (!token) {
                    throw new GraphError('Could not get an access token for Microsoft Graph.', 401);
                }
                const headers = new Headers(init.headers);
                headers.set('Authorization', `Bearer ${token.token}`);
                response = await fetch(url, { ...init, headers, signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
            } catch (err) {
                if (err instanceof GraphError) {
                    throw err;
                }
                // Network failure or timeout.
                if (attempt >= MAX_ATTEMPTS) {
                    throw new GraphError(`Microsoft Graph request failed: ${(err as Error).message}`, 0);
                }
                await delay(backoff(attempt));
                continue;
            }

            if (response.ok) {
                return response;
            }

            const retryable = response.status === 429 || response.status === 502 || response.status === 503 || response.status === 504;
            if (retryable && attempt < MAX_ATTEMPTS) {
                await response.body?.cancel();
                await delay(retryAfter(response) ?? backoff(attempt));
                continue;
            }
            throw await toGraphError(response);
        }
    }
}

async function toGraphError(response: Response): Promise<GraphError> {
    let code: string | undefined;
    let message = response.statusText;
    try {
        const body = (await response.json()) as { error?: { code?: string; message?: string } };
        code = body.error?.code;
        message = body.error?.message ?? message;
    } catch {
        // Body was not JSON.
    }
    return new GraphError(`Microsoft Graph returned ${response.status}${code ? ` ${code}` : ''}: ${message}`, response.status, code);
}

function retryAfter(response: Response): number | undefined {
    const header = response.headers.get('Retry-After');
    if (!header) {
        return undefined;
    }
    const seconds = Number(header);
    const ms = Number.isFinite(seconds) ? seconds * 1000 : Date.parse(header) - Date.now();
    return Number.isFinite(ms) ? Math.min(Math.max(ms, 0), MAX_RETRY_DELAY_MS) : undefined;
}

function backoff(attempt: number): number {
    return Math.min(1000 * 2 ** (attempt - 1), MAX_RETRY_DELAY_MS);
}

function delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
}
