import { ManagedIdentityCredential } from '@azure/identity';
import { BlobServiceClient, RestError, type ContainerClient } from '@azure/storage-blob';
import { ConfigError } from './config.js';

/** What the library sync remembers between runs. */
export interface SyncState {
    version: 1;
    /** Identifies the watched folder; watching a different folder starts over. */
    source: string;
    /** Content version (cTag) of every file the sync has handled, by drive item id. */
    files: Record<string, string>;
    /** When each file that keeps failing for temporary reasons first failed (ISO time), by drive item id. */
    failing?: Record<string, string>;
    updated: string;
}

export interface StateStore {
    load(): Promise<SyncState | undefined>;
    save(state: SyncState): Promise<void>;
}

const CONTAINER = 'songconverter';
const BLOB = 'sync-state.json';

/**
 * Keeps the sync state as a small JSON blob in the Function App's own storage account
 * (the AzureWebJobsStorage account every Function App already has), in the
 * "songconverter" container.
 */
export class BlobStateStore implements StateStore {
    constructor(private readonly container: ContainerClient) {}

    /** Connects with the same settings the Functions host uses: a connection string or an identity-based connection. */
    static fromEnvironment(env: NodeJS.ProcessEnv = process.env): BlobStateStore {
        const connectionString = env.AzureWebJobsStorage?.trim();
        let service: BlobServiceClient;
        if (connectionString) {
            service = BlobServiceClient.fromConnectionString(connectionString);
        } else {
            const accountName = env.AzureWebJobsStorage__accountName?.trim();
            const serviceUri = env.AzureWebJobsStorage__blobServiceUri?.trim() || (accountName ? `https://${accountName}.blob.core.windows.net` : undefined);
            if (!serviceUri) {
                throw new ConfigError('Library sync needs the AzureWebJobsStorage setting (or AzureWebJobsStorage__accountName) to remember which files it has seen.');
            }
            const clientId = env.AzureWebJobsStorage__clientId?.trim();
            service = new BlobServiceClient(serviceUri, clientId ? new ManagedIdentityCredential({ clientId }) : new ManagedIdentityCredential());
        }
        return new BlobStateStore(service.getContainerClient(CONTAINER));
    }

    async load(): Promise<SyncState | undefined> {
        let text: string;
        try {
            const buffer = await this.container.getBlockBlobClient(BLOB).downloadToBuffer();
            text = buffer.toString('utf8');
        } catch (err) {
            if (err instanceof RestError && err.statusCode === 404) {
                return undefined;
            }
            throw err;
        }
        const state = JSON.parse(text) as Partial<SyncState>;
        if (state.version !== 1 || typeof state.source !== 'string' || typeof state.files !== 'object' || state.files === null) {
            return undefined;
        }
        return state as SyncState;
    }

    async save(state: SyncState): Promise<void> {
        await this.container.createIfNotExists();
        const body = JSON.stringify(state);
        await this.container.getBlockBlobClient(BLOB).upload(body, Buffer.byteLength(body), {
            blobHTTPHeaders: { blobContentType: 'application/json' },
        });
    }
}

/** In-memory store, for tests and local experiments. */
export class MemoryStateStore implements StateStore {
    state: SyncState | undefined;

    async load(): Promise<SyncState | undefined> {
        return this.state && structuredClone(this.state);
    }

    async save(state: SyncState): Promise<void> {
        this.state = structuredClone(state);
    }
}
