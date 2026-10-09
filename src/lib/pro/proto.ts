import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import protobuf from 'protobufjs';

let presentationType: protobuf.Type | undefined;

/** Locates the proto/ folder whether running from src/ (tests via tsc output) or dist/. */
function findProtoDir(): string {
    let dir = path.dirname(fileURLToPath(import.meta.url));
    for (let i = 0; i < 6; i++) {
        const candidate = path.join(dir, 'proto');
        if (existsSync(path.join(candidate, 'presentation.proto'))) {
            return candidate;
        }
        dir = path.dirname(dir);
    }
    throw new Error('Could not find the proto/ folder containing presentation.proto.');
}

/** The ProPresenter 7 rv.data.Presentation message type, loaded once from the .proto files. */
export function getPresentationType(): protobuf.Type {
    if (!presentationType) {
        const protoDir = findProtoDir();
        const root = new protobuf.Root();
        root.resolvePath = (_origin, target) => (path.isAbsolute(target) ? target : path.join(protoDir, target));
        root.loadSync('presentation.proto', { keepCase: true });
        presentationType = root.lookupType('rv.data.Presentation');
    }
    return presentationType;
}
