import { loadConversionOptions, loadGraphAuthConfig, loadOutputConfig, type OutputConfig } from './config.js';
import type { ConversionOptions } from './convert.js';
import { GraphClient } from './graph.js';

// Shared, lazily created objects. Creating them on first use (rather than at import)
// means a missing setting is reported in the function's logs and response instead of
// stopping the whole Function App from starting.

let graph: GraphClient | undefined;

export function getGraphClient(): GraphClient {
    graph ??= new GraphClient(loadGraphAuthConfig());
    return graph;
}

export function getJobSettings(): { output: OutputConfig; options: ConversionOptions } {
    return { output: loadOutputConfig(), options: loadConversionOptions() };
}
