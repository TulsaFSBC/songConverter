// Converts PowerPoint files on this computer without Azure or SharePoint.
// Usage: npm run build && npm run convert -- <file.pptx|file.ppt>... [--out <folder>]
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { convertPowerPoint } from '../src/lib/convert.js';
import { loadConversionOptions } from '../src/lib/config.js';

async function main(): Promise<void> {
    const args = process.argv.slice(2);
    const outIndex = args.indexOf('--out');
    const outDir = outIndex >= 0 ? args[outIndex + 1] : undefined;
    const files = args.filter((_, i) => outIndex < 0 || (i !== outIndex && i !== outIndex + 1));
    if (files.length === 0) {
        console.error('Usage: npm run convert -- <file.pptx|file.ppt>... [--out <folder>]');
        process.exitCode = 2;
        return;
    }

    for (const file of files) {
        try {
            const result = await convertPowerPoint(await readFile(file), path.basename(file), loadConversionOptions());
            const target = path.join(outDir ?? path.dirname(file), result.fileName);
            await mkdir(path.dirname(target), { recursive: true });
            await writeFile(target, result.content);
            console.log(`${file} -> ${target} (${result.slideCount} slides)`);
            result.slides.forEach((lines, i) => {
                console.log(`  Slide ${i + 1}:`);
                for (const line of lines) {
                    console.log(`    ${line}`);
                }
            });
        } catch (err) {
            console.error(`${file}: ${(err as Error).message}`);
            process.exitCode = 1;
        }
    }
}

await main();
