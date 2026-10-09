import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import type { OutputConfig, SyncConfig } from '../src/lib/config.js';
import { DEFAULT_CONVERSION_OPTIONS, presentationNameFromFileName } from '../src/lib/convert.js';
import { UnsupportedFileError } from '../src/lib/extract/index.js';
import { convertDriveItem, isInFolder, parentFolderPath } from '../src/lib/job.js';
import { syncLibrary } from '../src/lib/sync.js';
import { FakeDrive, silentLogger } from './helpers/fakeDrive.js';
import { decodePresentation, rtfToLines, slideRtfs } from './helpers/pro.js';
import { buildPptx } from './helpers/pptx.js';

const output: OutputConfig = { folder: 'proPresenter Files', maxFileBytes: 50 * 1024 * 1024 };
const options = DEFAULT_CONVERSION_OPTIONS;

function song(...slides: string[][]): Promise<Uint8Array> {
    return buildPptx(slides.map((lines) => ({ shapes: [{ paragraphs: [lines] }] })));
}

function proText(content: Uint8Array | undefined): { name: unknown; slides: string[][] } {
    assert.ok(content, 'expected a .pro file');
    const decoded = decodePresentation(content);
    return { name: decoded.name, slides: slideRtfs(decoded).map(rtfToLines) };
}

describe('convertDriveItem', () => {
    it('converts a file and uploads "<name>.pro" to the output folder', async () => {
        const drive = new FakeDrive();
        const item = drive.add('Songs', 'Amazing Grace.PPTX', await song(['Amazing grace', 'How sweet the sound'], ['Verse 2']));
        const result = await convertDriveItem(drive, drive.driveId, item.id, output, options, silentLogger);

        assert.equal(result.status, 'converted');
        assert.equal(result.status === 'converted' && result.output.name, 'Amazing Grace.pro');
        assert.deepEqual(proText(drive.contentOf('proPresenter Files', 'Amazing Grace.pro')), {
            name: 'Amazing Grace',
            slides: [['Amazing grace', 'How sweet the sound'], ['Verse 2']],
        });
    });

    it('keeps simultaneous conversions separate (no shared temp files)', async () => {
        const drive = new FakeDrive();
        const names = ['Song A', 'Song B', 'Song C', 'Song D', 'Song E'];
        const items = await Promise.all(names.map(async (name) => drive.add('', `${name}.pptx`, await song([`${name} line`]))));
        // Finish downloads in reverse order to force interleaving.
        drive.downloadDelay = (_id, call) => (names.length - call) * 15;

        await Promise.all(items.map((item) => convertDriveItem(drive, drive.driveId, item.id, output, options, silentLogger)));

        for (const name of names) {
            assert.deepEqual(proText(drive.contentOf('proPresenter Files', `${name}.pro`)), { name, slides: [[`${name} line`]] });
        }
    });

    it('converts again when the file changes during conversion, so the newest version wins', async () => {
        const drive = new FakeDrive();
        const item = drive.add('', 'Song.pptx', await song(['Old words']));
        const newer = await song(['New words']);
        drive.onDownload = (id, call) => {
            if (call === 1) {
                drive.edit(id, newer);
            }
        };
        await convertDriveItem(drive, drive.driveId, item.id, output, options, silentLogger);
        assert.equal(drive.uploads.length, 2);
        assert.deepEqual(proText(drive.contentOf('proPresenter Files', 'Song.pro')).slides, [['New words']]);
    });

    it('skips files that are not PowerPoint presentations and files in the output folder', async () => {
        const drive = new FakeDrive();
        const pro = drive.add('proPresenter Files', 'Song.pro', new Uint8Array([1, 2, 3]));
        const inOutput = drive.add('proPresenter Files/Archive', 'Song.pptx', await song(['x']));
        const lockFile = drive.add('', '~$Song.pptx', new Uint8Array([1]));
        for (const item of [pro, inOutput, lockFile]) {
            const result = await convertDriveItem(drive, drive.driveId, item.id, output, options, silentLogger);
            assert.equal(result.status, 'skipped', item.name);
        }
        assert.equal(drive.uploads.length, 0);
    });

    it('fails without uploading when the presentation has no text', async () => {
        const drive = new FakeDrive();
        const item = drive.add('', 'Blank.pptx', await buildPptx([{ shapes: [] }]));
        await assert.rejects(convertDriveItem(drive, drive.driveId, item.id, output, options, silentLogger), UnsupportedFileError);
        assert.equal(drive.uploads.length, 0);
    });

    it('rejects files over the size limit before downloading them', async () => {
        const drive = new FakeDrive();
        const item = drive.add('', 'Huge.pptx', await song(['x']));
        await assert.rejects(
            convertDriveItem(drive, drive.driveId, item.id, { ...output, maxFileBytes: 10 }, options, silentLogger),
            /larger than/,
        );
    });

    it('drops hidden slides unless asked to keep them', async () => {
        const drive = new FakeDrive();
        const pptx = await buildPptx([{ shapes: [{ paragraphs: [['Shown']] }] }, { shapes: [{ paragraphs: [['Hidden']] }], hidden: true }]);
        const item = drive.add('', 'Song.pptx', pptx);
        await convertDriveItem(drive, drive.driveId, item.id, output, options, silentLogger);
        assert.deepEqual(proText(drive.contentOf('proPresenter Files', 'Song.pro')).slides, [['Shown']]);
        await convertDriveItem(drive, drive.driveId, item.id, output, { ...options, includeHiddenSlides: true }, silentLogger);
        assert.deepEqual(proText(drive.contentOf('proPresenter Files', 'Song.pro')).slides, [['Shown'], ['Hidden']]);
    });
});

describe('syncLibrary', () => {
    const config: SyncConfig = { driveId: 'drive-1', sourceFolder: 'Songs', includeSubfolders: false, maxFilesPerRun: 10, minAgeMs: 0 };

    it('converts new files, then only files changed since their .pro was written', async () => {
        const drive = new FakeDrive();
        const a = drive.add('Songs', 'A.pptx', await song(['A1']));
        drive.add('Songs', 'B.ppt.pptx', await song(['B1']));
        drive.add('Songs', 'notes.docx', new Uint8Array([1]));
        drive.add('Other', 'C.pptx', await song(['C1']));

        let summary = await syncLibrary(drive, config, output, options, silentLogger, drive.now() + 60_000);
        assert.deepEqual({ ...summary }, { scanned: 2, pending: 2, converted: 2, failed: 0, deferred: 0 });

        summary = await syncLibrary(drive, config, output, options, silentLogger, drive.now() + 60_000);
        assert.equal(summary.pending, 0, 'nothing changed, nothing to do');

        drive.edit(a.id, await song(['A2']));
        summary = await syncLibrary(drive, config, output, options, silentLogger, drive.now() + 60_000);
        assert.equal(summary.converted, 1);
        assert.deepEqual(proText(drive.contentOf('proPresenter Files', 'A.pro')).slides, [['A2']]);
    });

    it('waits for files that were just modified (may still be open in PowerPoint)', async () => {
        const drive = new FakeDrive();
        drive.add('Songs', 'A.pptx', await song(['A1']));
        const summary = await syncLibrary(drive, { ...config, minAgeMs: 120_000 }, output, options, silentLogger, drive.now() + 1000);
        assert.equal(summary.pending, 0);
        assert.equal(drive.uploads.length, 0);
    });

    it('limits the number of files per run and continues next run', async () => {
        const drive = new FakeDrive();
        for (let i = 0; i < 5; i++) {
            drive.add('Songs', `Song ${i}.pptx`, await song([`Line ${i}`]));
        }
        const first = await syncLibrary(drive, { ...config, maxFilesPerRun: 3 }, output, options, silentLogger, drive.now() + 60_000);
        assert.equal(first.converted, 3);
        assert.equal(first.deferred, 2);
        const second = await syncLibrary(drive, { ...config, maxFilesPerRun: 3 }, output, options, silentLogger, drive.now() + 60_000);
        assert.equal(second.converted, 2);
    });

    it('includes subfolders when enabled, but never the output folder', async () => {
        const drive = new FakeDrive();
        drive.add('', 'Root.pptx', await song(['Root']));
        drive.add('Sub/Deeper', 'Deep.pptx', await song(['Deep']));
        drive.add('proPresenter Files', 'Stray.pptx', await song(['Stray']));
        const summary = await syncLibrary(drive, { ...config, sourceFolder: '', includeSubfolders: true }, output, options, silentLogger, drive.now() + 60_000);
        assert.equal(summary.scanned, 2);
        assert.ok(drive.contentOf('proPresenter Files', 'Deep.pro'));
        assert.equal(drive.contentOf('proPresenter Files', 'Stray.pro'), undefined);
    });

    it('does not retry a file that cannot be converted until it changes', async () => {
        const drive = new FakeDrive();
        const bad = drive.add('Songs', 'Broken.pptx', new TextEncoder().encode('not a presentation'));
        let summary = await syncLibrary(drive, config, output, options, silentLogger, drive.now() + 60_000);
        assert.equal(summary.failed, 1);
        summary = await syncLibrary(drive, config, output, options, silentLogger, drive.now() + 60_000);
        assert.equal(summary.pending, 0);
        drive.edit(bad.id, await song(['Fixed']));
        summary = await syncLibrary(drive, config, output, options, silentLogger, drive.now() + 60_000);
        assert.equal(summary.converted, 1);
    });

    it('refuses a source folder inside the output folder', async () => {
        const drive = new FakeDrive();
        await assert.rejects(syncLibrary(drive, { ...config, sourceFolder: 'proPresenter Files/x' }, output, options, silentLogger));
    });
});

describe('path helpers', () => {
    it('reads the parent folder from parentReference.path', () => {
        assert.equal(parentFolderPath({ id: '1', name: 'a', parentReference: { path: '/drives/b!x/root:' } }), '');
        assert.equal(parentFolderPath({ id: '1', name: 'a', parentReference: { path: '/drives/b!x/root:/Songs/Sunday%20AM' } }), 'Songs/Sunday AM');
        assert.equal(parentFolderPath({ id: '1', name: 'a' }), undefined);
    });

    it('matches folders case-insensitively', () => {
        assert.equal(isInFolder('ProPresenter files', 'proPresenter Files'), true);
        assert.equal(isInFolder('proPresenter Files/Sub', 'proPresenter Files'), true);
        assert.equal(isInFolder('proPresenter Files 2', 'proPresenter Files'), false);
    });

    it('names presentations after the file', () => {
        assert.equal(presentationNameFromFileName('44 Song name.pptx'), '44 Song name');
        assert.equal(presentationNameFromFileName('Song.PPT'), 'Song');
        assert.equal(presentationNameFromFileName('Song.ppt.pptx'), 'Song.ppt');
        assert.equal(presentationNameFromFileName('.pptx'), 'Untitled');
    });
});
