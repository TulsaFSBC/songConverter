import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import type { OutputConfig, SyncConfig } from '../src/lib/config.js';
import { DEFAULT_CONVERSION_OPTIONS, presentationNameFromFileName, proFileName } from '../src/lib/convert.js';
import { UnsupportedFileError } from '../src/lib/extract/index.js';
import { convertDriveItem, isInFolder, parentFolderPath } from '../src/lib/job.js';
import { GraphError } from '../src/lib/graph.js';
import { MemoryStateStore } from '../src/lib/state.js';
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
    const config: SyncConfig = { driveId: 'drive-1', sourceFolder: 'Songs', includeSubfolders: false, maxFilesPerRun: 10, minAgeMs: 0, convertExisting: false };
    const later = (drive: FakeDrive) => drive.now() + 60_000;
    const newDrive = () => new FakeDrive().addFolder('Songs');
    const run = (drive: FakeDrive, store: MemoryStateStore, overrides: Partial<SyncConfig> = {}, now = later(drive)) =>
        syncLibrary(drive, store, { ...config, ...overrides }, output, options, silentLogger, now);

    it('first run only records existing files; later runs convert files added since the previous run', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        drive.add('Songs', 'Old 1.pptx', await song(['Old 1']));
        drive.add('Songs', 'Old 2.ppt.pptx', await song(['Old 2']));
        drive.add('Songs', 'notes.docx', new Uint8Array([1]));
        drive.add('Other', 'Elsewhere.pptx', await song(['Elsewhere']));

        const first = await run(drive, store);
        assert.deepEqual({ ...first }, { initialized: true, scanned: 2, added: 0, changed: 0, converted: 0, failed: 0, deferred: 0 });
        assert.equal(drive.uploads.length, 0);

        drive.add('Songs', 'New.pptx', await song(['New song']));
        const second = await run(drive, store);
        assert.deepEqual({ ...second }, { initialized: false, scanned: 3, added: 1, changed: 0, converted: 1, failed: 0, deferred: 0 });
        assert.deepEqual(drive.uploads.map((u) => u.name), ['New.pro']);

        const third = await run(drive, store);
        assert.equal(third.added + third.changed, 0, 'nothing new, nothing to do');
        assert.equal(drive.uploads.length, 1);
    });

    it('converts files whose content changed, but not metadata-only changes', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        const a = drive.add('Songs', 'A.pptx', await song(['A1']));
        const b = drive.add('Songs', 'B.pptx', await song(['B1']));
        await run(drive, store);

        drive.edit(a.id, await song(['A2']));
        drive.touch(b.id);
        const summary = await run(drive, store);
        assert.equal(summary.changed, 1);
        assert.deepEqual(drive.uploads.map((u) => u.name), ['A.pro']);
        assert.deepEqual(proText(drive.contentOf('proPresenter Files', 'A.pro')).slides, [['A2']]);
    });

    it('treats a file moved into the folder as added, even though its dates are old', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        const moved = drive.add('Archive', 'Moved.pptx', await song(['Moved']));
        await run(drive, store);
        drive.move(moved.id, 'Songs');
        const summary = await run(drive, store);
        assert.equal(summary.added, 1);
        assert.ok(drive.contentOf('proPresenter Files', 'Moved.pro'));
    });

    it('waits until a new file has not changed for SYNC_MIN_AGE_SECONDS', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        await run(drive, store);
        drive.add('Songs', 'Editing.pptx', await song(['Draft']));
        const soon = await run(drive, store, { minAgeMs: 120_000 }, drive.now() + 1000);
        assert.deepEqual([soon.added, soon.converted, soon.deferred], [1, 0, 1]);
        const afterwards = await run(drive, store, { minAgeMs: 120_000 }, drive.now() + 300_000);
        assert.equal(afterwards.converted, 1);
    });

    it('limits the number of files per run and continues next run', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        await run(drive, store);
        for (let i = 0; i < 5; i++) {
            drive.add('Songs', `Song ${i}.pptx`, await song([`Line ${i}`]));
        }
        const first = await run(drive, store, { maxFilesPerRun: 3 });
        assert.deepEqual([first.converted, first.deferred], [3, 2]);
        const second = await run(drive, store, { maxFilesPerRun: 3 });
        assert.deepEqual([second.added, second.converted], [2, 2]);
        assert.equal(drive.uploads.length, 5);
    });

    it('does not retry a file that cannot be converted until it changes', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        await run(drive, store);
        const bad = drive.add('Songs', 'Broken.pptx', new TextEncoder().encode('not a presentation'));
        assert.equal((await run(drive, store)).failed, 1);
        assert.equal((await run(drive, store)).failed, 0, 'not retried');
        drive.edit(bad.id, await song(['Fixed']));
        assert.equal((await run(drive, store)).converted, 1);
    });

    it('retries temporary failures on the next run, and gives up after a day of failures', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        await run(drive, store);
        const flaky = drive.add('Songs', 'Flaky.pptx', await song(['Flaky']));
        let failing = true;
        drive.downloadError = () => (failing ? new GraphError('Microsoft Graph returned 503', 503) : undefined);

        const start = later(drive);
        assert.equal((await run(drive, store, {}, start)).failed, 1);
        assert.equal((await run(drive, store, {}, start + 5 * 60_000)).failed, 1, 'retried');
        failing = false;
        assert.equal((await run(drive, store, {}, start + 10 * 60_000)).converted, 1, 'succeeds once SharePoint recovers');

        drive.edit(flaky.id, await song(['Flaky 2']));
        failing = true;
        const t = later(drive);
        assert.equal((await run(drive, store, {}, t)).failed, 1);
        assert.equal((await run(drive, store, {}, t + 25 * 60 * 60_000)).failed, 1);
        assert.equal((await run(drive, store, {}, t + 26 * 60 * 60_000)).failed, 0, 'gave up after 24 hours of failures');
    });

    it('forgets deleted files and starts a new baseline when the watched folder changes', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        const gone = drive.add('Songs', 'Gone.pptx', await song(['Gone']));
        drive.add('Other', 'Other.pptx', await song(['Other']));
        await run(drive, store);
        drive.delete(gone.id);
        await run(drive, store);
        assert.equal(Object.keys(store.state?.files ?? {}).length, 0);

        const switched = await run(drive, store, { sourceFolder: 'Other' });
        assert.equal(switched.initialized, true);
        assert.equal(drive.uploads.length, 0);
    });

    it('with SYNC_CONVERT_EXISTING, the first run converts existing files that have no up-to-date .pro', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        drive.add('Songs', 'Done.pptx', await song(['Done']));
        drive.add('proPresenter Files', 'Done.pro', new Uint8Array([1]));
        drive.add('Songs', 'Missing.pptx', await song(['Missing']));
        const summary = await run(drive, store, { convertExisting: true });
        assert.deepEqual([summary.initialized, summary.added, summary.converted], [true, 1, 1]);
        assert.deepEqual(drive.uploads.map((u) => u.name), ['Missing.pro']);
    });

    it('includes subfolders when enabled, but never the output folder', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        await run(drive, store, { sourceFolder: '', includeSubfolders: true });
        drive.add('', 'Root.pptx', await song(['Root']));
        drive.add('Sub/Deeper', 'Deep.pptx', await song(['Deep']));
        drive.add('proPresenter Files', 'Stray.pptx', await song(['Stray']));
        const summary = await run(drive, store, { sourceFolder: '', includeSubfolders: true });
        assert.deepEqual([summary.scanned, summary.converted], [2, 2]);
        assert.ok(drive.contentOf('proPresenter Files', 'Deep.pro'));
        assert.equal(drive.contentOf('proPresenter Files', 'Stray.pro'), undefined);
    });

    it('never opens the proPresenter Files folder when it sits inside the watched library', async () => {
        for (const includeSubfolders of [false, true]) {
            const drive = new FakeDrive();
            const store = new MemoryStateStore();
            drive.add('', 'Song.pptx', await song(['Song']));
            drive.add('Sunday', 'Other.pptx', await song(['Other']));
            drive.add('ProPresenter Files', 'Old.pro', new Uint8Array([1])); // different case, same folder
            drive.add('ProPresenter Files', 'Stray.pptx', await song(['Stray']));
            await run(drive, store, { sourceFolder: '', includeSubfolders });
            drive.add('', 'New.pptx', await song(['New']));
            await run(drive, store, { sourceFolder: '', includeSubfolders });

            assert.ok(drive.listed.length > 0);
            assert.ok(!drive.listed.some((folder) => folder.toLowerCase().startsWith('propresenter files')), `listed: ${drive.listed.join(', ')}`);
            assert.deepEqual(drive.uploads.map((u) => u.name), ['New.pro'], `includeSubfolders=${includeSubfolders}`);
        }
    });

    it('reports a source folder that does not exist instead of silently finding nothing', async () => {
        const drive = newDrive();
        await assert.rejects(run(drive, new MemoryStateStore(), { sourceFolder: 'Shared Documents/Songs' }), /SYNC_SOURCE_FOLDER "Shared Documents\/Songs" was not found/);
    });

    it('with SYNC_CONVERT_EXISTING, regenerates empty .pro files left by the old converter', async () => {
        const drive = newDrive();
        const store = new MemoryStateStore();
        drive.add('Songs', 'Empty.pptx', await song(['Words']));
        drive.add('proPresenter Files', 'Empty.pro', new Uint8Array(0));
        const summary = await run(drive, store, { convertExisting: true });
        assert.equal(summary.converted, 1);
        assert.ok((drive.contentOf('proPresenter Files', 'Empty.pro')?.length ?? 0) > 0);
    });

    it('refuses a source folder inside the output folder', async () => {
        const drive = newDrive();
        await assert.rejects(run(drive, new MemoryStateStore(), { sourceFolder: 'proPresenter Files/x' }));
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

    it('never produces a .pro file name that points into a subfolder', () => {
        assert.equal(proFileName('Song'), 'Song.pro');
        assert.equal(proFileName('AC/DC \\ Hymn'), 'AC-DC - Hymn.pro');
    });

    it('names presentations after the file', () => {
        assert.equal(presentationNameFromFileName('44 Song name.pptx'), '44 Song name');
        assert.equal(presentationNameFromFileName('Song.PPT'), 'Song');
        assert.equal(presentationNameFromFileName('Song.ppt.pptx'), 'Song.ppt');
        assert.equal(presentationNameFromFileName('.pptx'), 'Untitled');
    });
});
