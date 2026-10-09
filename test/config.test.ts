import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import { ConfigError, loadConversionOptions, loadGraphAuthConfig, loadOutputConfig, loadSyncConfig, normalizeFolderPath } from '../src/lib/config.js';
import { encodeDrivePath } from '../src/lib/graph.js';

const SETTINGS = [
    'TENANT_ID',
    'CLIENT_ID',
    'CLIENT_SECRET',
    'MANAGED_IDENTITY_CLIENT_ID',
    'OUTPUT_FOLDER',
    'OUTPUT_DRIVE_ID',
    'MAX_FILE_SIZE_MB',
    'SYNC_DRIVE_ID',
    'SYNC_SOURCE_FOLDER',
    'SYNC_INCLUDE_SUBFOLDERS',
    'SYNC_MAX_FILES_PER_RUN',
    'SYNC_MIN_AGE_SECONDS',
    'SYNC_CONVERT_EXISTING',
    'PRO_FONT_NAME',
    'PRO_FONT_FAMILY',
    'PRO_FONT_SIZE',
    'PRO_GROUP_NAME',
    'INCLUDE_HIDDEN_SLIDES',
];

describe('configuration', () => {
    afterEach(() => {
        for (const name of SETTINGS) {
            delete process.env[name];
        }
    });

    it('uses sensible defaults', () => {
        assert.deepEqual(loadOutputConfig(), { folder: 'proPresenter Files', driveId: undefined, maxFileBytes: 250 * 1024 * 1024 });
        assert.deepEqual(loadConversionOptions(), {
            fontName: 'TimesNewRomanPSMT',
            fontFamily: 'TimesNewRomanPSMT',
            fontSize: 90,
            groupName: 'Group',
            includeHiddenSlides: false,
        });
    });

    it('uses a client secret when one is set, otherwise managed identity', () => {
        assert.deepEqual(loadGraphAuthConfig(), { clientId: undefined });
        process.env.CLIENT_SECRET = 'secret';
        assert.throws(() => loadGraphAuthConfig(), ConfigError);
        process.env.TENANT_ID = 'tenant';
        process.env.CLIENT_ID = 'client';
        assert.deepEqual(loadGraphAuthConfig(), { tenantId: 'tenant', clientId: 'client', clientSecret: 'secret' });
    });

    it('validates numbers and booleans', () => {
        process.env.PRO_FONT_SIZE = 'big';
        assert.throws(() => loadConversionOptions(), /PRO_FONT_SIZE/);
        process.env.PRO_FONT_SIZE = '72';
        process.env.INCLUDE_HIDDEN_SLIDES = 'maybe';
        assert.throws(() => loadConversionOptions(), /INCLUDE_HIDDEN_SLIDES/);
        process.env.INCLUDE_HIDDEN_SLIDES = 'TRUE';
        process.env.PRO_FONT_NAME = 'ArialMT';
        assert.deepEqual(loadConversionOptions(), { fontName: 'ArialMT', fontFamily: 'ArialMT', fontSize: 72, groupName: 'Group', includeHiddenSlides: true });
    });

    it('reads sync settings', () => {
        assert.throws(() => loadSyncConfig(), /SYNC_DRIVE_ID/);
        process.env.SYNC_DRIVE_ID = 'b!abc';
        process.env.SYNC_SOURCE_FOLDER = '/Songs/PowerPoint/';
        process.env.SYNC_INCLUDE_SUBFOLDERS = 'yes';
        assert.deepEqual(loadSyncConfig(), {
            driveId: 'b!abc',
            sourceFolder: 'Songs/PowerPoint',
            includeSubfolders: true,
            maxFilesPerRun: 25,
            minAgeMs: 60_000,
            convertExisting: false,
        });
    });

    it('rejects the library root as the output folder', () => {
        process.env.OUTPUT_FOLDER = '/';
        assert.throws(() => loadOutputConfig(), ConfigError);
    });

    it('normalizes folder paths', () => {
        assert.equal(normalizeFolderPath(' /Songs\\Sunday AM// '), 'Songs/Sunday AM');
        assert.equal(normalizeFolderPath(''), '');
    });

    it('encodes drive paths for Graph', () => {
        assert.equal(encodeDrivePath('proPresenter Files/Song #1 & 50% O\'Neil?.pro'), "proPresenter%20Files/Song%20%231%20%26%2050%25%20O%27Neil%3F.pro");
        assert.equal(encodeDrivePath('/a//b/'), 'a/b');
    });
});
