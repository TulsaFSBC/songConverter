import assert from 'node:assert/strict';
import { afterEach, describe, it } from 'node:test';
import {
    asGuid,
    ConfigError,
    isSyncEnabled,
    loadConversionOptions,
    loadGraphAuthConfig,
    loadOutputConfig,
    loadSyncConfig,
    loadSyncLibrary,
    normalizeFolderPath,
} from '../src/lib/config.js';
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
    'SYNC_SITE_URL',
    'SYNC_LIBRARY',
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
        process.env.SYNC_SOURCE_FOLDER = '/Songs/PowerPoint/';
        process.env.SYNC_INCLUDE_SUBFOLDERS = 'yes';
        assert.deepEqual(loadSyncConfig('b!abc'), {
            driveId: 'b!abc',
            sourceFolder: 'Songs/PowerPoint',
            includeSubfolders: true,
            maxFilesPerRun: 25,
            minAgeMs: 60_000,
            convertExisting: false,
        });
    });

    it('identifies the synced library by drive id, or by site URL plus library name or id', () => {
        assert.equal(isSyncEnabled(), false);
        assert.throws(() => loadSyncLibrary(), /SYNC_SITE_URL/);

        process.env.SYNC_DRIVE_ID = 'b!abc';
        assert.equal(isSyncEnabled(), true);
        assert.deepEqual(loadSyncLibrary(), { driveId: 'b!abc' });

        // The library id from SharePoint's library settings page is not a drive id.
        process.env.SYNC_DRIVE_ID = '243caec2-9363-411a-869a-ae66e37ba368';
        assert.throws(() => loadSyncLibrary(), /library id, not a Microsoft Graph drive id.*SYNC_SITE_URL/s);
        process.env.SYNC_SITE_URL = 'https://contoso.sharepoint.com/sites/Worship';
        assert.deepEqual(loadSyncLibrary(), { siteUrl: 'https://contoso.sharepoint.com/sites/Worship', library: '243caec2-9363-411a-869a-ae66e37ba368' });

        delete process.env.SYNC_DRIVE_ID;
        assert.deepEqual(loadSyncLibrary(), { siteUrl: 'https://contoso.sharepoint.com/sites/Worship', library: 'Documents' });
        process.env.SYNC_LIBRARY = 'Song Library';
        assert.deepEqual(loadSyncLibrary(), { siteUrl: 'https://contoso.sharepoint.com/sites/Worship', library: 'Song Library' });
    });

    it('recognizes library ids in the forms SharePoint shows them', () => {
        for (const value of ['243caec2-9363-411a-869a-ae66e37ba368', '{243CAEC2-9363-411A-869A-AE66E37BA368}', '%7B243CAEC2-9363-411A-869A-AE66E37BA368%7D']) {
            assert.equal(asGuid(value), '243caec2-9363-411a-869a-ae66e37ba368', value);
        }
        assert.equal(asGuid('b!AbC123'), undefined);
        assert.equal(asGuid('Documents'), undefined);
    });

    it('rejects the library root as the output folder', () => {
        process.env.OUTPUT_FOLDER = '/';
        assert.throws(() => loadOutputConfig(), ConfigError);
    });

    it('only allows a single top-level output folder, never a subfolder path', () => {
        for (const value of ['proPresenter Files/Sunday', 'Songs/proPresenter Files', 'a\\b']) {
            process.env.OUTPUT_FOLDER = value;
            assert.throws(() => loadOutputConfig(), /single folder/, value);
        }
        process.env.OUTPUT_FOLDER = '/proPresenter Files/';
        assert.equal(loadOutputConfig().folder, 'proPresenter Files');
    });

    it('puts every .pro file in one library: OUTPUT_DRIVE_ID, else the synced library', () => {
        assert.equal(loadOutputConfig().driveId, undefined);
        assert.equal(loadOutputConfig('b!songs').driveId, 'b!songs');
        process.env.OUTPUT_DRIVE_ID = 'b!output';
        assert.equal(loadOutputConfig('b!songs').driveId, 'b!output');
        process.env.OUTPUT_DRIVE_ID = '243caec2-9363-411a-869a-ae66e37ba368';
        assert.throws(() => loadOutputConfig('b!songs'), /OUTPUT_DRIVE_ID/);
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
