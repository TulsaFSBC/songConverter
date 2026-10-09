# Song Converter

Converts song lyric PowerPoint files (`.pptx`, and the older `.ppt`) saved in SharePoint into
ProPresenter 7 presentations (`.pro`), and saves them back to SharePoint.

```
SharePoint library                 Azure Function App                         SharePoint library
  Songs/Amazing Grace.pptx  ──►  download → read slide text → build .pro  ──►  proPresenter Files/Amazing Grace.pro
                                  (all in memory, nothing written to disk)
```

Each PowerPoint slide becomes one ProPresenter slide, with each line of the slide as its own line
(white, centered, 90 pt Times New Roman by default; change it in ProPresenter or with the settings below).
The presentation is named after the PowerPoint file.

The app can be triggered two ways — use either or both:

| | **Library sync** (recommended) | **Power Automate flow** |
|---|---|---|
| How | A timer in the Function App checks a SharePoint folder every 5 minutes and converts new or changed files | A flow calls the function each time a file changes |
| Needs | Nothing else | Power Automate with the premium **HTTP** action |
| Turn on | Set `SYNC_DRIVE_ID` | Create the flow (below) |

## Setting up in Azure

### 1. Register an app so the function can read and write SharePoint

1. In the [Microsoft Entra admin center](https://entra.microsoft.com) go to **App registrations → New registration**.
   Name it `SongConverter`, leave the other options, and register.
2. Copy the **Application (client) ID** and **Directory (tenant) ID** from the overview page.
3. **Certificates & secrets → New client secret.** Copy the secret **Value** (it is shown only once).
   Note the expiry date: the app stops working when the secret expires, so put a reminder in your calendar.
4. **API permissions → Add a permission → Microsoft Graph → Application permissions**, then either:
   - **`Sites.Selected`** (recommended — access to only your song site; needs one extra step, below), or
   - **`Sites.ReadWrite.All`** (simplest — access to every SharePoint site in your organization).

   Then click **Grant admin consent**.
5. If you chose `Sites.Selected`, give the app write access to the song site. A SharePoint or global
   administrator can do this in [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer):
   - `GET https://graph.microsoft.com/v1.0/sites/{your-tenant}.sharepoint.com:/sites/{site-name}` — copy the `id`.
   - `POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions` with this body:
     ```json
     {
       "roles": ["write"],
       "grantedToIdentities": [{ "application": { "id": "<Application (client) ID>", "displayName": "SongConverter" } }]
     }
     ```
     (Graph Explorer needs the `Sites.FullControl.All` permission consented for this one request.)

### 2. Find the document library ID (needed for library sync)

In Graph Explorer:

```
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
```

Copy the `id` of the document library that holds the songs (it starts with `b!`). This is `SYNC_DRIVE_ID`.

### 3. Create the Function App

<!-- AZURE-PLAN -->

### 4. Application settings

In the Function App: **Settings → Environment variables → App settings**.

| Setting | Required | Description |
|---|---|---|
| `TENANT_ID` | yes* | Directory (tenant) ID of the app registration |
| `CLIENT_ID` | yes* | Application (client) ID of the app registration |
| `CLIENT_SECRET` | yes* | Client secret value. Better: store it in Azure Key Vault and use `@Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/<name>/)` as the value |
| `OUTPUT_FOLDER` | no | Folder in the document library for the `.pro` files. Default `proPresenter Files`. Created automatically |
| `OUTPUT_DRIVE_ID` | no | Put the `.pro` files in a different document library (its drive ID). Default: same library as the PowerPoint |
| `SYNC_DRIVE_ID` | for sync | Document library to watch. **Setting this turns library sync on** |
| `SYNC_SOURCE_FOLDER` | no | Folder in that library to watch, e.g. `Songs`. Default: the library root |
| `SYNC_INCLUDE_SUBFOLDERS` | no | `true` to also convert files in subfolders. Default `false` |
| `SYNC_SCHEDULE` | no | How often to check, as a [CRON expression](https://learn.microsoft.com/azure/azure-functions/functions-bindings-timer#ncrontab-expressions). Default `0 */5 * * * *` (every 5 minutes) |
| `SYNC_MIN_AGE_SECONDS` | no | Wait until a file has not changed for this long before converting it (it may still be open in PowerPoint). Default `60` |
| `SYNC_MAX_FILES_PER_RUN` | no | Most files converted per check; the rest wait for the next check. Default `25` |
| `PRO_FONT_NAME` | no | PostScript font name for the lyrics. Default `TimesNewRomanPSMT` (others: `ArialMT`, `Arial-BoldMT`, `Helvetica`) |
| `PRO_FONT_SIZE` | no | Font size in points. Default `90` |
| `PRO_GROUP_NAME` | no | Name of the ProPresenter group holding the slides. Default `Group` |
| `INCLUDE_HIDDEN_SLIDES` | no | `true` to also convert slides hidden in PowerPoint. Default `false` |
| `MAX_FILE_SIZE_MB` | no | Skip larger files. Default `250` |

\* Alternatively, leave `CLIENT_SECRET` empty to sign in with the Function App's managed identity
(`MANAGED_IDENTITY_CLIENT_ID` selects a user-assigned identity). Microsoft Graph permissions for a
managed identity can only be granted with PowerShell or the Azure CLI, so the client secret is easier.

### 5. Deploy the code

Pushing to `main` builds, tests and deploys with GitHub Actions
(`.github/workflows/main_songconverter.yml`). Pushes to `dev` and pull requests are built and tested but not deployed.

<!-- AZURE-DEPLOY -->

## Option: Power Automate flow

Use this if you prefer the flow to the built-in library sync (it reacts within a minute or two instead of up to 5 minutes).

1. Trigger: **SharePoint — When a file is created or modified (properties only)** on the song library.
   Add a trigger condition so only PowerPoint files start the flow:
   `@or(endsWith(toLower(triggerOutputs()?['body/{FilenameWithExtension}']), '.pptx'), endsWith(toLower(triggerOutputs()?['body/{FilenameWithExtension}']), '.ppt'))`
2. Action: **HTTP** — Method `POST`, URI = the function URL including its key
   (Function App → Functions → `fileIsModified` → **Get function URL** → default key), body:
   ```json
   { "driveId": "<document library id>", "driveItemId": "<file's drive item id>" }
   ```

The function answers with JSON:

| Status | Meaning | Power Automate retries? |
|---|---|---|
| `200` `"status": "converted"` | `.pro` file written (`output.name`, `output.webUrl`) | — |
| `200` `"status": "skipped"` | Not a PowerPoint file, or it is in the output folder | — |
| `400` | Request body is missing `driveId` / `driveItemId` | no |
| `404` | File was deleted or moved | no |
| `422` | File could not be converted (not a presentation, password protected, no text) | no |
| `500` / `503` | Settings problem or temporary SharePoint/Graph error (see the logs) | yes |

Requests without a valid function key are rejected with `401`.

## ProPresenter computers

Sync the output folder (`proPresenter Files`) to each ProPresenter computer with OneDrive
(**Sync** or **Add shortcut to My files** in SharePoint), then import the `.pro` files into a ProPresenter
library, or point a ProPresenter library at that folder.

## How it handles tricky cases

- **Slide order** comes from the presentation itself, not file names (slide 10 does not come before slide 2).
- **Line breaks**: each paragraph and each Shift+Enter line break in PowerPoint becomes its own line; text split
  across differently formatted runs stays on one line. Blank lines at the top/bottom of a slide are removed and
  repeated blank lines are reduced to one.
- **Any language and punctuation**: curly quotes, dashes, accented and Cyrillic letters, emoji, `"`, `\`, `{`, `}`
  are all written as escaped Unicode, so nothing is garbled.
- **Hidden slides** are left out (unless `INCLUDE_HIDDEN_SLIDES=true`); slide numbers, dates and footers are ignored.
- **Simultaneous conversions** never share files, so they cannot mix up names or content.
- **Edits during conversion**: if the PowerPoint changes while it is being converted, it is converted again so the
  `.pro` always matches the latest version.
- **Failures are reported**: a file that cannot be converted is never uploaded as an empty `.pro`; the error is logged
  and returned to Power Automate.

## Working on the code

Requires Node.js 22.

```bash
npm install
npm test                                     # build and run the tests
npm run convert -- "Amazing Grace.pptx"      # convert a local file (writes "Amazing Grace.pro" next to it)
npm start                                    # run locally with Azure Functions Core Tools (needs local.settings.json)
```

Copy `local.settings.example.json` to `local.settings.json` for local runs (it is ignored by git — never commit secrets).

| Path | Purpose |
|---|---|
| `src/functions/fileIsModified.ts` | HTTP function for Power Automate |
| `src/functions/syncLibrary.ts` | Timer function for library sync |
| `src/lib/extract/` | Reads slide text from `.pptx` and `.ppt` files |
| `src/lib/pro/` | Builds the ProPresenter 7 file (RTF text + protobuf) |
| `src/lib/graph.ts` | Microsoft Graph (SharePoint) client with retries |
| `src/lib/job.ts`, `src/lib/sync.ts` | Convert one file / sync a folder |
| `proto/` | ProPresenter 7 file format definitions |
| `test/fixtures/propresenter-7.16-structure.json` | Layout of a file saved by ProPresenter 7.16 (no lyrics); the tests check generated files match it |

## Troubleshooting

Logs are in the Function App under **Monitor** (Application Insights) or **Log stream**.

| Log message | Fix |
|---|---|
| `Application setting X is required but not set` | Add the setting (step 4) |
| `Microsoft Graph returned 401/403` / `not allowed to access this file` | Check admin consent and the site permission (step 1), and that the client secret has not expired |
| `Microsoft Graph returned 404` for sync | `SYNC_DRIVE_ID` or `SYNC_SOURCE_FOLDER` is wrong (step 2) |
| `PowerPoint file does not contain any slide text` | The slides contain only pictures (e.g. lyrics pasted as images) |
| `PowerPoint file is password protected` | Remove the password in PowerPoint |
