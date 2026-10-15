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

All `.pro` files go into **one folder, `proPresenter Files/`, with no subfolders** — even when the PowerPoints
come from different folders (or, with `SYNC_INCLUDE_SUBFOLDERS`, from subfolders). Two PowerPoints with the same
file name therefore share one `.pro` (the most recently converted wins).

The app can be triggered two ways — use either or both:

| | **Library sync** (recommended) | **Power Automate flow** |
|---|---|---|
| How | A timer in the Function App checks a SharePoint folder every 5 minutes and converts files added or changed since the previous check | A flow calls the function each time a file changes |
| Needs | Nothing else | Power Automate Premium license for the flow owner (the **HTTP** action is premium) |
| Turn on | Set `SYNC_SITE_URL` (and `SYNC_LIBRARY`); the `syncLibrary` function does nothing until then | Create the flow (below) |
| Turn off | Disable the `syncLibrary` function in the portal | Turn off the flow |

### How library sync decides what to convert

- **First run:** it records the PowerPoint files already in the folder and converts nothing, so turning sync on does
  not convert your whole existing library. (To convert existing songs once, set `SYNC_CONVERT_EXISTING=true` before
  the first run; files that already have an up-to-date, non-empty `.pro` are skipped.)
- **Replacing files made by the old version:** `.pro` files from the old converter may be misformatted. To rebuild them
  all, rename the old `proPresenter Files` folder (e.g. to `proPresenter Files (old)`), set `SYNC_CONVERT_EXISTING=true`,
  and turn sync on; every song is converted again, 25 per run.
- **Every run after that:** it converts files that were **added** since the previous run, and files whose **content
  changed** (an edited song gets a fresh `.pro`). Renaming a file or editing its SharePoint columns does not count as a change.
- The `proPresenter Files` folder is never opened while polling, even when it is inside the watched folder or
  library (it holds only `.pro` files). PowerPoints put in it by mistake are ignored.
- Files are tracked by their SharePoint id, not by date, so a file **moved or copied into** the folder from elsewhere
  counts as added even though it keeps its old dates.
- A file is left alone until it has not changed for a minute (it may still be open in PowerPoint), and at most 25 files
  are converted per run; anything left over is picked up by the next run.
- If a conversion fails because SharePoint is temporarily unavailable, the file is retried on following runs (for up
  to a day). A file that cannot be converted at all (no text, password protected) is logged and skipped until it changes.
- What the sync has seen is kept in the Function App's own storage account (container `songconverter`, blob
  `sync-state.json`). Deleting that blob makes the next run start over with a new baseline. Changing
  the watched library, `SYNC_SOURCE_FOLDER` or `SYNC_INCLUDE_SUBFOLDERS` also starts a new baseline.

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

### 2. Note the site address and library name (needed for library sync)

- **Site address:** open the SharePoint site in a browser and copy the address up to the site name, e.g.
  `https://contoso.sharepoint.com/sites/Worship`. This is `SYNC_SITE_URL`.
- **Library:** the name of the document library as SharePoint shows it (the default library is called `Documents`;
  `Shared Documents`, the name in its web address, also works). This is `SYNC_LIBRARY`. Instead of the name you can use
  the library's id from **Library settings** (the `List=` part of that page's address, with or without the `%7B…%7D`).

The app looks up the library from these. (Advanced: you can instead set `SYNC_DRIVE_ID` to the library's Microsoft
Graph drive id, which starts with `b!`, from `GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives` in
[Graph Explorer](https://developer.microsoft.com/graph/graph-explorer).)

### 3. Create the Function App

In the [Azure portal](https://portal.azure.com): **Create a resource → Function App**, then choose these options.

| Option | Choose | Why |
|---|---|---|
| Hosting plan | **Flex Consumption** | Microsoft's current serverless plan: you pay only while a conversion runs, and the monthly free allowance covers this app. (The older *Consumption* plan is now "legacy"; on Linux it is being retired.) |
| Runtime stack / version | **Node.js 24** | Supported until April 2028 (Node.js 22 support ends April 2027) |
| Operating system | Linux | The only option on Flex Consumption |
| Region | The one closest to you | |
| Instance size | **2048 MB** (default) | Enough for large presentations |
| Zone redundancy | Disabled | Not needed; costs more |
| Storage | Create new (Standard LRS) | Required by Azure Functions; costs cents per month |
| Always ready instances | **0** (default) | Anything above 0 is billed around the clock |
| Maximum instance count | **10** (Scale and concurrency settings) | Caps the cost if something goes wrong; plenty for this app |
| Public access | On | Power Automate and Microsoft Graph must reach the app |
| Application Insights | **Yes** | This is where the logs are |
| Continuous deployment (GitHub Actions) | Disable | This repository already has a workflow; step 5 connects it |

Notes:

- "Basic" and "Standard" are App Service plans with a fixed monthly price (and the Free tier cannot run Azure Functions),
  so Flex Consumption is the cheapest option. Its free monthly allowance applies to pay-as-you-go subscriptions;
  subscriptions running on credits or a free trial may not get it.
- Timer schedules on Flex Consumption run in UTC (the 5-minute check is not affected).
- Already have the old `SongConverter` app on the Windows *Consumption* plan? It still works: change it to
  Node.js 24 (**Settings → Configuration → General settings**), replace its app settings with the ones in step 4,
  and keep deploying with its publish profile. Do not create new apps on Linux *Consumption*.

### 4. Application settings

In the Function App: **Settings → Environment variables → App settings**.

| Setting | Required | Description |
|---|---|---|
| `TENANT_ID` | yes* | Directory (tenant) ID of the app registration |
| `CLIENT_ID` | yes* | Application (client) ID of the app registration |
| `CLIENT_SECRET` | yes* | Client secret value. Better: store it in Azure Key Vault and use `@Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/<name>/)` as the value |
| `OUTPUT_FOLDER` | no | The one folder, at the top of the document library, that holds every `.pro` file. Default `proPresenter Files`. Created automatically. Must be a single folder name — paths like `proPresenter Files/Sunday` are rejected |
| `OUTPUT_DRIVE_ID` | no | Graph drive id (`b!…`) of a different library for that folder. Default: the synced library; if sync is not set up, the PowerPoint's own library |
| `SYNC_SITE_URL` | for sync | SharePoint site to watch, e.g. `https://contoso.sharepoint.com/sites/Worship`. **Setting this turns library sync on** |
| `SYNC_LIBRARY` | no | Document library in that site: its name (default `Documents`) or its id from Library settings |
| `SYNC_DRIVE_ID` | no | Advanced alternative to the two settings above: the library's Graph drive id (`b!…`) |
| `SYNC_SOURCE_FOLDER` | no | Folder in that library to watch, e.g. `Songs`. Default: the library root |
| `SYNC_INCLUDE_SUBFOLDERS` | no | `true` to also convert files in subfolders. Default `false` |
| `SYNC_SCHEDULE` | no | How often to check, as a [CRON expression](https://learn.microsoft.com/azure/azure-functions/functions-bindings-timer#ncrontab-expressions). Default `0 */5 * * * *` (every 5 minutes) |
| `SYNC_MIN_AGE_SECONDS` | no | Wait until a file has not changed for this long before converting it (it may still be open in PowerPoint). Default `60` |
| `SYNC_MAX_FILES_PER_RUN` | no | Most files converted per check; the rest wait for the next check. Default `25` |
| `SYNC_CONVERT_EXISTING` | no | `true` to convert existing files that have no up-to-date `.pro` on the first run, instead of only recording them. Default `false` |
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

For a **Flex Consumption** app, GitHub signs in to Azure with OpenID Connect (Flex does not accept publish profiles).
Run this once in [Azure Cloud Shell](https://shell.azure.com) (Bash), with your names filled in:

```bash
RG=<resource group>; APP=<function app name>; REPO=TulsaFSBC/songConverter
az identity create -g $RG -n songconverter-deploy
az identity federated-credential create -g $RG --identity-name songconverter-deploy -n github-main \
  --issuer https://token.actions.githubusercontent.com \
  --subject repo:$REPO:ref:refs/heads/main --audiences api://AzureADTokenExchange
az role assignment create --role "Website Contributor" \
  --assignee-object-id $(az identity show -g $RG -n songconverter-deploy --query principalId -o tsv) \
  --assignee-principal-type ServicePrincipal \
  --scope $(az functionapp show -g $RG -n $APP --query id -o tsv)
az identity show -g $RG -n songconverter-deploy --query clientId -o tsv   # -> AZURE_CLIENT_ID
az account show --query "{tenant:tenantId, subscription:id}" -o table      # -> AZURE_TENANT_ID, AZURE_SUBSCRIPTION_ID
```

Then in GitHub (**Settings → Secrets and variables → Actions**):

- Secrets: `AZURE_CLIENT_ID`, `AZURE_TENANT_ID`, `AZURE_SUBSCRIPTION_ID` (the values printed above).
- Variables: `AZURE_FUNCTIONAPP_NAME` = your Function App name.

For a Windows *Consumption* app you can instead keep using a publish profile: download it from the Function App
overview (**Get publish profile**) and save it as the secret `AZURE_PUBLISH_PROFILE` (the existing
`AZUREAPPSERVICE_PUBLISHPROFILE_…` secret also still works).

After the first deployment the app has two functions: `fileIsModified` (HTTP, for Power Automate) and
`syncLibrary` (timer, every 5 minutes). `syncLibrary` only logs "Library sync is off" until `SYNC_SITE_URL` is set;
if you use only Power Automate, disable it (Functions → `syncLibrary` → **Disable**).

Settings from the old version that are no longer used and can be deleted: `PRO_PRESENTER_VERSION`, `KEY_VAULT_URL`
(and the `accessToken` secret in that Key Vault).

## Option: Power Automate flow

Use this if you prefer a flow to the built-in library sync (it reacts within a few minutes of a change).
The flow's owner needs a Power Automate Premium license, because the **HTTP** action is a premium connector.

1. Trigger: **SharePoint — When a file is created or modified (properties only)** on the song library.
   In the trigger's **Settings → Trigger conditions**, add this so only PowerPoint files start the flow:
   `@or(endsWith(triggerOutputs()?['body/{FilenameWithExtension}'], '.pptx'), endsWith(triggerOutputs()?['body/{FilenameWithExtension}'], '.ppt'))`
2. Action: **HTTP** — Method `POST`; URI = the function URL with its key (Function App → Functions →
   `fileIsModified` → **Get function URL** → *default (function key)*); Body:
   ```json
   {
     "siteUrl": "https://<tenant>.sharepoint.com/sites/<site>",
     "listId": "<library id>",
     "listItemId": @{triggerOutputs()?['body/ID']}
   }
   ```
   The library id is in the address bar of the library's **Library settings** page (`List=%7B…%7D`; paste it as is).
   An existing flow that sends `{ "driveId": "…", "driveItemId": "…" }` keeps working — just add the function key to its URL.

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
- **Hidden slides** are left out (unless `INCLUDE_HIDDEN_SLIDES=true`), and so are text boxes hidden with PowerPoint's
  Selection Pane; slide numbers, dates and footers are ignored.
- **Simultaneous conversions** never share files, so they cannot mix up names or content.
- **Edits during conversion**: if the PowerPoint changes while it is being converted, it is converted again so the
  `.pro` always matches the latest version.
- **Failures are reported**: a file that cannot be converted is never uploaded as an empty `.pro`; the error is logged
  and returned to Power Automate.

## Working on the code

Requires Node.js 22 or newer (Azure runs Node.js 24).

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
| `src/lib/state.ts` | Where library sync remembers which files it has handled |
| `proto/` | ProPresenter 7 file format definitions |
| `test/fixtures/propresenter-7.16-structure.json` | Layout of a file saved by ProPresenter 7.16 (no lyrics); the tests check generated files match it |

## Troubleshooting

Logs are in the Function App under **Monitor** (Application Insights) or **Log stream**.

| Log message | Fix |
|---|---|
| `Application setting X is required but not set` | Add the setting (step 4) |
| `Microsoft Graph returned 401/403` / `not allowed to access this file` | Check admin consent and the site permission (step 1), and that the client secret has not expired |
| `SYNC_SOURCE_FOLDER "…" was not found` | Use the folder path inside the library without the library name: `Songs`, not `Shared Documents/Songs` |
| `SharePoint site "…" was not found` / `No document library named "…"` | Check `SYNC_SITE_URL` and `SYNC_LIBRARY` (step 2); the message lists the libraries it found |
| `SYNC_DRIVE_ID "…" is a SharePoint library id, not a Microsoft Graph drive id` | Also set `SYNC_SITE_URL`, or move the id to `SYNC_LIBRARY` |
| `Library sync needs the AzureWebJobsStorage setting` | The Function App's storage setting was removed; it is created with the app and must stay |
| `Sync started: recorded the N PowerPoint files already in the folder` | Normal on the first run: only files added or changed after this are converted |
| `PowerPoint file does not contain any slide text` | The slides contain only pictures (e.g. lyrics pasted as images) |
| `PowerPoint file is password protected` | Remove the password in PowerPoint |
| `PowerPoint file is damaged and cannot be read` | Open the file in PowerPoint and save it again |
