# JLM Apps Script Backups

Source backups of Google Apps Script projects used across the Jim Lamb Music
catalog and related business automation. These are plain-text snapshots of the
live Apps Script projects in Google Drive, kept here for version history and
disaster recovery.

## Projects

| Folder | What it is | Live location |
| --- | --- | --- |
| `JLM_Pipeline/` | End-to-end audio intake pipeline. Scans the Drive catalog for new master audio, matches each to a canonical song row in the JLM_Command_Center sheet (by normalized song-folder name), transcribes via OpenAI, drafts lyrics + metadata, writes files into the song folder, and sets status to "Ready for Art & Video". | Plain-text `.gs` file in the Drive catalog root |
| `Dedupe_run/` | Drive catalog dedup runner. Walks the catalog root, dedupes binary assets, flags duplicate/empty folders with `DUPLICATE - ` / `EMPTY - ` prefixes. | Standalone Apps Script project |
| `DriveDedupeLibrary/` | Reusable dedup library — md5-based keeper selection (newest modified wins), mime-type safety filters (audio/video/image only). | Standalone Apps Script project |
| `Copy_of_DriveDedupeLibrary/` | Trimmed variant of the dedup library. | Standalone Apps Script project |
| `Inventory_capture_and_tracker/` | Mobile-friendly inventory scanner web app (HTML + Apps Script backend). | Standalone Apps Script project |
| `FundingQualifierForm/` | Apps Script form-builder for a funding qualifier intake form. | Standalone Apps Script project |
| `JLM_Command_Center_BOUND/` | **The live, actually-running code.** Bound to the JLM_Command_Center spreadsheet. Contains the "JLM AI Tools" menu, the AI metadata/lyrics engine (Gemini + OpenAI, queue-based processing over the Songs sheet), the WAV intake automation, and a duplicate copy of the pipeline scanner (`pipeline.gs`). | Bound Apps Script project on JLM_Command_Center sheet |

## Notes

- **Secrets**: API keys (e.g. `OPENAI_API_KEY`) are stored in Apps Script
  Script Properties, not in this source. No credentials are committed here.
- **IDs**: Folder IDs, sheet IDs, and the catalog root ID are present in the
  source (needed for the scripts to run). They are not secrets, but they do
  reference the live Drive structure.
- **Bound vs standalone**: `JLM_Pipeline/` is a plain-text copy kept in Drive as
  a separate file (not the live code). `JLM_Command_Center_BOUND/` is the actual
  live project bound to the spreadsheet — it's the one that runs when you use
  the "JLM AI Tools" and "JLM Pipeline" menus in the sheet.
- **Catalog cleanup (2026-07-19)**: the pipeline scanner skips any folder
  prefixed `EMPTY - ` or `DUPLICATE - ` (the flags applied during the Drive
  catalog cleanup). See the `SKIP_FOLDER_PATTERN` constant in
  `pipeline.gs` (both `JLM_Pipeline/` and `JLM_Command_Center_BOUND/` copies
  are in sync as of 2026-07-21).
- **Menu bug fixed (2026-07-21)**: the original `onOpen()` in
  `JLM_Command_Center_BOUND/Code.gs` chained `.createMenu("JLM AI Tools").createMenu("Pipeline")`,
  which is invalid — Apps Script's menu builder doesn't support calling
  `createMenu()` on the result of another `createMenu()`. This threw
  `TypeError: ...createMenu is not a function` on every sheet load, meaning
  the menus never actually rendered. Fixed using `.addSubMenu()` to properly
  nest "Pipeline" inside "JLM AI Tools", and wired in a call to
  `onOpenPipeline_()` so the separate "JLM Pipeline" top-level menu
  (Run Pipeline / Reset Queue / Show Log) also renders on load.
- **`Javascript.gs` removed (2026-07-21)**: it contained two unreferenced
  scratch functions (`getCloudHash`, `generateMetadataManifest`) pointing at
  a Cloud Function URL that was never deployed. Deleted from the live project
  as dead code; no longer present in this backup.

## Restore

To restore a project: create a new Apps Script project (standalone, or bound to
a sheet via Extensions → Apps Script), then paste in the `.gs` / `.html` files
from the corresponding folder. Set any required Script Properties (e.g.
`OPENAI_API_KEY`) manually after restoring.
