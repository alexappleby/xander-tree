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
  the "JLM AI Tools" menu in the sheet. As of 2026-07-21 these have **drifted**:
  the bound `pipeline.gs` is missing the folder-skip fix below. See "Known drift"
  before assuming either copy reflects production behavior.
- **Catalog cleanup (2026-07-19)**: the pipeline scanner was updated to skip any
  folder prefixed `EMPTY - ` or `DUPLICATE - ` (the flags applied during the
  Drive catalog cleanup). See the `SKIP_FOLDER_PATTERN` constant in
  `JLM_Pipeline/JLM_Pipeline.gs`. This fix has **not** been applied to the live
  bound script (`JLM_Command_Center_BOUND/pipeline.gs`) yet — see "Known drift".
- **`Javascript.gs` is dead code.** Both functions in
  `JLM_Command_Center_BOUND/Javascript.gs` (`getCloudHash`, `generateMetadataManifest`)
  are unreferenced anywhere else in the project — a placeholder Cloud Function URL
  was never deployed. Kept for reference only; safe to delete from the live
  project if unused.

## Known drift (as of 2026-07-21)

The live bound `pipeline.gs` does **not** have the `SKIP_FOLDER_PATTERN` fix that
`JLM_Pipeline/JLM_Pipeline.gs` has. To bring it in sync, open the JLM_Command_Center
sheet → Extensions → Apps Script → `pipeline.gs`, and apply:

1. Add to the `PIPE` config object:
   ```js
   SKIP_FOLDER_PATTERN: /^\s*(EMPTY|DUPLICATE)\s*-/i,
   ```
2. Add a helper function:
   ```js
   function jlmIsFlaggedFolder_(name) {
     return PIPE.SKIP_FOLDER_PATTERN.test(String(name || ""));
   }
   ```
3. In the folder-scan function, skip flagged folders early:
   ```js
   if (PIPE.SKIP_FOLDER_PATTERN.test(fname)) return; // EMPTY - / DUPLICATE - flagged folder, any depth
   ```
4. In the recursive subfolder walk, skip flagged folders before recursing:
   ```js
   while (subs.hasNext()) {
     const sub = subs.next();
     if (jlmIsFlaggedFolder_(sub.getName())) continue; // skip EMPTY - / DUPLICATE - flagged folders
     jlmCollectAudio_(sub, candidates, depth + 1);
   }
   ```

Full reference diff: compare `JLM_Command_Center_BOUND/pipeline.gs` (live, pre-fix)
against `JLM_Pipeline/JLM_Pipeline.gs` (updated, post-fix) in this repo.

## Restore

To restore a project: create a new Apps Script project (standalone, or bound to
a sheet via Extensions → Apps Script), then paste in the `.gs` / `.html` files
from the corresponding folder. Set any required Script Properties (e.g.
`OPENAI_API_KEY`) manually after restoring.
