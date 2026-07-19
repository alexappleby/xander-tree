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

## Notes

- **Secrets**: API keys (e.g. `OPENAI_API_KEY`) are stored in Apps Script
  Script Properties, not in this source. No credentials are committed here.
- **IDs**: Folder IDs, sheet IDs, and the catalog root ID are present in the
  source (needed for the scripts to run). They are not secrets, but they do
  reference the live Drive structure.
- **Bound vs standalone**: `JLM_Pipeline/` is the plain-text copy kept in Drive.
  The *bound* script project on the JLM_Command_Center spreadsheet (the one that
  runs the "JLM Pipeline" menu live) is not included here — if it diverges from
  this file, back it up separately via Extensions → Apps Script in the sheet.
- **Catalog cleanup (2026-07-19)**: the pipeline scanner was updated to skip any
  folder prefixed `EMPTY - ` or `DUPLICATE - ` (the flags applied during the
  Drive catalog cleanup). See the `SKIP_FOLDER_PATTERN` constant in
  `JLM_Pipeline/JLM_Pipeline.gs`.

## Restore

To restore a project: create a new Apps Script project (standalone, or bound to
a sheet via Extensions → Apps Script), then paste in the `.gs` / `.html` files
from the corresponding folder. Set any required Script Properties (e.g.
`OPENAI_API_KEY`) manually after restoring.
