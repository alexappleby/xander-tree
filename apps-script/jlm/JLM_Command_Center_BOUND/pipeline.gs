/**
 * ============================================================
 * JLM Pipeline — end-to-end audio intake
 * Google Sheet: JLM_Command_Center  (tab: "Songs")
 *
 * What it does (run from the "JLM Pipeline" menu, manual):
 *   1. Scans the Drive catalog root for audio masters not yet processed
 *      (skips DUPLICATE- , preview/video, and archive folders).
 *   2. For each new master, matches it to a canonical song row by
 *      normalized SONG-FOLDER NAME (never appends one row per file).
 *      Creates exactly one new row if no match; blocks if ambiguous.
 *   3. Writes Folder Link + audio-linking columns (AL-AQ) on that row.
 *   4. Transcribes the master (OpenAI gpt-4o-mini-transcribe).
 *   5. Drafts clean lyrics + metadata (OpenAI gpt-4.1-mini).
 *   6. Writes lyrics/metadata/note files into the song folder + links to the row.
 *   7. Sets AI Generation Status = "Ready for Art & Video" so the
 *      art + lyric-video worker (run separately) picks it up.
 *
 * Idempotent: a hidden "Pipeline_Processed" sheet tracks processed file IDs,
 * so re-running the menu continues the queue without redoing finished work.
 *
 * Self-contained: does NOT depend on Code.gs helpers. Requires the Script
 * Property OPENAI_API_KEY.
 * ============================================================
 */

const PIPE = {
  SONGS_SHEET: "Songs",
  LOG_SHEET: "System_Logs",
  CATALOG_ROOT_ID: "1lMcL9DniWH33Nyb6dpZItQESwKtciGdy",
  PROCESSED_SHEET: "Pipeline_Processed",

  // Folders under the catalog root to skip during the scan
  SKIP_FOLDERS: ["_Quarantine", "_Press Kit", "Z-Archive (Duplicates)",
                 "Duplicate tracks", "Artwork", "JLM_Release_Package.zip"],

  MAX_PER_RUN: 3,                 // Apps Script 6-min limit: process a few per run
  MAX_AUDIO_BYTES: 24 * 1024 * 1024,

  AUDIO_EXTENSIONS: [".mp3", ".m4a", ".wav", ".webm", ".flac", ".ogg", ".mpga"],

  // Column headers we read/write (mapped by name at runtime, never by letter)
  H: {
    TITLE: "Song Title",
    RELEASE: "EP / Release Name",
    FOLDER: "Folder Link",
    ART: "Art?",
    VIDEO: "Video?",
    ART_PROMPT: "Artwork Prompt",
    LYRICS_RAW: "Lyrics Raw",
    LYRICS_CLEAN: "Lyrics Clean",
    METADATA: "Metadata File",
    DISTROKID_NOTES: "DistroKid Notes",
    RIGHTS_REVIEW: "Rights Review",
    AI_STATUS: "AI Generation Status",
    AI_LAST_RUN: "AI Last Run",
    AI_ERROR: "AI Error",
    EXPLICIT: "Explicit Review",
    COVER_SONG: "Cover Song Review",
    SAMPLES: "Samples Review",
    SONGWRITER: "Songwriter Review",
    PUBLISHER: "Publisher Review",
    MASTER_OWNER: "Master Owner Review",
    ISRC_REVIEW: "ISRC Review",
    AUDIO_USED: "Audio File Used",
    SAF_ID: "Source Audio File ID",
    SAF_URL: "Source Audio File URL",
    SAF_NAME: "Source Audio File Name",
    SAF_FOLDER: "Source Audio Folder",
    INTAKE: "Drive Intake Status",
    INTAKE_LAST: "Drive Intake Last Run"
  },

  FILES: {
    RAW_LYRICS: "lyrics_raw.txt",
    CLEAN_LYRICS: "lyrics_clean_distrokid.txt",
    METADATA: "metadata.txt",
    DISTROKID_NOTES: "distrokid_upload_notes.txt",
    RIGHTS_REVIEW: "rights_review_needed.txt"
  }
};


// ============================================================
// MENU
// ============================================================
function onOpenPipeline_() {
  SpreadsheetApp.getUi()
    .createMenu("JLM Pipeline")
    .addItem("Run Pipeline (scan + process new audio)", "jlmRunPipeline")
    .addSeparator()
    .addItem("Reset Processed Queue (re-run everything)", "jlmResetProcessed")
    .addItem("Show Processed Log", "jlmShowProcessed")
    .addToUi();
}
// If this project has no onOpen yet, uncomment:
// function onOpen() { onOpenPipeline_(); }
// Otherwise add the JLM Pipeline menu items to the existing onOpen.


// ============================================================
// MAIN ENTRY
// ============================================================
function jlmRunPipeline() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(PIPE.SONGS_SHEET);
  if (!sheet) { ui.alert("No '" + PIPE.SONGS_SHEET + "' tab."); return; }

  const hmap = jlmHeaderMap_(sheet);
  const procSheet = jlmEnsureProcessedSheet_(ss);
  const processed = jlmLoadProcessedIds_(procSheet);

  // Scan catalog for song folders with a new master audio
  const items = [];
  jlmScanForNewAudio_(DriveApp.getFolderById(PIPE.CATALOG_ROOT_ID), processed, items, 0);

  if (!items.length) {
    ui.alert("JLM Pipeline", "No new audio masters found. Everything already processed.", ui.ButtonSet.OK);
    return;
  }

  let done = 0, blocked = 0;
  for (const it of items) {
    if (done >= PIPE.MAX_PER_RUN) break;
    try {
      const res = jlmProcessOne_(ss, sheet, hmap, procSheet, it);
      if (res === "blocked") blocked++;
      else done++;
    } catch (err) {
      jlmLog_(ss, "Pipeline error on '" + it.folderName + "': " + err);
      blocked++;
    }
  }

  const remaining = items.length - done - blocked;
  ui.alert(
    "JLM Pipeline",
    "Processed " + done + " song(s) this run.\n" +
    "Blocked: " + blocked + "\n" +
    "Remaining in queue: " + remaining + "\n\n" +
    (remaining > 0 ? "Run the menu again to continue." : ""),
    ui.ButtonSet.OK
  );
}


// ============================================================
// SCAN — find song folders containing an unprocessed master audio
// ============================================================
function jlmScanForNewAudio_(folder, processedIds, out, depth) {
  if (depth > 5) return;
  const fname = folder.getName();
  if (depth > 0 && PIPE.SKIP_FOLDERS.indexOf(fname) !== -1) return;

  // Does this folder directly contain a qualifying master audio?
  const best = jlmFindBestAudioInFolder_(folder);
  if (best && processedIds.indexOf(best.getId()) === -1) {
    out.push({
      folderId: folder.getId(),
      folderName: fname,
      folderUrl: folder.getUrl(),
      audioFile: best
    });
  }

  // Recurse into subfolders
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    jlmScanForNewAudio_(subs.next(), processedIds, out, depth + 1);
  }
}

function jlmFindBestAudioInFolder_(folder) {
  // depth-limited recursive search; returns the best-scoring master, or null.
  const candidates = [];
  jlmCollectAudio_(folder, candidates, 0);
  if (!candidates.length) return null;
  candidates.sort(function (a, b) { return b.score - a.score; });
  return candidates[0].file;
}

function jlmCollectAudio_(folder, candidates, depth) {
  if (depth > 2) return;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const lower = name.toLowerCase();

    if (!jlmIsSupportedAudio_(lower)) continue;
    // Skip DUPLICATE- prefixed files (catalog dedup markers)
    if (/^duplicate\s*-/i.test(name)) continue;
    // Skip preview/video/visual artifacts even if audio-ext named
    if (/(preview|video|9x16|16x9|spectrogram|waveform|lyricvideo|lyric\s*video)/i.test(lower)) continue;

    let score = 0;
    if (lower.indexOf("master") !== -1) score += 100;
    if (lower.indexOf("final") !== -1) score += 50;
    if (lower.indexOf("demo") !== -1) score -= 20;
    if (lower.endsWith(".wav") || lower.endsWith(".flac")) score += 40;
    if (lower.endsWith(".mp3") || lower.endsWith(".m4a")) score += 25;
    try { if (String(file.getMimeType()).indexOf("audio/") === 0) score += 30; } catch (e) {}
    score += Math.min(20, Math.floor(file.getSize() / (1024 * 1024)));
    candidates.push({ file: file, score: score });
  }
  const subs = folder.getFolders();
  while (subs.hasNext()) jlmCollectAudio_(subs.next(), candidates, depth + 1);
}

function jlmIsSupportedAudio_(lowerName) {
  for (let i = 0; i < PIPE.AUDIO_EXTENSIONS.length; i++) {
    if (lowerName.endsWith(PIPE.AUDIO_EXTENSIONS[i])) return true;
  }
  return false;
}


// ============================================================
// PROCESS ONE SONG
// ============================================================
function jlmProcessOne_(ss, sheet, hmap, procSheet, item) {
  const audio = item.audioFile;
  const folderName = item.folderName;

  // 1. Match / create canonical row by normalized folder name
  const rowRes = jlmFindOrCreateCanonicalRow_(sheet, hmap, folderName);
  if (rowRes.blocked) {
    jlmLog_(ss, "Blocked '" + folderName + "': " + rowRes.reason);
    jlmMarkProcessed_(procSheet, audio.getId(), audio.getName(), folderName, "", "Blocked", rowRes.reason);
    return "blocked";
  }
  const row = rowRes.row;
  const title = String(sheet.getRange(row, hmap[PIPE.H.TITLE]).getValue() || folderName).trim();

  // 2. Write Folder Link + audio-linking columns
  const songFolder = DriveApp.getFolderById(item.folderId);
  jlmSetRow_(sheet, row, hmap, [
    [PIPE.H.FOLDER, item.folderUrl],
    [PIPE.H.AUDIO_USED, audio.getName()],
    [PIPE.H.SAF_ID, audio.getId()],
    [PIPE.H.SAF_URL, audio.getUrl()],
    [PIPE.H.SAF_NAME, audio.getName()],
    [PIPE.H.SAF_FOLDER, folderName],
    [PIPE.H.INTAKE, "Audio Detected"],
    [PIPE.H.INTAKE_LAST, new Date()],
    [PIPE.H.AI_STATUS, "Transcribing"],
    [PIPE.H.AI_LAST_RUN, new Date()],
    [PIPE.H.AI_ERROR, ""]
  ]);

  // 3. Size guard
  if (audio.getSize() > PIPE.MAX_AUDIO_BYTES) {
    jlmSetRow_(sheet, row, hmap, [
      [PIPE.H.AI_STATUS, "Blocked - Audio Too Large"],
      [PIPE.H.AI_ERROR, "Audio over 24MB. Make a smaller MP3/M4A transcription copy."]
    ]);
    jlmMarkProcessed_(procSheet, audio.getId(), audio.getName(), folderName, row, "Blocked", "audio too large");
    return "blocked";
  }

  // 4. Transcribe
  let transcript = "";
  try {
    transcript = jlmTranscribe_(audio, title);
  } catch (err) {
    jlmSetRow_(sheet, row, hmap, [
      [PIPE.H.AI_STATUS, "Error - Transcription"],
      [PIPE.H.AI_ERROR, String(err && err.message ? err.message : err)]
    ]);
    jlmMarkProcessed_(procSheet, audio.getId(), audio.getName(), folderName, row, "Error", "transcription");
    return "blocked";
  }
  if (!transcript) {
    jlmSetRow_(sheet, row, hmap, [
      [PIPE.H.AI_STATUS, "Error - Transcription"],
      [PIPE.H.AI_ERROR, "Transcription returned empty."]
    ]);
    jlmMarkProcessed_(procSheet, audio.getId(), audio.getName(), folderName, row, "Error", "empty transcript");
    return "blocked";
  }

  // 5. Draft lyrics + metadata
  let draft;
  try {
    draft = jlmDraft_({
      songTitle: title,
      releaseName: String(sheet.getRange(row, hmap[PIPE.H.RELEASE]).getValue() || "").trim(),
      audioFileName: audio.getName(),
      transcript: transcript
    });
  } catch (err) {
    jlmSetRow_(sheet, row, hmap, [
      [PIPE.H.AI_STATUS, "Error - Drafting"],
      [PIPE.H.AI_ERROR, String(err && err.message ? err.message : err)]
    ]);
    jlmMarkProcessed_(procSheet, audio.getId(), audio.getName(), folderName, row, "Error", "drafting");
    return "blocked";
  }

  // 6. Write generated files into the song folder
  const rawLyrics = jlmReplaceTextFile_(songFolder, PIPE.FILES.RAW_LYRICS, transcript);
  const cleanLyrics = jlmReplaceTextFile_(songFolder, PIPE.FILES.CLEAN_LYRICS, draft.cleanLyrics || ("NEEDS REVIEW\n\n" + transcript));
  const metadata = jlmReplaceTextFile_(songFolder, PIPE.FILES.METADATA, draft.metadata || jlmFallbackMetadata_(title, audio.getName()));
  const notes = jlmReplaceTextFile_(songFolder, PIPE.FILES.DISTROKID_NOTES, draft.distrokidNotes || jlmDefaultNotes_(title));
  const rights = jlmReplaceTextFile_(songFolder, PIPE.FILES.RIGHTS_REVIEW, draft.rightsReview || jlmDefaultRights_(title));

  // 7. Write links + reviews + status to the row
  jlmSetRow_(sheet, row, hmap, [
    [PIPE.H.LYRICS_RAW, rawLyrics.getUrl()],
    [PIPE.H.LYRICS_CLEAN, cleanLyrics.getUrl()],
    [PIPE.H.METADATA, metadata.getUrl()],
    [PIPE.H.DISTROKID_NOTES, notes.getUrl()],
    [PIPE.H.RIGHTS_REVIEW, rights.getUrl()],
    [PIPE.H.EXPLICIT, "Review Needed"],
    [PIPE.H.COVER_SONG, "Review Needed"],
    [PIPE.H.SAMPLES, "Review Needed"],
    [PIPE.H.SONGWRITER, "Review Needed"],
    [PIPE.H.PUBLISHER, "Review Needed"],
    [PIPE.H.MASTER_OWNER, "Review Needed"],
    [PIPE.H.ISRC_REVIEW, "Review Needed"],
    [PIPE.H.AI_STATUS, "Ready for Art & Video"],
    [PIPE.H.AI_LAST_RUN, new Date()],
    [PIPE.H.AI_ERROR, ""]
  ]);

  // 8. Mark processed
  jlmMarkProcessed_(procSheet, audio.getId(), audio.getName(), folderName, row, "Complete", "");
  jlmLog_(ss, "Pipeline processed '" + title + "' (row " + row + "). Ready for art & video.");
  return "done";
}


// ============================================================
// CANONICAL ROW MATCH / CREATE
// ============================================================
function jlmFindOrCreateCanonicalRow_(sheet, hmap, folderName) {
  const norm = jlmNormalizeTitle_(folderName);
  const titleCol = hmap[PIPE.H.TITLE];
  const values = sheet.getDataRange().getValues();

  const matches = [];
  for (let i = 1; i < values.length; i++) {
    const t = String(values[i][titleCol - 1] || "").trim();
    if (!t) continue;
    if (!jlmIsCanonicalTitle_(t)) continue;          // ignore clutter rows
    if (jlmNormalizeTitle_(t) === norm) matches.push(i + 1);
  }

  if (matches.length === 1) return { row: matches[0] };
  if (matches.length > 1) {
    return { blocked: true, row: null, reason: "Ambiguous: " + matches.length + " canonical rows match '" + folderName + "' (rows " + matches.join(",") + ")" };
  }
  // No match -> create exactly one canonical row
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, titleCol).setValue(folderName);
  return { row: newRow };
}

function jlmIsCanonicalTitle_(t) {
  t = String(t).trim();
  if (!t) return false;
  if (/^\d{8}[\s_-]?\d{6}$/.test(t)) return false;             // timestamp
  if (/^duplicate/i.test(t)) return false;                    // duplicate-prefixed
  if (/^\d{1,3}\s*[-._)]?\s+/.test(t)) return false;           // numbered re-list
  if (/(?:lyric\s*video|lyricvideo|video|preview)\s*(?:9x16|16x9)?\s*$/i.test(t)) return false; // asset suffix
  if (/(?:audio|vet\s*promo)\s*$/i.test(t)) return false;
  return true;
}

function jlmNormalizeTitle_(t) {
  let s = String(t);
  s = s.replace(/^\d{1,3}\s*[-._)]?\s+/, " ");
  s = s.replace(/^duplicate\s*\d*\s*/i, " ");
  s = s.toLowerCase();
  s = s.replace(/[^a-z0-9\s]/g, " ");
  return s.replace(/\s+/g, " ").trim();
}


// ============================================================
// OPENAI — transcription + drafting
// ============================================================
function jlmTranscribe_(audioFile, songTitle) {
  const apiKey = jlmGetProp_("OPENAI_API_KEY", "");
  if (!apiKey) throw new Error("OPENAI_API_KEY Script Property is missing.");
  const model = jlmGetProp_("OPENAI_TRANSCRIBE_MODEL", "gpt-4o-mini-transcribe");
  const blob = audioFile.getBlob().setName(audioFile.getName());

  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions", {
    method: "post",
    headers: { Authorization: "Bearer " + apiKey },
    payload: {
      model: model,
      file: blob,
      response_format: "json",
      prompt: "This is a country/Americana song by Jim Lamb Music. Transcribe the sung lyrics as accurately as possible for the song titled: " + songTitle
    },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error("OpenAI transcription HTTP " + code + ": " + text);
  return String(JSON.parse(text).text || "").trim();
}

function jlmDraft_(info) {
  const apiKey = jlmGetProp_("OPENAI_API_KEY", "");
  const model = jlmGetProp_("OPENAI_TEXT_MODEL", "gpt-4.1-mini");

  const prompt =
    "You are helping prepare a music release package for DistroKid.\n\n" +
    "Artist: Jim Lamb Music\n" +
    "Release: " + info.releaseName + "\n" +
    "Song Title: " + info.songTitle + "\n" +
    "Audio File: " + info.audioFileName + "\n\n" +
    "Transcript:\n" + info.transcript + "\n\n" +
    "Create draft outputs for this song.\n\n" +
    "Rules:\n" +
    "- Do not claim rights, ownership, publishing, songwriter splits, ISRC, UPC, sample clearance, cover-song status, or explicit status as final.\n" +
    "- Mark all rights/licensing fields as Review Needed.\n" +
    "- Clean lyrics for plain lyric submission: remove timestamps, avoid weird spacing, avoid all-caps except proper words, use stanza breaks.\n" +
    "- If lyrics are uncertain, mark [unclear] sparingly.\n\n" +
    "Return EXACTLY in this format with these headings:\n\n" +
    "=== CLEAN_LYRICS ===\n[cleaned lyrics]\n\n" +
    "=== METADATA ===\n" +
    "Song Title:\nArtist: Jim Lamb Music\nRelease:\nAudio File:\nLanguage:\nGenre: Country\nSubgenre: Americana\nMood:\nThemes:\nShort Description:\n" +
    "Explicit: Review Needed\nSongwriter(s): Review Needed\nPublisher: Review Needed\nMaster Owner: Review Needed\n" +
    "Composition Owner: Review Needed\nFeatured Artist: Review Needed\nISRC: Review Needed\nUPC: Album-level / DistroKid generated\n" +
    "Cover Song?: Review Needed\nSamples?: Review Needed\nReady Status: Needs Review\n\n" +
    "=== DISTROKID_NOTES ===\n[upload notes]\n\n" +
    "=== RIGHTS_REVIEW ===\n[rights and licensing checklist]\n";

  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
    method: "post",
    headers: { Authorization: "Bearer " + apiKey, "Content-Type": "application/json" },
    payload: JSON.stringify({ model: model, input: prompt }),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error("OpenAI draft HTTP " + code + ": " + text);
  const out = jlmExtractResponseText_(JSON.parse(text));
  return {
    cleanLyrics: jlmExtractSection_(out, "CLEAN_LYRICS"),
    metadata: jlmExtractSection_(out, "METADATA"),
    distrokidNotes: jlmExtractSection_(out, "DISTROKID_NOTES"),
    rightsReview: jlmExtractSection_(out, "RIGHTS_REVIEW"),
    fullText: out
  };
}

function jlmExtractResponseText_(data) {
  if (data && data.output_text) return data.output_text;
  if (data && Array.isArray(data.output)) {
    const parts = [];
    for (const item of data.output) {
      if (item && item.content && Array.isArray(item.content)) {
        for (const c of item.content) { if (c && c.type === "output_text" && c.text) parts.push(c.text); }
      }
    }
    if (parts.length) return parts.join("\n");
  }
  return String(data && data.text ? data.text : JSON.stringify(data));
}

function jlmExtractSection_(text, heading) {
  const re = new RegExp("===\\s*" + heading + "\\s*===([\\s\\S]*?)(?:===\\s*[A-Z_]+\\s*===|$)", "i");
  const m = re.exec(text);
  return m ? m[1].trim() : "";
}


// ============================================================
// DRIVE / SHEET HELPERS
// ============================================================
function jlmReplaceTextFile_(folder, filename, content) {
  const existing = folder.getFilesByName(filename);
  if (existing.hasNext()) existing.next().setTrashed(true);
  return folder.createFile(filename, content, "text/plain");
}

function jlmSetRow_(sheet, row, hmap, kv) {
  // kv: array of [headerName, value]
  const updates = [];
  for (const [name, val] of kv) {
    const col = hmap[name];
    if (col) updates.push({ row: row, col: col, val: val });
  }
  // Batch by writing each (small N per song)
  for (const u of updates) sheet.getRange(u.row, u.col).setValue(u.val);
}

function jlmHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const name = String(headers[i] || "").trim();
    if (name && !map[name]) map[name] = i + 1;
  }
  return map;
}

function jlmGetProp_(name, def) {
  const v = PropertiesService.getScriptProperties().getProperty(name);
  return v || def;
}


// ============================================================
// PROCESSED-TRACKING SHEET (hidden)
// ============================================================
function jlmEnsureProcessedSheet_(ss) {
  let s = ss.getSheetByName(PIPE.PROCESSED_SHEET);
  if (!s) {
    s = ss.insertSheet(PIPE.PROCESSED_SHEET);
    s.getRange(1, 1, 1, 7).setValues([[
      "File ID", "File Name", "Song Folder", "Song Row", "Status", "Note", "Processed At"
    ]]).setFontWeight("bold");
    s.hideSheet();
  }
  return s;
}

function jlmLoadProcessedIds_(procSheet) {
  const v = procSheet.getDataRange().getValues();
  const ids = [];
  for (let i = 1; i < v.length; i++) {
    const id = String(v[i][0] || "").trim();
    if (id) ids.push(id);
  }
  return ids;
}

function jlmMarkProcessed_(procSheet, fileId, fileName, folder, row, status, note) {
  const r = procSheet.getLastRow() + 1;
  procSheet.getRange(r, 1, 1, 7).setValues([[
    fileId, fileName, folder, row || "", status, note, new Date()
  ]]);
}

function jlmResetProcessed() {
  const ss = SpreadsheetApp.getActive();
  const s = ss.getSheetByName(PIPE.PROCESSED_SHEET);
  if (!s) { SpreadsheetApp.getUi().alert("No processed log yet."); return; }
  const last = s.getLastRow();
  if (last > 1) s.deleteRows(2, last - 1);
  SpreadsheetApp.getUi().alert("Processed queue cleared. Re-run the pipeline to reprocess all audio.");
}

function jlmShowProcessed() {
  const ss = SpreadsheetApp.getActive();
  const s = ss.getSheetByName(PIPE.PROCESSED_SHEET);
  if (s) { s.showSheet(); ss.setActiveSheet(s); }
}


// ============================================================
// FALLBACK DRAFTS
// ============================================================
function jlmFallbackMetadata_(title, audioFile) {
  return "Song Title: " + title + "\nArtist: Jim Lamb Music\nAudio File: " + audioFile +
    "\nLanguage: English\nGenre: Country\nSubgenre: Americana\nReady Status: Needs Review\n";
}
function jlmDefaultNotes_(title) {
  return "Upload notes for " + title + ":\n- Confirm final master WAV used\n- Confirm ISRC assigned\n- Confirm explicit/cover/sample status before release\n";
}
function jlmDefaultRights_(title) {
  return "Rights review for " + title + ":\n- Songwriter(s): Review Needed\n- Publisher: Review Needed\n- Master Owner: Review Needed\n- Cover song?: Review Needed\n- Samples?: Review Needed\n";
}


// ============================================================
// LOG
// ============================================================
function jlmLog_(ss, message) {
  let s = ss.getSheetByName(PIPE.LOG_SHEET);
  if (!s) return;
  try {
    const r = s.getLastRow() + 1;
    s.getRange(r, 1, 1, Math.min(3, s.getLastColumn() || 3)).setValues([[new Date(), "Pipeline", message]]);
  } catch (e) {}
}
