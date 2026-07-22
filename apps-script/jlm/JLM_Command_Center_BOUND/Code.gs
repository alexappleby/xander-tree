
/**
 * JLM Catalog AI Engine
 * Google Sheet: JLM_Command_Center
 *
 * What this installs:
 * - JLM AI Tools menu
 * - Metadata + lyrics queue
 * - Per-song generation
 * - Release-level queue generation
 * - Whole-catalog queue generation
 *
 * Required Script Property:
 * OPENAI_API_KEY
 *
 * Optional Script Properties:
 * OPENAI_TRANSCRIBE_MODEL = gpt-4o-mini-transcribe
 * OPENAI_TEXT_MODEL = gpt-4.1-mini
 */

const JLM_CONFIG = {
  SONGS_SHEET: "Songs",
  QUEUE_SHEET: "Metadata Lyrics Queue",
  LOG_SHEET: "AI Logs",

  TITLE_HEADER: "Song Title",
  RELEASE_HEADER: "EP / Release Name",
  STATUS_HEADER: "Workflow Status",
  FOLDER_HEADER: "Folder Link",

  MAX_AUDIO_BYTES: 24 * 1024 * 1024, // keep under common 25MB API upload limit
  MAX_QUEUE_ITEMS_PER_RUN: 2,

  AUDIO_EXTENSIONS: [
    ".mp3", ".m4a", ".wav", ".webm", ".flac", ".ogg", ".mp4", ".mpeg", ".mpga"
  ],

  GENERATED_FILES: {
    RAW_LYRICS: "lyrics_raw.txt",
    CLEAN_LYRICS: "lyrics_clean_distrokid.txt",
    METADATA: "metadata.txt",
    DISTROKID_NOTES: "distrokid_upload_notes.txt",
    RIGHTS_REVIEW: "rights_review_needed.txt"
  },

  AI_COLUMNS: [
    "Lyrics Raw",
    "Lyrics Clean",
    "Metadata File",
    "DistroKid Notes",
    "Rights Review",
    "AI Generation Status",
    "AI Last Run",
    "AI Error",
    "Explicit Review",
    "Cover Song Review",
    "Samples Review",
    "Songwriter Review",
    "Publisher Review",
    "Master Owner Review",
    "ISRC Review",
    "Audio File Used"
  ],

  QUEUE_HEADERS: [
    "Queue ID",
    "Song Row",
    "Release",
    "Song Title",
    "Folder Link",
    "Audio File",
    "Audio File ID",
    "Queue Status",
    "Last Attempt",
    "Error",
    "Generated Files",
    "Review Notes"
  ]
};


/**
 * Adds custom menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("JLM AI Tools")
    .createMenu("Pipeline")
    .addItem("1. Install / Repair AI Columns", "installJlmAiEngine")
    .addSeparator()
    .addItem("Generate Lyrics + Metadata for Selected Song", "generateLyricsMetadataForSelectedSong")
    .addItem("Queue Current Release", "queueCurrentRelease")
    .addItem("Process Catalog Metadata Queue", "processCatalogMetadataQueue")
    .addSeparator()
    .addItem("Rebuild Whole Catalog Queue", "rebuildMetadataLyricsQueue")
    .addItem("Clear Completed Queue Items", "clearCompletedQueueItems")
    .addToUi();
}


/**
 * Run this first.
 */
function installJlmAiEngine() {
  const ss = SpreadsheetApp.getActive();
  const songsSheet = getSheetOrThrow_(ss, JLM_CONFIG.SONGS_SHEET);

  ensureColumns_(songsSheet, JLM_CONFIG.AI_COLUMNS);
  ensureQueueSheet_(ss);
  ensureLogSheet_(ss);

  SpreadsheetApp.getActive().toast(
    "JLM AI Engine installed. Add OPENAI_API_KEY in Script Properties before generating.",
    "JLM AI Tools"
  );
}


/**
 * Queue every catalog song that has a folder link and is missing generated files.
 */
function rebuildMetadataLyricsQueue() {
  const ss = SpreadsheetApp.getActive();
  const songsSheet = getSheetOrThrow_(ss, JLM_CONFIG.SONGS_SHEET);
  const queueSheet = ensureQueueSheet_(ss);

  const map = getHeaderMap_(songsSheet);
  requireHeaders_(map, [
    JLM_CONFIG.TITLE_HEADER,
    JLM_CONFIG.RELEASE_HEADER,
    JLM_CONFIG.FOLDER_HEADER
  ]);

  const values = songsSheet.getDataRange().getValues();
  let added = 0;

  for (let r = 2; r <= values.length; r++) {
    const row = values[r - 1];

    const title = String(row[map[JLM_CONFIG.TITLE_HEADER] - 1] || "").trim();
    const release = String(row[map[JLM_CONFIG.RELEASE_HEADER] - 1] || "").trim();
    const folderLink = String(row[map[JLM_CONFIG.FOLDER_HEADER] - 1] || "").trim();

    if (!title || !folderLink) continue;
    if (isLikelyReleaseHeaderRow_(title, release, folderLink)) continue;

    const status = getRowAiStatus_(songsSheet, r);
    if (status === "Approved" || status === "Generated / Needs Review") continue;

    const exists = queueItemExists_(queueSheet, r, title);
    if (exists) continue;

    const audio = findBestAudioInFolderLink_(folderLink, title);

    appendQueueRow_(queueSheet, {
      songRow: r,
      release: release,
      title: title,
      folderLink: folderLink,
      audioName: audio ? audio.getName() : "",
      audioId: audio ? audio.getId() : "",
      status: audio ? "Queued" : "Blocked - No Audio Found",
      error: audio ? "" : "No supported audio file found in song folder."
    });

    added++;
  }

  SpreadsheetApp.getActive().toast(
    "Queue rebuilt. Added " + added + " item(s).",
    "JLM AI Tools"
  );
}


/**
 * Queue only the active row's release.
 */
function queueCurrentRelease() {
  const ss = SpreadsheetApp.getActive();
  const songsSheet = getSheetOrThrow_(ss, JLM_CONFIG.SONGS_SHEET);
  const queueSheet = ensureQueueSheet_(ss);

  const activeRow = songsSheet.getActiveRange().getRow();
  if (activeRow < 2) {
    SpreadsheetApp.getUi().alert("Select a song row first.");
    return;
  }

  const map = getHeaderMap_(songsSheet);
  requireHeaders_(map, [
    JLM_CONFIG.TITLE_HEADER,
    JLM_CONFIG.RELEASE_HEADER,
    JLM_CONFIG.FOLDER_HEADER
  ]);

  const release = String(songsSheet.getRange(activeRow, map[JLM_CONFIG.RELEASE_HEADER]).getValue() || "").trim();

  if (!release) {
    SpreadsheetApp.getUi().alert("The selected row does not have a release name.");
    return;
  }

  const values = songsSheet.getDataRange().getValues();
  let added = 0;

  for (let r = 2; r <= values.length; r++) {
    const row = values[r - 1];

    const rowRelease = String(row[map[JLM_CONFIG.RELEASE_HEADER] - 1] || "").trim();
    const title = String(row[map[JLM_CONFIG.TITLE_HEADER] - 1] || "").trim();
    const folderLink = String(row[map[JLM_CONFIG.FOLDER_HEADER] - 1] || "").trim();

    if (rowRelease !== release) continue;
    if (!title || !folderLink) continue;
    if (queueItemExists_(queueSheet, r, title)) continue;

    const audio = findBestAudioInFolderLink_(folderLink, title);

    appendQueueRow_(queueSheet, {
      songRow: r,
      release: rowRelease,
      title: title,
      folderLink: folderLink,
      audioName: audio ? audio.getName() : "",
      audioId: audio ? audio.getId() : "",
      status: audio ? "Queued" : "Blocked - No Audio Found",
      error: audio ? "" : "No supported audio file found in song folder."
    });

    added++;
  }

  SpreadsheetApp.getActive().toast(
    "Queued " + added + " song(s) for release: " + release,
    "JLM AI Tools"
  );
}


/**
 * Generate lyrics + metadata for the selected song row.
 */
function generateLyricsMetadataForSelectedSong() {
  const ss = SpreadsheetApp.getActive();
  const songsSheet = getSheetOrThrow_(ss, JLM_CONFIG.SONGS_SHEET);
  const row = songsSheet.getActiveRange().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert("Select a song row first.");
    return;
  }

  installJlmAiEngine();

  const result = processSongRow_(songsSheet, row);

  SpreadsheetApp.getActive().toast(
    result.message,
    "JLM AI Tools"
  );
}


/**
 * Process queued songs safely.
 * Run repeatedly until queue is complete.
 */
function processCatalogMetadataQueue() {
  installJlmAiEngine();

  const ss = SpreadsheetApp.getActive();
  const songsSheet = getSheetOrThrow_(ss, JLM_CONFIG.SONGS_SHEET);
  const queueSheet = ensureQueueSheet_(ss);

  const values = queueSheet.getDataRange().getValues();
  if (values.length < 2) {
    SpreadsheetApp.getActive().toast("Queue is empty.", "JLM AI Tools");
    return;
  }

  const headers = getHeaderMap_(queueSheet);
  let processed = 0;

  for (let r = 2; r <= values.length; r++) {
    if (processed >= JLM_CONFIG.MAX_QUEUE_ITEMS_PER_RUN) break;

    const status = String(queueSheet.getRange(r, headers["Queue Status"]).getValue() || "").trim();
    if (status !== "Queued" && status !== "Retry") continue;

    const songRow = Number(queueSheet.getRange(r, headers["Song Row"]).getValue());
    if (!songRow || songRow < 2) {
      setQueueStatus_(queueSheet, r, "Blocked", "Invalid song row.");
      continue;
    }

    try {
      setQueueStatus_(queueSheet, r, "Processing", "");

      const result = processSongRow_(songsSheet, songRow);

      if (result.ok) {
        setQueueStatus_(queueSheet, r, "Complete", "");
        queueSheet.getRange(r, headers["Generated Files"]).setValue(result.files.join("\n"));
      } else {
        setQueueStatus_(queueSheet, r, "Blocked", result.message);
      }

      processed++;
    } catch (err) {
      setQueueStatus_(queueSheet, r, "Error", String(err && err.message ? err.message : err));
      processed++;
    }
  }

  SpreadsheetApp.getActive().toast(
    "Processed " + processed + " queue item(s). Run again for more.",
    "JLM AI Tools"
  );
}


/**
 * Clear completed queue items.
 */
function clearCompletedQueueItems() {
  const ss = SpreadsheetApp.getActive();
  const queueSheet = ensureQueueSheet_(ss);
  const map = getHeaderMap_(queueSheet);
  const values = queueSheet.getDataRange().getValues();

  for (let r = values.length; r >= 2; r--) {
    const status = String(queueSheet.getRange(r, map["Queue Status"]).getValue() || "").trim();
    if (status === "Complete") {
      queueSheet.deleteRow(r);
    }
  }

  SpreadsheetApp.getActive().toast("Completed queue items cleared.", "JLM AI Tools");
}


/**
 * Main worker for one song row.
 */
function processSongRow_(songsSheet, rowNumber) {
  const map = getHeaderMap_(songsSheet);

  requireHeaders_(map, [
    JLM_CONFIG.TITLE_HEADER,
    JLM_CONFIG.RELEASE_HEADER,
    JLM_CONFIG.FOLDER_HEADER
  ]);

  const title = String(songsSheet.getRange(rowNumber, map[JLM_CONFIG.TITLE_HEADER]).getValue() || "").trim();
  const release = String(songsSheet.getRange(rowNumber, map[JLM_CONFIG.RELEASE_HEADER]).getValue() || "").trim();
  const folderLink = String(songsSheet.getRange(rowNumber, map[JLM_CONFIG.FOLDER_HEADER]).getValue() || "").trim();

  if (!title) return { ok: false, message: "Missing song title.", files: [] };
  if (!folderLink) return { ok: false, message: "Missing Folder Link.", files: [] };

  updateSongAiStatus_(songsSheet, rowNumber, {
    "AI Generation Status": "Processing",
    "AI Last Run": new Date(),
    "AI Error": ""
  });

  const folderId = extractDriveId_(folderLink);
  const songFolder = DriveApp.getFolderById(folderId);

  const audioFile = findBestAudioInFolder_(songFolder, title);
  if (!audioFile) {
    updateSongAiStatus_(songsSheet, rowNumber, {
      "AI Generation Status": "Blocked - No Audio Found",
      "AI Error": "No supported audio file found."
    });
    return { ok: false, message: "No supported audio file found.", files: [] };
  }

  updateSongAiStatus_(songsSheet, rowNumber, {
    "Audio File Used": audioFile.getName()
  });

  if (audioFile.getSize() > JLM_CONFIG.MAX_AUDIO_BYTES) {
    const note =
      "Audio file is too large for direct transcription in this workflow.\n\n" +
      "Song: " + title + "\n" +
      "Audio: " + audioFile.getName() + "\n" +
      "Size MB: " + roundMb_(audioFile.getSize()) + "\n\n" +
      "Next action: create a smaller MP3/M4A transcription copy under 24 MB, then rerun generation.";

    const blockedFile = replaceTextFile_(songFolder, "ai_blocked_audio_too_large.txt", note);

    updateSongAiStatus_(songsSheet, rowNumber, {
      "AI Generation Status": "Blocked - Audio Too Large",
      "AI Error": "Audio file over size limit. Created ai_blocked_audio_too_large.txt."
    });

    return {
      ok: false,
      message: "Blocked: audio too large. Created note file.",
      files: [blockedFile.getUrl()]
    };
  }

  const transcript = openAiTranscribe_(audioFile, title);
  if (!transcript) {
    updateSongAiStatus_(songsSheet, rowNumber, {
      "AI Generation Status": "Error",
      "AI Error": "Transcription returned empty text."
    });
    return { ok: false, message: "Transcription returned empty text.", files: [] };
  }

  const draft = openAiDraftLyricsMetadata_({
    songTitle: title,
    releaseName: release,
    artist: "Jim Lamb Music",
    audioFileName: audioFile.getName(),
    transcript: transcript
  });

  const generated = [];

  const rawLyricsFile = replaceTextFile_(
    songFolder,
    JLM_CONFIG.GENERATED_FILES.RAW_LYRICS,
    transcript
  );
  generated.push(rawLyricsFile.getUrl());

  const cleanLyricsFile = replaceTextFile_(
    songFolder,
    JLM_CONFIG.GENERATED_FILES.CLEAN_LYRICS,
    draft.cleanLyrics || "NEEDS REVIEW\n\n" + transcript
  );
  generated.push(cleanLyricsFile.getUrl());

  const metadataFile = replaceTextFile_(
    songFolder,
    JLM_CONFIG.GENERATED_FILES.METADATA,
    draft.metadata || buildFallbackMetadata_(title, release, audioFile.getName())
  );
  generated.push(metadataFile.getUrl());

  const notesFile = replaceTextFile_(
    songFolder,
    JLM_CONFIG.GENERATED_FILES.DISTROKID_NOTES,
    draft.distrokidNotes || buildDefaultDistroKidNotes_(title)
  );
  generated.push(notesFile.getUrl());

  const rightsFile = replaceTextFile_(
    songFolder,
    JLM_CONFIG.GENERATED_FILES.RIGHTS_REVIEW,
    draft.rightsReview || buildDefaultRightsReview_(title)
  );
  generated.push(rightsFile.getUrl());

  updateSongAiStatus_(songsSheet, rowNumber, {
    "Lyrics Raw": rawLyricsFile.getUrl(),
    "Lyrics Clean": cleanLyricsFile.getUrl(),
    "Metadata File": metadataFile.getUrl(),
    "DistroKid Notes": notesFile.getUrl(),
    "Rights Review": rightsFile.getUrl(),
    "AI Generation Status": "Generated / Needs Review",
    "AI Last Run": new Date(),
    "AI Error": "",
    "Explicit Review": "Review Needed",
    "Cover Song Review": "Review Needed",
    "Samples Review": "Review Needed",
    "Songwriter Review": "Review Needed",
    "Publisher Review": "Review Needed",
    "Master Owner Review": "Review Needed",
    "ISRC Review": "Review Needed"
  });

  logAi_(songsSheet.getParent(), "Generated", title, "Created lyrics, metadata, DistroKid notes, rights review.");

  return {
    ok: true,
    message: "Generated lyrics + metadata for " + title + ". Needs review.",
    files: generated
  };
}


/**
 * OpenAI audio transcription.
 */
function openAiTranscribe_(audioFile, songTitle) {
  const apiKey = getOpenAiApiKey_();
  const model = getScriptProperty_("OPENAI_TRANSCRIBE_MODEL", "gpt-4o-mini-transcribe");

  const blob = audioFile.getBlob().setName(audioFile.getName());

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions", {
    method: "post",
    headers: {
      Authorization: "Bearer " + apiKey
    },
    payload: {
      model: model,
      file: blob,
      response_format: "json",
      prompt: "This is a country/Americana song by Jim Lamb Music. Transcribe the sung lyrics as accurately as possible for the song titled: " + songTitle
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error("OpenAI transcription failed: HTTP " + code + " — " + text);
  }

  const data = JSON.parse(text);
  return String(data.text || "").trim();
}


/**
 * OpenAI metadata + cleaned lyrics drafting.
 */
function openAiDraftLyricsMetadata_(info) {
  const apiKey = getOpenAiApiKey_();
  const model = getScriptProperty_("OPENAI_TEXT_MODEL", "gpt-4.1-mini");

  const prompt =
    "You are helping prepare a music release package for DistroKid.\n\n" +
    "Artist: Jim Lamb Music\n" +
    "Release: " + info.releaseName + "\n" +
    "Song Title: " + info.songTitle + "\n" +
    "Audio File: " + info.audioFileName + "\n\n" +
    "Transcript:\n" +
    info.transcript + "\n\n" +
    "Create draft outputs for this song.\n\n" +
    "Rules:\n" +
    "- Do not claim rights, ownership, publishing, songwriter splits, ISRC, UPC, sample clearance, cover-song status, or explicit status as final.\n" +
    "- Mark all rights/licensing fields as Review Needed.\n" +
    "- Clean lyrics for plain lyric submission: remove timestamps, avoid weird spacing, avoid all-caps except proper words, use stanza breaks.\n" +
    "- If lyrics are uncertain, mark [unclear] sparingly.\n" +
    "- Keep metadata practical for a music release dashboard.\n\n" +
    "Return EXACTLY in this format with these headings:\n\n" +
    "=== CLEAN_LYRICS ===\n" +
    "[cleaned lyrics]\n\n" +
    "=== METADATA ===\n" +
    "Song Title:\n" +
    "Artist:\n" +
    "Release:\n" +
    "Audio File:\n" +
    "Language:\n" +
    "Genre:\n" +
    "Subgenre:\n" +
    "Mood:\n" +
    "Themes:\n" +
    "Short Description:\n" +
    "Explicit: Review Needed\n" +
    "Songwriter(s): Review Needed\n" +
    "Publisher: Review Needed\n" +
    "Master Owner: Review Needed\n" +
    "Composition Owner: Review Needed\n" +
    "Featured Artist: Review Needed\n" +
    "ISRC: Review Needed\n" +
    "UPC: Album-level / DistroKid generated unless reused\n" +
    "Cover Song?: Review Needed\n" +
    "Samples?: Review Needed\n" +
    "Ready Status: Needs Review\n\n" +
    "=== DISTROKID_NOTES ===\n" +
    "[upload notes]\n\n" +
    "=== RIGHTS_REVIEW ===\n" +
    "[rights and licensing checklist]\n";

  const body = {
    model: model,
    input: prompt
  };

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
    method: "post",
    headers: {
      Authorization: "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error("OpenAI metadata draft failed: HTTP " + code + " — " + text);
  }

  const data = JSON.parse(text);
  const outputText = extractResponseText_(data);

  return {
    cleanLyrics: extractSection_(outputText, "CLEAN_LYRICS"),
    metadata: extractSection_(outputText, "METADATA"),
    distrokidNotes: extractSection_(outputText, "DISTROKID_NOTES"),
    rightsReview: extractSection_(outputText, "RIGHTS_REVIEW"),
    fullText: outputText
  };
}


/**
 * Finds best audio file inside a folder.
 */
function findBestAudioInFolderLink_(folderLink, songTitle) {
  const folderId = extractDriveId_(folderLink);
  if (!folderId) return null;

  const folder = DriveApp.getFolderById(folderId);
  return findBestAudioInFolder_(folder, songTitle);
}


function findBestAudioInFolder_(folder, songTitle) {
  const candidates = [];

  scanFolderForAudio_(folder, songTitle, candidates, 0);

  if (!candidates.length) return null;

  candidates.sort(function(a, b) {
    return b.score - a.score;
  });

  return candidates[0].file;
}


function scanFolderForAudio_(folder, songTitle, candidates, depth) {
  if (depth > 2) return;

  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const lower = name.toLowerCase();

    if (!isSupportedAudioName_(lower)) continue;

    let score = 0;

    if (lower.indexOf("master") !== -1) score += 100;
    if (lower.indexOf("final") !== -1) score += 50;
    if (lower.indexOf("preview") !== -1) score -= 75;
    if (lower.indexOf("spectrogram") !== -1) score -= 100;
    if (lower.indexOf("waveform") !== -1) score -= 100;

    const normalizedTitle = normalizeText_(songTitle);
    const normalizedName = normalizeText_(name);

    if (normalizedName.indexOf(normalizedTitle) !== -1) score += 75;

    if (lower.endsWith(".wav") || lower.endsWith(".flac")) score += 40;
    if (lower.endsWith(".mp3") || lower.endsWith(".m4a")) score += 25;
    if (String(file.getMimeType()).indexOf("audio/") === 0) score += 30;

    score += Math.min(20, Math.floor(file.getSize() / (1024 * 1024)));

    candidates.push({
      file: file,
      score: score
    });
  }

  const folders = folder.getFolders();
  while (folders.hasNext()) {
    scanFolderForAudio_(folders.next(), songTitle, candidates, depth + 1);
  }
}


function isSupportedAudioName_(lowerName) {
  return JLM_CONFIG.AUDIO_EXTENSIONS.some(function(ext) {
    return lowerName.endsWith(ext);
  });
}


/**
 * File writing.
 * Existing generated files are moved into _AI Generated Archive, not deleted.
 */
function replaceTextFile_(folder, fileName, contents) {
  archiveExistingFiles_(folder, fileName);

  return folder.createFile(
    fileName,
    contents,
    MimeType.PLAIN_TEXT
  );
}


function archiveExistingFiles_(folder, fileName) {
  const existing = folder.getFilesByName(fileName);
  if (!existing.hasNext()) return;

  const archiveFolder = getOrCreateChildFolder_(folder, "_AI Generated Archive");
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");

  while (existing.hasNext()) {
    const file = existing.next();
    file.setName("ARCHIVE " + stamp + " - " + fileName);
    file.moveTo(archiveFolder);
  }
}


function getOrCreateChildFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}


/**
 * Sheet setup.
 */
function ensureColumns_(sheet, headersToEnsure) {
  const existing = getHeaderMap_(sheet);
  let lastCol = sheet.getLastColumn();

  headersToEnsure.forEach(function(header) {
    if (!existing[header]) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(header);
      existing[header] = lastCol;
    }
  });
}


function ensureQueueSheet_(ss) {
  let sheet = ss.getSheetByName(JLM_CONFIG.QUEUE_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(JLM_CONFIG.QUEUE_SHEET);
  }

  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === "") {
    sheet.getRange(1, 1, 1, JLM_CONFIG.QUEUE_HEADERS.length).setValues([JLM_CONFIG.QUEUE_HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}


function ensureLogSheet_(ss) {
  let sheet = ss.getSheetByName(JLM_CONFIG.LOG_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(JLM_CONFIG.LOG_SHEET);
    sheet.getRange(1, 1, 1, 5).setValues([[
      "Timestamp",
      "Type",
      "Song",
      "Message",
      "User"
    ]]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}


/**
 * Queue helpers.
 */
function appendQueueRow_(queueSheet, item) {
  queueSheet.appendRow([
    Utilities.getUuid(),
    item.songRow,
    item.release,
    item.title,
    item.folderLink,
    item.audioName,
    item.audioId,
    item.status,
    "",
    item.error || "",
    "",
    "AI output must be human-reviewed before DistroKid upload."
  ]);
}


function queueItemExists_(queueSheet, songRow, title) {
  const values = queueSheet.getDataRange().getValues();
  if (values.length < 2) return false;

  const map = getHeaderMap_(queueSheet);

  for (let r = 2; r <= values.length; r++) {
    const rowSongRow = Number(queueSheet.getRange(r, map["Song Row"]).getValue());
    const rowTitle = String(queueSheet.getRange(r, map["Song Title"]).getValue() || "").trim();
    const rowStatus = String(queueSheet.getRange(r, map["Queue Status"]).getValue() || "").trim();

    if (rowSongRow === songRow && rowTitle === title && rowStatus !== "Complete") {
      return true;
    }
  }

  return false;
}


function setQueueStatus_(queueSheet, row, status, error) {
  const map = getHeaderMap_(queueSheet);

  queueSheet.getRange(row, map["Queue Status"]).setValue(status);
  queueSheet.getRange(row, map["Last Attempt"]).setValue(new Date());

  if (error !== undefined) {
    queueSheet.getRange(row, map["Error"]).setValue(error);
  }
}


/**
 * Song row status helpers.
 */
function getRowAiStatus_(songsSheet, row) {
  const map = getHeaderMap_(songsSheet);
  if (!map["AI Generation Status"]) return "";
  return String(songsSheet.getRange(row, map["AI Generation Status"]).getValue() || "").trim();
}


function updateSongAiStatus_(songsSheet, rowNumber, updates) {
  ensureColumns_(songsSheet, Object.keys(updates));

  const map = getHeaderMap_(songsSheet);

  Object.keys(updates).forEach(function(header) {
    songsSheet.getRange(rowNumber, map[header]).setValue(updates[header]);
  });
}


/**
 * Parsing helpers.
 */
function extractResponseText_(data) {
  if (data.output_text) return String(data.output_text).trim();

  let text = "";

  if (data.output && Array.isArray(data.output)) {
    data.output.forEach(function(item) {
      if (item.content && Array.isArray(item.content)) {
        item.content.forEach(function(contentItem) {
          if (contentItem.text) text += contentItem.text + "\n";
          if (contentItem.type === "output_text" && contentItem.text) text += contentItem.text + "\n";
        });
      }
    });
  }

  return text.trim();
}


function extractSection_(text, sectionName) {
  const pattern = new RegExp(
    "===\\s*" + sectionName + "\\s*===\\s*([\\s\\S]*?)(?=\\n===\\s*[A-Z_]+\\s*===|$)",
    "i"
  );

  const match = text.match(pattern);
  return match ? match[1].trim() : "";
}


/**
 * Fallback builders.
 */
function buildFallbackMetadata_(title, release, audioFile) {
  return [
    "Song Title: " + title,
    "Artist: Jim Lamb Music",
    "Release: " + release,
    "Audio File: " + audioFile,
    "Language: English",
    "Genre: Country / Americana",
    "Subgenre: Review Needed",
    "Mood: Review Needed",
    "Themes: Review Needed",
    "Short Description: Review Needed",
    "Explicit: Review Needed",
    "Songwriter(s): Review Needed",
    "Publisher: Review Needed",
    "Master Owner: Review Needed",
    "Composition Owner: Review Needed",
    "Featured Artist: Review Needed",
    "ISRC: Review Needed",
    "UPC: Album-level / DistroKid generated unless reused",
    "Cover Song?: Review Needed",
    "Samples?: Review Needed",
    "Ready Status: Needs Review"
  ].join("\n");
}


function buildDefaultDistroKidNotes_(title) {
  return [
    "DistroKid Upload Notes",
    "",
    "Song: " + title,
    "Artist: Jim Lamb Music",
    "",
    "Before upload, verify:",
    "- Final song title spelling",
    "- Songwriter legal name(s)",
    "- Publisher information",
    "- Master owner",
    "- Composition owner",
    "- Explicit status",
    "- Cover song status",
    "- Sample clearance",
    "- ISRC reuse if previously released",
    "",
    "Status: Needs Review"
  ].join("\n");
}


function buildDefaultRightsReview_(title) {
  return [
    "Rights / Licensing Review Needed",
    "",
    "Song: " + title,
    "",
    "Confirm before DistroKid upload:",
    "- Is this an original song?",
    "- Is this a cover song?",
    "- Does it use any samples?",
    "- Who wrote the composition?",
    "- Who owns the master recording?",
    "- Who owns publishing?",
    "- Any featured artist permissions needed?",
    "- Existing ISRC to reuse?",
    "- Explicit lyrics?",
    "",
    "AI cannot verify rights. Human review required."
  ].join("\n");
}


/**
 * Generic helpers.
 */
function getSheetOrThrow_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error("Missing sheet: " + name);
  return sheet;
}


function getHeaderMap_(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const map = {};
  headers.forEach(function(header, index) {
    const key = String(header || "").trim();
    if (key) map[key] = index + 1;
  });

  return map;
}


function requireHeaders_(map, headers) {
  const missing = headers.filter(function(h) {
    return !map[h];
  });

  if (missing.length) {
    throw new Error("Missing required header(s): " + missing.join(", "));
  }
}


function extractDriveId_(urlOrId) {
  const text = String(urlOrId || "").trim();

  if (!text) return "";

  const folderMatch = text.match(/folders\/([a-zA-Z0-9_-]+)/);
  if (folderMatch) return folderMatch[1];

  const fileMatch = text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (fileMatch) return fileMatch[1];

  const idMatch = text.match(/[a-zA-Z0-9_-]{20,}/);
  return idMatch ? idMatch[0] : "";
}


function normalizeText_(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/\(feat[^)]*\)/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}


function isLikelyReleaseHeaderRow_(title, release, folderLink) {
  return title && !release && folderLink;
}


function getOpenAiApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  if (!key) {
    throw new Error("Missing Script Property: OPENAI_API_KEY");
  }

  return key;
}


function getScriptProperty_(name, fallback) {
  return PropertiesService.getScriptProperties().getProperty(name) || fallback;
}


function roundMb_(bytes) {
  return Math.round((bytes / 1024 / 1024) * 10) / 10;
}


function logAi_(ss, type, song, message) {
  const sheet = ensureLogSheet_(ss);

  sheet.appendRow([
    new Date(),
    type,
    song,
    message,
    Session.getActiveUser().getEmail()
  ]);
}
/**
 * JLM WAV Intake Scanner
 *
 * Purpose:
 * When a new WAV/audio file is saved into Jim Lamb Music Drive folders,
 * this scanner adds or updates the song row in the Songs tab.
 *
 * Run first:
 * jlmInstallWavIntakeAutomation
 *
 * Then it scans every 5 minutes.
 */

const JLM_WAV_INTAKE_CONFIG = {
  SONGS_SHEET: "Songs",
  LOG_SHEET: "Drive WAV Intake Log",

  // Canonical Jim Lamb Music folders
  STAGE_FOLDERS: [
    {
      name: "01-Pre-Production",
      id: "1WAgXnoYlk9WgvVxxG_vUdE2h1_v-ISZN",
      status: "Pre-Production"
    },
    {
      name: "02-In Progress",
      id: "1i6znQDeA7aMVpIDxeob_OARl4nnQvf9q",
      status: "In Progress"
    },
    {
      name: "03-Released",
      id: "1fu89V9-mPyPBm1ALL85VtdnwS0w-13bC",
      status: "Released"
    }
  ],

  AUDIO_EXTENSIONS: [
    ".wav", ".mp3", ".m4a", ".flac", ".aiff", ".aif", ".ogg", ".webm", ".mp4", ".mpeg", ".mpga"
  ],

  REQUIRED_SONG_COLUMNS: [
    "Song Title",
    "EP / Release Name",
    "Suggested Release Date",
    "Workflow Status",
    "Art?",
    "Video?",
    "Posted?",
    "DistroKid?",
    "Streaming?",
    "Notes",
    "Folder Link",
    "Migration Status",
    "lyrics",
    "isrc",
    "upc",
    "RAW WAV?",
    "MASTER WAV?",
    "Naming Issue",
    "Last Asset Scan",
    "Source Audio File ID",
    "Source Audio File URL",
    "Source Audio File Name",
    "Source Audio Folder",
    "Drive Intake Status",
    "Drive Intake Last Run"
  ],

  LOG_HEADERS: [
    "Timestamp",
    "Action",
    "Song Title",
    "Release",
    "Audio File",
    "Audio File ID",
    "Song Folder",
    "Message"
  ],

  MAX_DEPTH: 4
};


/**
 * Run this once.
 * It installs columns, creates log tab, and creates a 5-minute trigger.
 */
function jlmInstallWavIntakeAutomation() {
  const ss = SpreadsheetApp.getActive();
  const songsSheet = jlmWavGetSheetOrThrow_(ss, JLM_WAV_INTAKE_CONFIG.SONGS_SHEET);

  jlmWavEnsureColumns_(songsSheet, JLM_WAV_INTAKE_CONFIG.REQUIRED_SONG_COLUMNS);
  jlmWavEnsureLogSheet_(ss);

  // Remove old duplicate triggers first
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "jlmScanForNewWavFiles") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("jlmScanForNewWavFiles")
    .timeBased()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getActive().toast(
    "WAV intake automation installed. New audio files will be scanned every 5 minutes.",
    "JLM WAV Intake"
  );
}


/**
 * Manual scan button/function.
 * You can run this any time.
 */
function jlmScanForNewWavFiles() {
  const ss = SpreadsheetApp.getActive();
  const songsSheet = jlmWavGetSheetOrThrow_(ss, JLM_WAV_INTAKE_CONFIG.SONGS_SHEET);

  jlmWavEnsureColumns_(songsSheet, JLM_WAV_INTAKE_CONFIG.REQUIRED_SONG_COLUMNS);
  jlmWavEnsureLogSheet_(ss);

  let stats = {
    scanned: 0,
    added: 0,
    updated: 0,
    moved: 0,
    skipped: 0,
    errors: 0
  };

  JLM_WAV_INTAKE_CONFIG.STAGE_FOLDERS.forEach(function(stage) {
    try {
      const stageFolder = DriveApp.getFolderById(stage.id);
      jlmWavScanFolderRecursive_(ss, songsSheet, stage, stageFolder, [], stats, 0);
    } catch (err) {
      stats.errors++;
      jlmWavLog_(ss, "ERROR", "", "", "", "", "", "Stage scan failed for " + stage.name + ": " + err.message);
    }
  });

  SpreadsheetApp.getActive().toast(
    "WAV scan done. Added: " + stats.added +
    ", Updated: " + stats.updated +
    ", Moved: " + stats.moved +
    ", Scanned: " + stats.scanned,
    "JLM WAV Intake"
  );
}


/**
 * Optional: remove the automatic trigger.
 */
function jlmRemoveWavIntakeAutomation() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "jlmScanForNewWavFiles") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  SpreadsheetApp.getActive().toast(
    "WAV intake trigger removed.",
    "JLM WAV Intake"
  );
}


/**
 * Recursively scans folders for supported audio files.
 */
function jlmWavScanFolderRecursive_(ss, songsSheet, stage, folder, relativePath, stats, depth) {
  if (depth > JLM_WAV_INTAKE_CONFIG.MAX_DEPTH) return;

  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();

    if (!jlmWavIsSupportedAudioFile_(file)) {
      continue;
    }

    stats.scanned++;

    try {
      const result = jlmWavProcessAudioFile_(ss, songsSheet, stage, folder, relativePath, file);

      if (result.action === "added") stats.added++;
      if (result.action === "updated") stats.updated++;
      if (result.moved) stats.moved++;
      if (result.action === "skipped") stats.skipped++;

    } catch (err) {
      stats.errors++;
      jlmWavLog_(
        ss,
        "ERROR",
        "",
        "",
        file.getName(),
        file.getId(),
        folder.getName(),
        err.message
      );
    }
  }

  const folders = folder.getFolders();

  while (folders.hasNext()) {
    const child = folders.next();
    const nextPath = relativePath.concat([{
      id: child.getId(),
      name: child.getName()
    }]);

    jlmWavScanFolderRecursive_(ss, songsSheet, stage, child, nextPath, stats, depth + 1);
  }
}


/**
 * Processes one audio file:
 * - Detects release
 * - Detects song title
 * - Creates song folder if needed
 * - Moves loose audio into song folder
 * - Adds or updates Songs row
 */
function jlmWavProcessAudioFile_(ss, songsSheet, stage, currentFolder, relativePath, file) {
  const fileName = file.getName();
  const fileId = file.getId();

  const songTitle = jlmWavTitleFromFilename_(fileName);
  const isMaster = jlmWavIsMasterFile_(fileName);

  let releaseName = "";
  let releaseFolder = null;
  let songFolder = currentFolder;
  let moved = false;

  // Case 1: file directly inside stage folder
  // Put it under _Inbox New WAV Imports / Song Title
  if (relativePath.length === 0) {
    releaseName = "_Inbox New WAV Imports";
    releaseFolder = jlmWavGetOrCreateChildFolder_(currentFolder, releaseName);
    songFolder = jlmWavGetOrCreateChildFolder_(releaseFolder, songTitle);
    file.moveTo(songFolder);
    moved = true;
  }

  // Case 2: file directly inside album/release folder
  // Example: 02-In Progress / Americana 250 / song.wav
  // Create: 02-In Progress / Americana 250 / Song Title / song.wav
  if (relativePath.length === 1) {
    releaseName = relativePath[0].name;
    releaseFolder = currentFolder;
    songFolder = jlmWavGetOrCreateChildFolder_(releaseFolder, songTitle);
    file.moveTo(songFolder);
    moved = true;
  }

  // Case 3: file already inside song folder
  // Example: 02-In Progress / Americana 250 / Streets of LA / master.wav
  if (relativePath.length >= 2) {
    releaseName = relativePath[0].name;
    songFolder = currentFolder;
  }

  const songFolderUrl = "https://drive.google.com/drive/folders/" + songFolder.getId();
  const fileUrl = file.getUrl();

  const existingRow = jlmWavFindExistingSongRow_(songsSheet, {
    fileId: fileId,
    songTitle: songTitle,
    releaseName: releaseName
  });

  const rowData = {
    "Song Title": songTitle,
    "EP / Release Name": releaseName,
    "Workflow Status": stage.status,
    "Art?": "No",
    "Video?": "No",
    "Posted?": "No",
    "DistroKid?": "No",
    "Streaming?": stage.status === "Released" ? "Review Needed" : "No",
    "Notes": "Auto-added from Drive audio intake: " + fileName,
    "Folder Link": songFolderUrl,
    "Migration Status": "Drive Intake Linked",
    "RAW WAV?": isMaster ? false : true,
    "MASTER WAV?": isMaster ? true : false,
    "Naming Issue": jlmWavDetectNamingIssue_(songTitle, fileName),
    "Last Asset Scan": new Date(),
    "Source Audio File ID": fileId,
    "Source Audio File URL": fileUrl,
    "Source Audio File Name": fileName,
    "Source Audio Folder": songFolder.getName(),
    "Drive Intake Status": "Audio Detected",
    "Drive Intake Last Run": new Date()
  };

  if (existingRow) {
    jlmWavUpdateSongRow_(songsSheet, existingRow, rowData);

    jlmWavLog_(
      ss,
      "UPDATED",
      songTitle,
      releaseName,
      fileName,
      fileId,
      songFolder.getName(),
      "Updated existing song row " + existingRow + "."
    );

    return {
      action: "updated",
      moved: moved
    };
  }

  const newRow = songsSheet.getLastRow() + 1;
  jlmWavUpdateSongRow_(songsSheet, newRow, rowData);

  jlmWavLog_(
    ss,
    "ADDED",
    songTitle,
    releaseName,
    fileName,
    fileId,
    songFolder.getName(),
    "Added new song row " + newRow + "."
  );

  return {
    action: "added",
    moved: moved
  };
}


/**
 * Finds existing song by Source Audio File ID first,
 * then by Song Title + Release.
 */
function jlmWavFindExistingSongRow_(songsSheet, info) {
  const map = jlmWavGetHeaderMap_(songsSheet);
  const values = songsSheet.getDataRange().getValues();

  const fileIdCol = map["Source Audio File ID"];
  const titleCol = map["Song Title"];
  const releaseCol = map["EP / Release Name"];

  for (let r = 2; r <= values.length; r++) {
    const row = values[r - 1];

    if (fileIdCol) {
      const existingFileId = String(row[fileIdCol - 1] || "").trim();
      if (existingFileId && existingFileId === info.fileId) {
        return r;
      }
    }

    const existingTitle = String(row[titleCol - 1] || "").trim();
    const existingRelease = String(row[releaseCol - 1] || "").trim();

    if (
      jlmWavNormalize_(existingTitle) === jlmWavNormalize_(info.songTitle) &&
      jlmWavNormalize_(existingRelease) === jlmWavNormalize_(info.releaseName)
    ) {
      return r;
    }
  }

  return null;
}


/**
 * Writes values by header name.
 */
function jlmWavUpdateSongRow_(sheet, rowNumber, data) {
  jlmWavEnsureColumns_(sheet, Object.keys(data));

  const map = jlmWavGetHeaderMap_(sheet);

  Object.keys(data).forEach(function(header) {
    if (!map[header]) return;
    sheet.getRange(rowNumber, map[header]).setValue(data[header]);
  });
}


/**
 * Detects title from filename.
 */
function jlmWavTitleFromFilename_(fileName) {
  let base = String(fileName || "")
    .replace(/\.[^.]+$/, "")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  // Remove common production suffixes
  base = base
    .replace(/\bmaster\b/gi, "")
    .replace(/\bfinal\b/gi, "")
    .replace(/\bmix\b/gi, "")
    .replace(/\bversion\b/gi, "")
    .replace(/\bremaster\b/gi, "")
    .replace(/\bexport\b/gi, "")
    .replace(/\s+/g, " ")
    .trim();

  // Known cleanup aliases
  const normalized = jlmWavNormalize_(base);

  if (normalized === "street of la" || normalized === "streets of la") return "Streets of LA";
  if (normalized === "take a ride") return "Take a Ride";
  if (normalized === "american soldier") return "American Soldier";
  if (normalized === "white boy") return "White Boy";
  if (normalized === "crippled oldier" || normalized === "crippled soldier") return "Crippled Soldier";
  if (normalized === "alexander") return "Alexander (feat Evee)";

  return jlmWavTitleCase_(base);
}


function jlmWavTitleCase_(text) {
  return String(text || "")
    .toLowerCase()
    .split(" ")
    .filter(Boolean)
    .map(function(word) {
      if (word === "la") return "LA";
      if (word === "usa") return "USA";
      if (word === "ai") return "AI";
      return word.charAt(0).toUpperCase() + word.slice(1);
    })
    .join(" ");
}


function jlmWavIsMasterFile_(fileName) {
  const lower = String(fileName || "").toLowerCase();

  return (
    lower.indexOf("master") !== -1 ||
    lower.indexOf("final") !== -1 ||
    lower.endsWith(".wav") ||
    lower.endsWith(".flac") ||
    lower.endsWith(".aiff") ||
    lower.endsWith(".aif")
  );
}


function jlmWavDetectNamingIssue_(songTitle, fileName) {
  const titleNorm = jlmWavNormalize_(songTitle);
  const fileNorm = jlmWavNormalize_(fileName.replace(/\.[^.]+$/, ""));

  if (fileNorm.indexOf(titleNorm) === -1) {
    return "Possible filename mismatch: " + fileName;
  }

  if (fileName.indexOf("_") !== -1) {
    return "Filename uses underscores: " + fileName;
  }

  return "";
}


function jlmWavIsSupportedAudioFile_(file) {
  const name = String(file.getName() || "").toLowerCase();
  const mime = String(file.getMimeType() || "").toLowerCase();

  const extensionMatch = JLM_WAV_INTAKE_CONFIG.AUDIO_EXTENSIONS.some(function(ext) {
    return name.endsWith(ext);
  });

  return extensionMatch || mime.indexOf("audio/") === 0;
}


function jlmWavGetOrCreateChildFolder_(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();

  return parent.createFolder(name);
}


/**
 * Sheet helpers.
 */
function jlmWavEnsureColumns_(sheet, headersToEnsure) {
  const map = jlmWavGetHeaderMap_(sheet);
  let lastCol = sheet.getLastColumn();

  headersToEnsure.forEach(function(header) {
    if (!map[header]) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(header);
      map[header] = lastCol;
    }
  });
}


function jlmWavEnsureLogSheet_(ss) {
  let sheet = ss.getSheetByName(JLM_WAV_INTAKE_CONFIG.LOG_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(JLM_WAV_INTAKE_CONFIG.LOG_SHEET);
  }

  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === "") {
    sheet.getRange(1, 1, 1, JLM_WAV_INTAKE_CONFIG.LOG_HEADERS.length)
      .setValues([JLM_WAV_INTAKE_CONFIG.LOG_HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}


function jlmWavLog_(ss, action, songTitle, release, audioFile, audioFileId, songFolder, message) {
  const sheet = jlmWavEnsureLogSheet_(ss);

  sheet.appendRow([
    new Date(),
    action,
    songTitle,
    release,
    audioFile,
    audioFileId,
    songFolder,
    message
  ]);
}


function jlmWavGetSheetOrThrow_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("Missing sheet: " + sheetName);
  return sheet;
}


function jlmWavGetHeaderMap_(sheet) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const map = {};

  headers.forEach(function(header, index) {
    const key = String(header || "").trim();
    if (key) map[key] = index + 1;
  });

  return map;
}


function jlmWavNormalize_(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/\.[^.]+$/, "")
    .replace(/\(feat[^)]*\)/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();

}
function getGeminiApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

  if (!key) {
    throw new Error("Missing Script Property: GEMINI_API_KEY");
  }

  return key;
}
function callGeminiForLyricsMetadata_(audioFile, songTitle, releaseName) {
  const apiKey = getGeminiApiKey_();

  const maxBytes = 18 * 1024 * 1024;
  if (audioFile.getSize() > maxBytes) {
    throw new Error(
      "Audio file is too large for inline Gemini processing. Create a smaller MP3/M4A transcription copy under 18 MB."
    );
  }

  const blob = audioFile.getBlob();
  const base64Audio = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType() || "audio/wav";

  const prompt =
    "You are preparing a music release package for Jim Lamb Music.\n\n" +
    "Song Title: " + songTitle + "\n" +
    "Release / Album: " + releaseName + "\n" +
    "Artist: Jim Lamb Music\n\n" +
    "Tasks:\n" +
    "1. Transcribe the lyrics as accurately as possible.\n" +
    "2. Create cleaned DistroKid-friendly plain lyrics.\n" +
    "3. Create metadata for the song.\n" +
    "4. Create DistroKid upload notes.\n" +
    "5. Create a rights/licensing review checklist.\n\n" +
    "Important rules:\n" +
    "- Do not mark songwriter, publishing, master ownership, composition ownership, ISRC, UPC, cover-song status, sample clearance, or explicit status as final.\n" +
    "- Mark those fields as Review Needed.\n" +
    "- If any lyric is unclear, use [unclear] sparingly.\n\n" +
    "Return exactly these sections:\n\n" +
    "=== RAW_LYRICS ===\n" +
    "=== CLEAN_LYRICS ===\n" +
    "=== METADATA ===\n" +
    "=== DISTROKID_NOTES ===\n" +
    "=== RIGHTS_REVIEW ===";

  const payload = {
    model: "gemini-3.5-flash",
    input: [
      {
        type: "text",
        text: prompt
      },
      {
        type: "audio",
        data: base64Audio,
        mime_type: mimeType
      }
    ]
  };

  const response = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/interactions",
    {
      method: "post",
      contentType: "application/json",
      headers: {
        "x-goog-api-key": apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error("Gemini API failed: HTTP " + code + " — " + text);
  }

  const data = JSON.parse(text);
  return String(data.output_text || "").trim();
}
function jlmShowRootAudioScanFolder() {
  const folder = DriveApp.getRootFolder();

  Logger.log("Scanning ONLY the top level of My Drive:");
  Logger.log(folder.getName());
  Logger.log(folder.getUrl());
}