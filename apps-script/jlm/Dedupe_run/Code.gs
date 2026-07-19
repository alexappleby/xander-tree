const ROOT_ID = "1lMcL9DniWH33Nyb6dpZItQESwKtciGdy";

const IGNORE_FOLDER_NAMES = {
  ".DS_Store": true,
  "__MACOSX": true,
  "Archive": true,
  "Archives": true,
  "Trash": true
};

const PROP_PREFIX = "DEDUPE_V1_";
const SAFE_RUNTIME_MS = 4.5 * 60 * 1000;
const RESUME_DELAY_MS = 60 * 1000;

function dryRunStart() {
  resetDedupeState_();
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_PREFIX + "MODE", "DRY");
  props.setProperty(PROP_PREFIX + "ROOT_ID", ROOT_ID);
  props.setProperty(PROP_PREFIX + "STARTED_AT", new Date().toISOString());
  props.setProperty(PROP_PREFIX + "PENDING_FOLDERS", JSON.stringify([ROOT_ID]));
  props.setProperty(PROP_PREFIX + "ACTIONS_COUNT", "0");
  props.deleteProperty(PROP_PREFIX + "CURRENT_FOLDER_ID");
  props.deleteProperty(PROP_PREFIX + "CURRENT_FILE_TOKEN");
  props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_TOKEN");
  props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_QUEUE");
  dedupeResume_();
}

function applyRunStart() {
  resetDedupeState_();
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_PREFIX + "MODE", "APPLY");
  props.setProperty(PROP_PREFIX + "ROOT_ID", ROOT_ID);
  props.setProperty(PROP_PREFIX + "STARTED_AT", new Date().toISOString());
  props.setProperty(PROP_PREFIX + "PENDING_FOLDERS", JSON.stringify([ROOT_ID]));
  props.setProperty(PROP_PREFIX + "ACTIONS_COUNT", "0");
  props.deleteProperty(PROP_PREFIX + "CURRENT_FOLDER_ID");
  props.deleteProperty(PROP_PREFIX + "CURRENT_FILE_TOKEN");
  props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_TOKEN");
  props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_QUEUE");
  dedupeResume_();
}

function dedupeResume_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log("Another dedupe run is already active.");
    return;
  }

  try {
    const started = Date.now();
    const props = PropertiesService.getScriptProperties();
    const mode = props.getProperty(PROP_PREFIX + "MODE");
    const rootId = props.getProperty(PROP_PREFIX + "ROOT_ID");

    if (!mode || !rootId) {
      Logger.log("No dedupe job is active.");
      return;
    }

    const doApply = mode === "APPLY";
    const rootFolder = DriveApp.getFolderById(rootId);
    const songFolders = getSongFolders_(rootFolder);

    Logger.log("Mode: " + mode);
    Logger.log("Song folders indexed: " + songFolders.length);

    while (!isTimeUp_(started)) {
      let currentFolderId = props.getProperty(PROP_PREFIX + "CURRENT_FOLDER_ID");

      if (!currentFolderId) {
        const pending = getJsonProp_(PROP_PREFIX + "PENDING_FOLDERS", []);
        if (!pending.length) {
          Logger.log("Done. Total actions: " + props.getProperty(PROP_PREFIX + "ACTIONS_COUNT"));
          clearResumeTrigger_();
          clearWorkingState_();
          return;
        }

        currentFolderId = pending.shift();
        setJsonProp_(PROP_PREFIX + "PENDING_FOLDERS", pending);
        props.setProperty(PROP_PREFIX + "CURRENT_FOLDER_ID", currentFolderId);
        props.deleteProperty(PROP_PREFIX + "CURRENT_FILE_TOKEN");
        props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_TOKEN");
        props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_QUEUE");
      }

      const doneWithFolder = processOneFolderChunk_(currentFolderId, songFolders, doApply, started);
      if (doneWithFolder) {
        props.deleteProperty(PROP_PREFIX + "CURRENT_FOLDER_ID");
        props.deleteProperty(PROP_PREFIX + "CURRENT_FILE_TOKEN");
        props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_TOKEN");
        props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_QUEUE");
      }
    }

    scheduleResume_();
    Logger.log("Paused safely. Will resume automatically.");
  } finally {
    lock.releaseLock();
  }
}

function processOneFolderChunk_(folderId, songFolders, doApply, startedMs) {
  const props = PropertiesService.getScriptProperties();
  const folder = DriveApp.getFolderById(folderId);

  let files;
  const fileToken = props.getProperty(PROP_PREFIX + "CURRENT_FILE_TOKEN");
  if (fileToken) {
    files = DriveApp.continueFileIterator(fileToken);
  } else {
    files = folder.getFiles();
  }

  while (files.hasNext()) {
    const file = files.next();
    processFile_(file, songFolders, doApply);

    if (isTimeUp_(startedMs)) {
      if (files.hasNext()) {
        props.setProperty(PROP_PREFIX + "CURRENT_FILE_TOKEN", files.getContinuationToken());
      } else {
        props.deleteProperty(PROP_PREFIX + "CURRENT_FILE_TOKEN");
      }
      return false;
    }
  }
  props.deleteProperty(PROP_PREFIX + "CURRENT_FILE_TOKEN");

  let subfolders;
  const subfolderToken = props.getProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_TOKEN");
  if (subfolderToken) {
    subfolders = DriveApp.continueFolderIterator(subfolderToken);
  } else {
    subfolders = folder.getFolders();
  }

  let queue = getJsonProp_(PROP_PREFIX + "CURRENT_SUBFOLDER_QUEUE", []);

  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    const name = sub.getName();

    if (!IGNORE_FOLDER_NAMES[name]) {
      queue.push(sub.getId());
    }

    if (isTimeUp_(startedMs)) {
      setJsonProp_(PROP_PREFIX + "CURRENT_SUBFOLDER_QUEUE", queue);
      if (subfolders.hasNext()) {
        props.setProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_TOKEN", subfolders.getContinuationToken());
      } else {
        props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_TOKEN");
      }
      return false;
    }
  }

  props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_TOKEN");

  const pending = getJsonProp_(PROP_PREFIX + "PENDING_FOLDERS", []);
  setJsonProp_(PROP_PREFIX + "PENDING_FOLDERS", queue.concat(pending));
  props.deleteProperty(PROP_PREFIX + "CURRENT_SUBFOLDER_QUEUE");

  return true;
}

function processFile_(file, songFolders, doApply) {
  const fileName = file.getName();
  if (!fileName) return;

  const toks = tokenize_(fileName);
  if (!toks.length) return;

  const best = findBestSongFolder_(toks, songFolders);
  if (!best) return;

  const targetFolder = DriveApp.getFolderById(best.id);

  let alreadyThere = false;
  const parents = file.getParents();
  while (parents.hasNext()) {
    const p = parents.next();
    if (p.getId() === targetFolder.getId()) {
      alreadyThere = true;
      break;
    }
  }
  if (alreadyThere) return;

  incrementActionCount_();

  Logger.log(JSON.stringify({
    action: doApply ? "MOVE" : "DRY_RUN_MATCH",
    fileName: fileName,
    fileId: file.getId(),
    targetFolderName: best.name,
    targetFolderId: best.id
  }));

  if (doApply) {
    file.moveTo(targetFolder);
  }
}

function getSongFolders_(rootFolder) {
  const songFolders = [];
  const stageFolders = rootFolder.getFolders();

  while (stageFolders.hasNext()) {
    const stage = stageFolders.next();
    if (IGNORE_FOLDER_NAMES[stage.getName()]) continue;

    const children = stage.getFolders();
    while (children.hasNext()) {
      const sf = children.next();
      const nm = sf.getName();

      if (IGNORE_FOLDER_NAMES[nm]) continue;
      if (/^_/.test(nm)) continue;

      const toks = tokenize_(nm);
      if (!toks.length) continue;

      const set = {};
      toks.forEach(function(t) {
        set[t] = true;
      });

      songFolders.push({
        id: sf.getId(),
        name: nm,
        tokenSet: set,
        tokenCount: toks.length
      });
    }
  }

  return songFolders;
}

function findBestSongFolder_(fileTokens, songFolders) {
  let best = null;
  let bestScore = 0;

  for (let i = 0; i < songFolders.length; i++) {
    const sf = songFolders[i];
    let score = 0;

    for (let j = 0; j < fileTokens.length; j++) {
      if (sf.tokenSet[fileTokens[j]]) score++;
    }

    if (score > bestScore) {
      bestScore = score;
      best = sf;
    } else if (score === bestScore && best && sf.tokenCount < best.tokenCount) {
      best = sf;
    }
  }

  return bestScore > 0 ? best : null;
}

function tokenize_(text) {
  return String(text)
    .toLowerCase()
    .replace(/.[^.]+$/, "")
    .replace(/[(){}[],.!?_ -]+/g, " ")
    .replace(/s+/g, " ")
    .trim()
    .split(" ")
    .filter(Boolean)
    .filter(function(t) {
      return t.length > 1;
    });
}

function ensureFolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

function pickKeeper_(files) {
  files.sort(function(a, b) {
    return b.getLastUpdated().getTime() - a.getLastUpdated().getTime();
  });
  return files[0];
}

function resetDedupeState_() {
  clearResumeTrigger_();
  clearWorkingState_();
}

function clearWorkingState_() {
  const props = PropertiesService.getScriptProperties();
  [
    "MODE",
    "ROOT_ID",
    "STARTED_AT",
    "PENDING_FOLDERS",
    "CURRENT_FOLDER_ID",
    "CURRENT_FILE_TOKEN",
    "CURRENT_SUBFOLDER_TOKEN",
    "CURRENT_SUBFOLDER_QUEUE",
    "ACTIONS_COUNT"
  ].forEach(function(k) {
    props.deleteProperty(PROP_PREFIX + k);
  });
}

function stopDedupe() {
  resetDedupeState_();
  Logger.log("Dedupe job stopped and state cleared.");
}

function scheduleResume_() {
  clearResumeTrigger_();
  ScriptApp.newTrigger("dedupeResume_")
    .timeBased()
    .after(RESUME_DELAY_MS)
    .create();
}

function clearResumeTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "dedupeResume_") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function isTimeUp_(startedMs) {
  return Date.now() - startedMs > SAFE_RUNTIME_MS;
}

function incrementActionCount_() {
  const props = PropertiesService.getScriptProperties();
  const n = Number(props.getProperty(PROP_PREFIX + "ACTIONS_COUNT") || "0") + 1;
  props.setProperty(PROP_PREFIX + "ACTIONS_COUNT", String(n));
}

function getJsonProp_(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v ? JSON.parse(v) : fallback;
}

function setJsonProp_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(value));
}