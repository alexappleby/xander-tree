/**
 * Drive Dedupe Library
 * - Dedupe binary assets by md5Checksum
 * - Keeper = newest modified
 * - Safety #1: only audio/video/image (configurable)
 * - Safety #2: if keeper already in any song folder, do NOT move it across stages
 * - Duplicates are moved to _Quarantine/_Duplicates (never deleted)
 *
 * REQUIREMENTS:
 * - Apps Script: Services -> Advanced Google services -> Drive API ON
 * - Google Cloud: Enable Google Drive API
 */

const IGNORE_FOLDER_NAMES = {
  "_Quarantine": true,
  "_Duplicates": true,
  "Duplicate tracks": true
};

// Safety knob #1: only process these mime prefixes (add "application/pdf" if you want)
const INCLUDE_MIME_PREFIXES = ["audio/", "video/", "image/"];

function mimeAllowed_(mimeType) {
  if (!mimeType) return false;
  if (INCLUDE_MIME_PREFIXES.includes(mimeType)) return true;
  return INCLUDE_MIME_PREFIXES.some(p => p.endsWith("/") ? mimeType.startsWith(p) : mimeType === p);
}

function normalizeTitle_(s) {
  return (s || "")
    .toLowerCase()
    .replace(/\.[^/.]+$/, "")                 // drop extension
    .replace(/[_\-]+/g, " ")
    .replace(/[’']/g, "")                     // normalize apostrophes
    .replace(/\b(copy of|final|master|mix|mv|v\d+)\b/g, "")
    .replace(/\b(20\d{2}[\s\-_]?\d{2}[\s\-_]?\d{2})\b/g, "") // strip date blobs like 2026-05-28
    .replace(/\s+/g, " ")
    .trim();
}

function tokenize_(s) {
  return normalizeTitle_(s)
    .split(" ")
    .filter(t => t && t.length >= 3); // drop tiny tokens like "a", "to"
}

function ensureFolder_(parentFolder, name) {
  const it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

function listAllFilesRecursive_(folder) {
  let out = [];
  const folders = folder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (IGNORE_FOLDER_NAMES[f.getName()]) continue;
    out = out.concat(listAllFilesRecursive_(f));
  }
  const files = folder.getFiles();
  while (files.hasNext()) out.push(files.next());
  return out;
}

/**
 * Stage-first assumption:
 * root/
 *   01-Pre-Production/
 *     Brown eyes/
 *   02-In Progress/
 *     Brown eyes/
 *   03-Released/
 *     Brown eyes/
 */
function getSongFolders_(rootFolder) {
  const songFolders = []; // {id, name, tokenSet, tokenCount}
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
      if (toks.length === 0) continue;

      const set = {};
      toks.forEach(t => set[t] = true);

      songFolders.push({ id: sf.getId(), name: nm, tokenSet: set, tokenCount: toks.length });
    }
  }
  return songFolders;
}

function buildSongFolderIdSet_(songFolders) {
  const set = {};
  songFolders.forEach(sf => set[sf.id] = true);
  return set;
}

function scoreSongFolder_(songFolder, fileName) {
  const ftoks = tokenize_(fileName);
  if (ftoks.length === 0) return 0;

  let shared = 0;
  ftoks.forEach(t => { if (songFolder.tokenSet[t]) shared++; });

  // overlap ratio
  return shared / Math.max(1, Math.min(ftoks.length, songFolder.tokenCount));
}

function bestSongFolderId_(songFolders, fileName) {
  let best = null;
  let bestScore = 0;

  songFolders.forEach(sf => {
    const score = scoreSongFolder_(sf, fileName);
    if (score > bestScore) {
      bestScore = score;
      best = sf;
    }
  });

  // require meaningful overlap
  if (!best || bestScore < 0.6) return null;
  return best.id;
}

function pickKeeper_(files) {
  files.sort((a, b) => b.getLastUpdated().getTime() - a.getLastUpdated().getTime());
  return files[0];
}

function moveFileToFolder_(file, targetFolder) {
  targetFolder.addFile(file);
  const parents = file.getParents();
  while (parents.hasNext()) {
    parents.next().removeFile(file);
  }
}

/**
 * Main entrypoint
 * @param {string} rootFolderId
 * @param {boolean} apply false=dry run, true=execute moves
 * @return {Array<Object>} actions log
 */
function dedupeToSongFolders(rootFolderId, apply) {
  if (!rootFolderId || typeof rootFolderId !== "string") {
    throw new Error("rootFolderId is missing. Pass a valid folder ID string.");
  }

  apply = !!apply;

  const root = DriveApp.getFolderById(rootFolderId);
  const quarantine = ensureFolder_(root, "_Quarantine");
  const dupesFolder = ensureFolder_(quarantine, "_Duplicates");

  const songFolders = getSongFolders_(root);
  const songFolderIdSet = buildSongFolderIdSet_(songFolders);

  const allFiles = listAllFilesRecursive_(root);

  // Group by md5Checksum
  const byHash = {};
  allFiles.forEach(f => {
    if (f.getMimeType() === MimeType.FOLDER) return;

    const mime = f.getMimeType();
    if (!mimeAllowed_(mime)) return; // Safety #1

    const meta = Drive.Files.get(f.getId()); // Advanced Drive service
    const md5 = meta.md5Checksum;
    if (!md5) return; // skip Google-native files

    if (!byHash[md5]) byHash[md5] = [];
    byHash[md5].push(f);
  });

  const actions = [];

  Object.keys(byHash).forEach(h => {
    const group = byHash[h];
    if (group.length < 2) return;

    const keeper = pickKeeper_(group);

    // Safety #2: if keeper already in any song folder, do not move it
    let keeperAlreadyInSongFolder = false;
    const parents = keeper.getParents();
    while (parents.hasNext()) {
      const p = parents.next();
      if (songFolderIdSet[p.getId()]) {
        keeperAlreadyInSongFolder = true;
        break;
      }
    }

    if (keeperAlreadyInSongFolder) {
      actions.push({
        kind: "KEEPER_LEFT_IN_PLACE",
        file: keeper.getName(),
        keeperId: keeper.getId(),
        note: "Keeper already in a song folder; not moving across stages."
      });
    } else {
      const targetSongFolderId = bestSongFolderId_(songFolders, keeper.getName());
      if (targetSongFolderId) {
        actions.push({
          kind: "KEEPER_TO_SONG",
          file: keeper.getName(),
          keeperId: keeper.getId(),
          targetSongFolderId: targetSongFolderId
        });
        if (apply) {
          moveFileToFolder_(keeper, DriveApp.getFolderById(targetSongFolderId));
        }
      } else {
        actions.push({
          kind: "KEEPER_UNMATCHED",
          file: keeper.getName(),
          keeperId: keeper.getId(),
          note: "No confident song-folder match; keeper left in place."
        });
      }
    }

    // Quarantine duplicates
    group.forEach(f => {
      if (f.getId() === keeper.getId()) return;

      actions.push({
        kind: "DUPE_TO_QUARANTINE",
        file: f.getName(),
        fileId: f.getId(),
        targetFolderId: dupesFolder.getId()
      });

      if (apply) {
        moveFileToFolder_(f, dupesFolder);
      }
    });
  });

  return actions;
}