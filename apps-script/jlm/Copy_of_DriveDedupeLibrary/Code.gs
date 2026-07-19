// --- Safety knob #1: only process these mime prefixes ---
const INCLUDE_MIME_PREFIXES = ["audio/", "video/", "image/"]; // add "application/pdf" if desired

function mimeAllowed_(mimeType) {
  if (!mimeType) return false;
  // exact match option
  if (INCLUDE_MIME_PREFIXES.includes(mimeType)) return true;
  // prefix match option
  return INCLUDE_MIME_PREFIXES.some(p => p.endsWith("/") ? mimeType.startsWith(p) : mimeType === p);
}

function buildSongFolderIdSet_(songFolders) {
  const set = {};
  songFolders.forEach(sf => set[sf.id] = true);
  return set;
}

/**
 * @param {string} rootFolderId
 * @param {boolean} apply false=dry run, true=execute moves
 * @return {Array<Object>} actions log
 */
function dedupeToSongFolders(rootFolderId, apply) {
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
    if (!mimeAllowed_(mime)) return; // Safety knob #1

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

    // Safety knob #2: if keeper is already inside ANY recognized song folder, do not move it.
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
      const keeperTargetSongFolderId = bestSongFolderId_(songFolders, keeper.getName());
      if (keeperTargetSongFolderId) {
        actions.push({
          kind: "KEEPER_TO_SONG",
          file: keeper.getName(),
          keeperId: keeper.getId(),
          targetSongFolderId: keeperTargetSongFolderId
        });
        if (apply) {
          moveFileToFolder_(keeper, DriveApp.getFolderById(keeperTargetSongFolderId));
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

    // Always quarantine duplicates (never delete)
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