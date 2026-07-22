/**
 * Calls the external serverless function to hash a large WAV file.
 */
function getCloudHash(fileId) {
  // Replace with your deployed Cloud Function or Cloudflare Worker URL
  const CLOUD_FUNCTION_URL = 'https://REGION-PROJECT_ID.cloudfunctions.net/hashDriveFile'; 
  const token = ScriptApp.getOAuthToken(); 
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      fileId: fileId,
      accessToken: token
    }),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(CLOUD_FUNCTION_URL, options);
  
  if (response.getResponseCode() === 200) {
    return response.getContentText();
  } else {
    Logger.log('Hashing Service Error: ' + response.getContentText());
    return 'Error';
  }
}

/**
 * Generates a manifest.json file from sheet data and saves it to the Drive folder.
 */
function generateMetadataManifest(sheet, row, folderUrl) {
  // Extract the folder ID from the URL to save the JSON in the right place
  const folderId = folderUrl.match(/[-\w]{25,}/);
  if (!folderId) return;

  const folder = DriveApp.getFolderById(folderId[0]);
  
  // Pull data from your specific columns
  const metadata = {
    songTitle: sheet.getRange(row, 1).getValue(),
    releaseName: sheet.getRange(row, 2).getValue(),
    workflowStatus: sheet.getRange(row, 4).getValue(),
    isrc: sheet.getRange(row, 13).getValue() || "Pending",
    upc: sheet.getRange(row, 14).getValue() || "Pending",
    timestamp: new Date().toISOString()
  };

  const jsonString = JSON.stringify(metadata, null, 2);
  
  // Check if manifest already exists to overwrite, otherwise create new
  const existingFiles = folder.getFilesByName('manifest.json');
  if (existingFiles.hasNext()) {
    existingFiles.next().setContent(jsonString);
  } else {
    folder.createFile('manifest.json', jsonString, MimeType.PLAIN_TEXT);
  }
  
  Logger.log(`Manifest created for ${metadata.songTitle}`);
}
