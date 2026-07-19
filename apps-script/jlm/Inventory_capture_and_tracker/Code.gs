function doGet() {
  return HtmlService.createHtmlOutputFromFile('MobileIndex')
    .setTitle('Inventory Scanner');
}

function processInventory(data) {
  try {
    var sheetId = '1hU4FEShRz2BnKiD1trOBfiVkR-Aj1a1OgthHAts4JCo';
    var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
    var row = sheet.getLastRow() + 1;
    
    var imageUrl = 'No image';
    var label = data.label || 'Unknown';
    var estimatedValue = '$0.00';
    
    if (data.image && data.image !== 'no-image') {
      imageUrl = uploadImageToDrive(data.image);
    }
    
    if (data.image && data.image !== 'no-image') {
      var aiLabel = detectObjectLabel(data.image);
      if (aiLabel !== 'API_NOT_CONFIGURED') {
        label = aiLabel;
      }
    }
    
    estimatedValue = estimateValue(label);
    var listPrice = data.price || estimatedValue.replace('$', '');
    
    sheet.getRange(row, 1, 1, 7).setValues([[
      imageUrl,
      label,
      estimatedValue,
      '$' + listPrice,
      new Date(),
      data.notes || '',
      'eBay pending approval'
    ]]);
    
    if (row === 2) {
      sheet.getRange(1, 1, 1, 7).setValues([[
        'Image', 'Item', 'Est. Value', 'List Price', 'Date', 'Notes', 'eBay Status'
      ]]);
    }
    
    return { 
      success: true, 
      row: row, 
      label: label,
      value: estimatedValue,
      ebay: 'Ready when eBay approves'
    };
    
  } catch (error) {
    return { 
      success: false, 
      error: error.toString() 
    };
  }
}

function uploadImageToDrive(imageData) {
  try {
    var FOLDER_ID = '1t-VW7XWsZWMJFjLkY4GfkUyir1RmIae6';
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var blob = Utilities.newBlob(
      Utilities.base64Decode(imageData.split(',')[1]), 
      'image/jpeg', 
      'inventory_' + Date.now() + '.jpg'
    );
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return 'Upload failed';
  }
}

function detectObjectLabel(imageData) {
  try {
    // NOTE: API key redacted from source backup. Store in Script Properties
    // as VISION_API_KEY and load via PropertiesService — never hardcode.
    var API_KEY = PropertiesService.getScriptProperties().getProperty('VISION_API_KEY') || 'YOUR_VISION_API_KEY';
    
    if (API_KEY === 'YOUR_VISION_API_KEY') {
      return 'API_NOT_CONFIGURED';
    }
    
    var url = 'https://vision.googleapis.com/v1/images:annotate?key=' + API_KEY;
    var imageBytes = imageData.split(',')[1];
    
    var payload = {
      requests: [{
        image: { content: imageBytes },
        features: [
          { type: 'LABEL_DETECTION', maxResults: 1 },
          { type: 'OBJECT_LOCALIZATION', maxResults: 1 }
        ]
      }]
    };
    
    var response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    var json = JSON.parse(response.getContentText());
    
    if (json.responses[0].localizedObjectAnnotations && 
        json.responses[0].localizedObjectAnnotations.length > 0) {
      return json.responses[0].localizedObjectAnnotations[0].name;
    }
    if (json.responses[0].labelAnnotations && 
        json.responses[0].labelAnnotations.length > 0) {
      return json.responses[0].labelAnnotations[0].description;
    }
    
    return 'Unknown item';
    
  } catch (e) {
    return 'AI detection failed';
  }
}

function estimateValue(itemName) {
  try {
    var APP_ID = 'YOUR_EBAY_APP_ID';
    
    if (APP_ID === 'YOUR_EBAY_APP_ID') {
      return '$0.00';
    }
    
    var url = 'https://svcs.ebay.com/services/search/FindingService/v1?' +
      'OPERATION-NAME=findCompletedItems&' +
      'SERVICE-VERSION=1.0.0&' +
      'SECURITY-APPNAME=' + APP_ID + '&' +
      'RESPONSE-DATA-FORMAT=JSON&' +
      'REST-PAYLOAD&' +
      'keywords=' + encodeURIComponent(itemName + ' used') + '&' +
      'itemFilter(0).name=SoldItemsOnly&' +
      'itemFilter(0).value=true&' +
      'sortOrder=PricePlusShippingLowest&' +
      'paginationInput.entriesPerPage=5';
    
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(response.getContentText());
    
    var items = json.findCompletedItemsResponse[0].searchResult[0].item;
    if (items && items.length > 0) {
      var total = 0;
      var count = 0;
      for (var i = 0; i < items.length; i++) {
        var price = parseFloat(items[i].sellingStatus[0].currentPrice[0].__value__);
        total += price;
        count++;
      }
      var avg = (total / count).toFixed(2);
      return '$' + avg;
    }
    
    return '$0.00';
    
  } catch (e) {
    return '$0.00';
  }
}