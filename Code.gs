// ===== BACKEND - Google Apps Script =====

// Hàm doGet - Entry point
function doGet() {
  var activeSpreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var template = HtmlService.createTemplateFromFile('Index');
  template.sheetUrl = activeSpreadsheetUrl;
  return template
    .evaluate()
    .setTitle('Ghi chú (gsheets.vn)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl("https://cdn-icons-png.flaticon.com/512/8262/8262990.png");
}

// Include HTML files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== API ENDPOINTS =====

// Get all data from all sheets using Sheets API V4
function getAllData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var spreadsheetId = ss.getId();
    
    // Use Sheets API V4 to read data
    // Now Prompts/Notes include a `tags` column after `order` (A..I)
    // Accounts: Cột I: isDeleted, Tasks: Cột J: isDeleted
    var ranges = ['Prompts!A2:I', 'Notes!A2:I', 'Accounts!A2:J', 'Tasks!A2:I'];
    var result = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, {ranges: ranges});
    
    var data = {
      prompts: parseSheetData(result.valueRanges[0].values || []),
      notes: parseSheetData(result.valueRanges[1].values || []),
      accounts: parseAccountData(result.valueRanges[2].values || []),
      tasks: parseTaskData(result.valueRanges[3].values || [])
    };
    
    return data;
  } catch (e) {
    Logger.log('Error in getAllData: ' + e.toString());
    return {error: e.toString()};
  }
}

// Parse sheet data helper
function parseSheetData(rows) {
  if (!rows) return [];
  
  return rows
    .filter(function(row) { 
      return row[0] && row[0].toString().trim(); 
    })
    .map(function(row) {
      return {
        id: row[0] || '',
        parentId: row[1] || '',
        title: row[2] || '',
        content: row[3] || '',
        level: parseInt(row[4]) || 0,
        order: parseInt(row[5]) || 0,
        tags: normalizeTags(row[6] || ''), // new tags column (G) - normalized
        isPinned: row[7] === 'TRUE' || row[7] === true || false, // Cột H: isPinned
        isDeleted: row[8] === 'TRUE' || row[8] === true || false // Cột I: isDeleted
      };
    });
}

// Normalize tags string or array: lowercase, trim, split/join by comma, unique
function normalizeTags(tags) {
  try {
    if (!tags && tags !== 0) return '';
    var arr = [];
    if (Object.prototype.toString.call(tags) === '[object Array]') {
      arr = tags.slice();
    } else {
      arr = tags.toString().split(',');
    }

    var seen = {};
    var out = [];
    for (var i = 0; i < arr.length; i++) {
      var t = (arr[i] || '').toString().toLowerCase().trim();
      if (!t) continue;
      // Remove leading # if provided
      if (t.charAt(0) === '#') t = t.substring(1);
      if (!t) continue;
      if (!seen[t]) {
        seen[t] = true;
        out.push(t);
      }
    }
    return out.join(',');
  } catch (e) {
    return '';
  }
}

// Parse account data helper
function parseAccountData(rows) {
  if (!rows) return []; 
  
  return rows
    .filter(function(row) { 
      return row[0] && row[0].toString().trim();
    })
    .map(function(row) {
      return {
        id: row[0] || '',
        parentId: row[1] || '',
        accountType: row[2] || '',
        username: row[3] || '',
        password: row[4] || '',
        phone: row[5] || '',
        notes: row[6] || '',
        level: parseInt(row[7]) || 0,
        order: parseInt(row[8]) || 0
      };
    });
}


function parseTaskData(rows) {
  if (!rows) return [];
  
  return rows
    .filter(function(row) { 
      return row[0] && row[0].toString().trim();
    })
    .map(function(row) {
      return {
        id: row[0] || '',
        parentId: row[1] || '',
        date: row[2] || '',
        title: row[3] || '',
        content: row[4] || '',
        level: parseInt(row[5]) || 0,
        completed: row[6] === 'TRUE' || row[6] === true,
        order: parseInt(row[7]) || 0,
        isDeleted: row[8] === 'TRUE' || row[8] === true || false // Cột I: isDeleted
      };
    });
}


// Create item in sheet
function createItem(sheetName, item) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return {success: false, error: 'Không tìm thấy sheet'};
    }

    var lastRow = sheet.getLastRow();
    var targetRow = lastRow + 1;
    
    if (lastRow < 2) {
      targetRow = 2;
    }

    var row;
    
    if (sheetName === 'Accounts') {
      var phoneValue = item.phone ? "'" + item.phone : '';
      
      row = [
        item.id,
        item.parentId || '',
        item.accountType || '',
        item.username || '',
        item.password || '',
        phoneValue,
        item.notes || '',
        item.level || 0,
        item.order || 0,
        item.isDeleted || false // Cột J: isDeleted
      ];
    } else if (sheetName === 'Tasks') {
      row = [
        item.id,
        item.parentId || '',
        item.date || '',
        item.title || '',
        item.content || '',
        item.level || 0,
        item.completed || false,
        item.order || 0,
        item.isDeleted || false // Cột I: isDeleted
      ];
    } else {
      row = [
        item.id,
        item.parentId || '',
        item.title || '',
        item.content || '',
        item.level || 0,
        item.order || 0,
        normalizeTags(item.tags || ''), // Cột G: tags (mới) - normalized
        item.isPinned || false, // Cột H: isPinned
        item.isDeleted || false // Cột I: isDeleted
      ];
    }
    
    if (targetRow > sheet.getMaxRows()) {
      sheet.insertRowAfter(lastRow);
    }
    
    var numCols = row.length;
    sheet.getRange(targetRow, 1, 1, numCols).setValues([row]);
    
    return {success: true, item: item};
  } catch (e) {
    Logger.log('Error in createItem: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

// Update item in sheet using Sheets API V4
function updateItem(sheetName, itemId, updates) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return {success: false, error: 'Không tìm thấy sheet'};
    }
    
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;
    
    // Find row by ID
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === itemId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return {success: false, error: 'Không tìm thấy mục'};
    }
    
    // Update using setValues
    if (sheetName === 'Accounts') {
      var phoneValue = updates.phone !== undefined ? 
        (updates.phone ? "'" + updates.phone : '') : 
        data[rowIndex-1][5];
      
      var currentCols = data[rowIndex-1].length;
      var isDeletedValue = updates.isDeleted !== undefined ? updates.isDeleted : 
                          (currentCols > 9 ? (data[rowIndex-1][9] === 'TRUE' || data[rowIndex-1][9] === true) : false);
      
      var rowData = [
        itemId,
        updates.parentId !== undefined ? updates.parentId : data[rowIndex-1][1],
        updates.accountType !== undefined ? updates.accountType : data[rowIndex-1][2],
        updates.username !== undefined ? updates.username : data[rowIndex-1][3],
        updates.password !== undefined ? updates.password : data[rowIndex-1][4],
        phoneValue,
        updates.notes !== undefined ? updates.notes : data[rowIndex-1][6],
        updates.level !== undefined ? updates.level : data[rowIndex-1][7],
        updates.order !== undefined ? updates.order : data[rowIndex-1][8],
        isDeletedValue // Cột J: isDeleted
      ];
      
      sheet.getRange(rowIndex, 1, 1, 10).setValues([rowData]);
    } else if (sheetName === 'Tasks') {
      var currentCols = data[rowIndex-1].length;
      var isDeletedValue = updates.isDeleted !== undefined ? updates.isDeleted : 
                          (currentCols > 8 ? (data[rowIndex-1][8] === 'TRUE' || data[rowIndex-1][8] === true) : false);
      
      var rowData = [
        itemId,
        updates.parentId !== undefined ? updates.parentId : data[rowIndex-1][1],
        updates.date !== undefined ? updates.date : data[rowIndex-1][2],
        updates.title !== undefined ? updates.title : data[rowIndex-1][3],
        updates.content !== undefined ? updates.content : data[rowIndex-1][4],
        updates.level !== undefined ? updates.level : data[rowIndex-1][5],
        updates.completed !== undefined ? updates.completed : data[rowIndex-1][6],
        updates.order !== undefined ? updates.order : data[rowIndex-1][7],
        isDeletedValue // Cột I: isDeleted
      ];
      
      sheet.getRange(rowIndex, 1, 1, 9).setValues([rowData]);
    } else {
      // Kiểm tra số cột hiện tại
      var currentCols = data[rowIndex-1].length;
      var tagsValue = updates.tags !== undefined ? normalizeTags(updates.tags) : (currentCols > 6 ? normalizeTags(data[rowIndex-1][6]) : '');
      var isPinnedValue = updates.isPinned !== undefined ? updates.isPinned : 
                         (currentCols > 7 ? (data[rowIndex-1][7] === 'TRUE' || data[rowIndex-1][7] === true) : false);
      var isDeletedValue = updates.isDeleted !== undefined ? updates.isDeleted : 
                          (currentCols > 8 ? (data[rowIndex-1][8] === 'TRUE' || data[rowIndex-1][8] === true) : false);

      var rowData = [
        itemId,
        updates.parentId !== undefined ? updates.parentId : data[rowIndex-1][1],
        updates.title !== undefined ? updates.title : data[rowIndex-1][2],
        updates.content !== undefined ? updates.content : data[rowIndex-1][3],
        updates.level !== undefined ? updates.level : data[rowIndex-1][4],
        updates.order !== undefined ? updates.order : data[rowIndex-1][5],
        tagsValue, // Cột G: tags
        isPinnedValue, // Cột H: isPinned
        isDeletedValue // Cột I: isDeleted
      ];

      sheet.getRange(rowIndex, 1, 1, 9).setValues([rowData]);
    }
    
    return {success: true};
  } catch (e) {
    Logger.log('Error in updateItem: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

// Soft Delete: Set isDeleted = TRUE instead of deleting rows
function deleteItem(sheetName, itemId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return {success: false, error: 'Không tìm thấy sheet'};
    }
    
    var data = sheet.getDataRange().getValues();
    var idsToDelete = [itemId];
    
    // Find all children and descendants
    var toProcess = [itemId];
    while (toProcess.length > 0) {
      var currentId = toProcess.shift();
      for (var i = 1; i < data.length; i++) {
        if (data[i][1] === currentId) { // parentId column
          var childId = data[i][0];
          idsToDelete.push(childId);
          toProcess.push(childId);
        }
      }
    }
    
    // Update isDeleted = TRUE for all items and their children
    // Determine column index for isDeleted based on sheet type
    var isDeletedColIndex;
    if (sheetName === 'Accounts') {
      isDeletedColIndex = 10; // Column J (0-indexed: 9)
    } else if (sheetName === 'Tasks') {
      isDeletedColIndex = 9; // Column I (0-indexed: 8)
    } else {
      isDeletedColIndex = 9; // Column I for Prompts/Notes (tags added before)
    }
    
    // Update each item's isDeleted status
    for (var i = 1; i < data.length; i++) {
      if (idsToDelete.indexOf(data[i][0]) !== -1) {
        sheet.getRange(i + 1, isDeletedColIndex).setValue('TRUE');
      }
    }
    
    return {success: true};
  } catch (e) {
    Logger.log('Error in deleteItem: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

// Batch update order for multiple items
function batchUpdateOrder(sheetName, updates) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return {success: false, error: 'Không tìm thấy sheet'};
    }
    
    var data = sheet.getDataRange().getValues();
    
    // Update each item's order
    updates.forEach(function(update) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === update.id) {
          var orderColIndex = sheetName === 'Accounts' ? 8 : 
                            (sheetName === 'Tasks' ? 7 : 5);
          sheet.getRange(i + 1, orderColIndex + 1).setValue(update.order);
          break;
        }
      }
    });
    
    return {success: true};
  } catch (e) {
    Logger.log('Error in batchUpdateOrder: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

// Hard delete an item by ID from the specified sheet (do not remove header)
function deleteItemById(sheetName, itemId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return {success: false, error: 'Sheet not found'};
    }

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { // start at row index 1 => sheet row 2
      if (data[i][0] === itemId) {
        var rowNumber = i + 1;
        sheet.deleteRow(rowNumber);
        return {success: true};
      }
    }

    return {success: false, error: 'ID not found'};
  } catch (e) {
    Logger.log('Error in deleteItemById: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}

// ===== SETTINGS FUNCTIONS =====

// Get all settings from Settings sheet
function getSettings() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Settings');
    
    if (!sheet) {
      // Create Settings sheet if it doesn't exist
      sheet = ss.insertSheet('Settings');
      sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
      return {};
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return {};
    }
    
    var settings = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        settings[data[i][0]] = data[i][1];
      }
    }
    
    return settings;
  } catch (e) {
    Logger.log('Error in getSettings: ' + e.toString());
    return {};
  }
}

// Save settings to Settings sheet
function saveSettings(settings) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Settings');
    
    if (!sheet) {
      sheet = ss.insertSheet('Settings');
      sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    }
    
    // Clear existing data (except header)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    
    // Write new settings
    var rows = [['key', 'value']];
    for (var key in settings) {
      if (settings.hasOwnProperty(key)) {
        rows.push([key, settings[key]]);
      }
    }
    
    if (rows.length > 1) {
      sheet.getRange(1, 1, rows.length, 2).setValues(rows);
    }
    
    return {success: true};
  } catch (e) {
    Logger.log('Error in saveSettings: ' + e.toString());
    return {success: false, error: e.toString()};
  }
}