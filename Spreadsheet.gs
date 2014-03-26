/*********************************************
 * Functions to use with Google spreadsheet
 *
 * @Author: Nicolas Trinh
 *********************************************/

/**
 * Adds a custom menu to the active spreadsheet
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened. 
 */
function onOpen() {
  var _sheet = SpreadsheetApp.getActiveSpreadsheet();
  var _entries = [
    {
      name : "Initialize Parameters",
      functionName : "initParamSheet"
    },
    {
      name : "Process Data",
      functionName : "main"
    }
  ];
  _sheet.addMenu("Zipabox", _entries);
  
  Logger.clear();   
}


/**
 * Get a column index by its name
 * For use with Google Spreadsheets
 */
function getColIndexByName(colName, tabName) {
  if(!tabName) {
    tabName = PropertiesService.getScriptProperties().getProperty("paramSheet");
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  var numColumns = sheet.getLastColumn();
  var row = sheet.getRange(1, 1, 1, numColumns).getValues();
  for (i in row[0]) {
    var name = row[0][i];
    if (name == colName) {
      return parseInt(i) + 1;
    }
  }
  return -1;
}

/**
 * Get a row index by its name
 * For use with Google Spreadsheets
 */
function getRowIndexByName(rowName, tabName) {
  if(!tabName) {
    tabName = PropertiesService.getScriptProperties().getProperty("paramSheet");
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  var numRows = sheet.getLastRow();
  
  var row = sheet.getRange(1, 1, numRows, 1).getValues();
  for (i=0; i<row.length; i++) {
    var name = row[i][0];
    if (name == rowName) {
      return parseInt(i) + 1;
    }
  }
  return -1;
}

/**
 * Get the row range by its UUID
 * For use with Google Spreadsheets
 */
function getRowByUUID(uuid, tabName) {
  if(!tabName) {
    tabName = PropertiesService.getScriptProperties().getProperty("paramSheet");
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  var numRows = sheet.getLastRow();
  
  var row = sheet.getRange(1, 1, numRows, 5).getValues();
  
  for (i=0; i<row.length; i++) {
    var name = row[i][1];
    if (name == uuid) {
      return sheet.getRange(i+1, 1, 1, 5);
    }
  }
  
  return -1;  
}

/**
 * Get properties range by its section name
 */
function getPropertiesRangeByName(sectionName, tabName) {
  if(!tabName) {
    tabName = PropertiesService.getScriptProperties().getProperty("paramSheet");
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
  var rowPos = getRowIndexByName(sectionName, tabName);
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  // If the section is not found return false
  if(rowPos == -1) return false;
  
  // Return the section range    
  var search = sheet.getRange(rowPos, 1, lastRow-rowPos+1, 4).getValues();
  var find = 0;
  for (i=0; i<search.length; i++) {
    if (search[i][0] != '') {
      find ++;
    } else {
      break;
    }
  }
    
  return sheet.getRange(rowPos, 1, find, lastColumn);
}


/**
 * Insert a new record in the sheet
 * @param name: name of the sheet
 * @param value: value of the record
 */
function _insertRecord(name, value) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(name);
  var lastRow = sheet.getLastRow();
  var activeRange = sheet.getRange(lastRow+1, 1, lastRow+2, 2);
  var dateFormat = (typeof zipabox.dateFormat != "undefined") ? zipabox.dateFormat : "dd/MM/yyyy HH:mm:ss";
  
  activeRange.getCell(1, 1).setValue(new Date());
  activeRange.getCell(1, 1).setNumberFormat(dateFormat);
  activeRange.getCell(1, 2).setValue(value);
}


/**
 * Get the feedId from the params sheet
 * @param typeDevice : type of the device (meters, sensors, lights...)
 * @param deviceName : the device name, not used at this time
 * @param deviceId : the deviceId from the zipabox
 * @param attributeName : name of the attribute
 * @return the feedID matching the device, return 0 if not found
 */
function getFeedID(typeDevice, deviceName, deviceID, attributeName) {
  writelog("==> getFeedID ***");
  writelog("checking for typeDevice["+typeDevice+"] / deviceName["+deviceName+"] / deviceID["+deviceID+"]");
  
  var feedID = 0;
  var paramName = CacheService.getPrivateCache().get("paramSheet");
  var range = getPropertiesRangeByName(typeDevice, paramName);
  
  // If no device is found return 0
  if(!range) {
    writelog("No feedID found");
    return 0;
  }
  
  var listDevices = range.getValues();
  
  for(var i=1; i<listDevices.length; i++) {    
    // get the sense feedID of the device
    if (listDevices[i][1] == deviceID && listDevices[i][2] == attributeName) {
      feedID = listDevices[i][3];
      writelog("FeedID found = "+feedID+" for device ["+deviceName+"]");
      
      // check if device should be send to sense
      if(!listDevices[i][4]) {
        writelog("Device set to FALSE => Not sending");
        feedID = 0;
      }
      
      break;
    }
  }
  
  if(parseInt(feedID) == 0) writelog("No feedID found");
  
  return parseInt(feedID);
}

/**
 * Create a sheet name
 * @param name: name of the device
 * @param type: type of the device
 */
function createSheetName(name, type) {
  var sheetName = name+" - "+type;
  
  if(sheetName.length > 50) {
    var sup = sheetName.length-50;
    var short = name.length-sup-3;
    sheetName = name.substr(0, short)+"... - "+type;
  }
  
  return sheetName;  
}
