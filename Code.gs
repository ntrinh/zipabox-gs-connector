/** 
@Author: Nicolas TRINH

Contributions: 
[1] This project was greatly inspired by Zipabox Connector written in Python by Frédéric Ravetier.
http://sourceforge.net/projects/zipabox-connector/
[2] This project was greatly inspired by "Zipabox" API written in javascript for node.js by djoulz22.
https://github.com/djoulz22/zipabox
*/


/**
 * Init Zipabox with data from the Spreadsheets
 * 
 */
function _initLogin() {
  // Getting the parameters sheet    
  var paramName = PropertiesService.getScriptProperties().getProperty("paramSheet");
  
  if(!paramName) {
    paramName = "Paramètres"; //Default value
  }
  
  CacheService.getPrivateCache().put("paramSheet", paramName);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(paramName);
  var range = getPropertiesRangeByName("ZIPABOX");
  var search = range.getValues();    
  
  //Logger.clear();
  
  // Convert plain text password in the spreadsheet to SHA1 digest
  for (i=0; i<search.length; i++) {
    if (search[i][0] == 'password') {
      // If length = 40 then it is already a SHA1 hash
      if(search[i][1].length == 40) break; 
      
      // Modifying password with SHA1 hash
      var passwordSHA1 = getSHA1(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, 
                                                         search[i][1], 
                                                         Utilities.Charset.UTF_8));
      range.getCell(i+1, 2).setValue(passwordSHA1);
      
      break;
    }
  }
  
  // Retrieving data from "ZIPABOX" section in the spreadsheet
  var zipaboxSection = range.getValues();
  
  for(var i=0; i< zipaboxSection.length; i++) {
    switch(zipaboxSection[i][0]) {
      case 'password':
        zipabox.password = zipaboxSection[i][1];
        break;
        
      case 'user':
        zipabox.username = zipaboxSection[i][1];        
        break;
        
      case 'url':
        zipabox.url = zipaboxSection[i][1];        
        break;
        
      case 'dateFormat':
        zipabox.dateFormat = zipaboxSection[i][1];
    }
  }
}


/**
 * Collect values for creating a feed
 * @param deviceType: type of the device (meters, sensors, lights...)
 * @param uuid: uuid of the device (check logs to see the uuid of the device)
 * @param deviceName: name of the device
 * @param attribute: attribute of the device
 */
function CollectValuesForFeeds(deviceType, uuid, deviceName, attribute, forcedValue){
  var value = attribute.value;
  var attributeName = attribute.name;
  var feedID = getFeedID(deviceType, deviceName, uuid, attributeName);  
  
  if (feedID != 0){
    writelog("Collecting device value for feeding...");
    
    if(typeof forcedValue != "undefined")
      value = forcedValue;
    
    if (!isNaN(parseFloat(value)))
      value = parseFloat(value).toFixed(2);
    
    opensense.feeds[feedID] = {
      "feed_id": feedID,
      "value": value,
      "deviceName": deviceName,
      "uuid": uuid,
      "attributeName": attributeName
    };
    writelog("feedID: "+feedID+"\tValue: "+value+"\tUUID: "+uuid);                
  }
}


/**
 * Get the temperature of a device
 * @param attributeValue : attribute in JSON returned by the zipabox
 * @param name : name of the device
 * @param deviceId : uuid of the device (check logs to see the uuid of the device)
 */
function _getTemperature(attributeValue, name, deviceId) {
  Logger.log("==> _getTemperature ***");
  
  var type = "TEMP";
  
  // Apply a filter on TEMPERATURE                
  if (attributeValue['definition']['name'] == "TEMPERATURE" || attributeValue['definition']['name'] == "TEMPERATURE_IN_ROOM") {
    CollectValuesForFeeds("meters", deviceId, name, attributeValue);
    
    // Insert a record in the spreadsheet    
    if(attributeValue.canCreateSheet) {    
      var sheetName = createSheetName(name, type);
      var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
      
      if (!spreadSheet.getSheetByName(sheetName)) {
        spreadSheet.insertSheet(sheetName);
        spreadSheet.getSheetByName(sheetName).appendRow(["Temperature", deviceId]);
        spreadSheet.getSheetByName(sheetName).appendRow(["Timestamp", "Value"]);
        spreadSheet.getSheetByName(sheetName).getRange("A2:B2").setBackground("purple");
        spreadSheet.getSheetByName(sheetName).getRange("A2:B2").setFontColor("white");
      }
    
      var value = parseFloat(attributeValue['value']); 
      writelog("Insert values in spreadsheet for device ["+name+"] / Value: "+value);
      _insertRecord(sheetName, value);
    }
  }
}


/**
 * Get the humidity of a device
 * @param attributeValue : attribute in JSON returned by the zipabox
 * @param name : name of the device
 * @param deviceId : uuid of the device (check logs to see the uuid of the device)
 */
function _getHumidity(attributeValue, name, deviceId) {
  Logger.log("==> _getHumidity ***");
  
  var type = "HUM";
  
  // Apply a filter on HUMIDITY               
  if (attributeValue['definition']['name'] == "HUMIDITY") {
    CollectValuesForFeeds("meters", deviceId, name, attributeValue);        
    
    // Insert a record in the spreadsheet
    if(attributeValue.canCreateSheet) {
      var sheetName = createSheetName(name, type);
      var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();    
      if (!spreadSheet.getSheetByName(sheetName)) {
        spreadSheet.insertSheet(sheetName);
        spreadSheet.getSheetByName(sheetName).appendRow(["Humidity", deviceId]);
        spreadSheet.getSheetByName(sheetName).appendRow(["Timestamp", "Value"]);
        spreadSheet.getSheetByName(sheetName).getRange("A2:B2").setBackground("purple");
        spreadSheet.getSheetByName(sheetName).getRange("A2:B2").setFontColor("white");
      }
      
      var value = parseFloat(attributeValue['value']); 
      writelog("Insert values in spreadsheet for device ["+name+"] / Value: "+value);
      _insertRecord(sheetName, value);
    }
  }
}


/**
 * Get the luminance of a device
 * @param attributeValue : attribute in JSON returned by the zipabox
 * @param name : name of the device
 * @param deviceId : uuid of the device (check logs to see the uuid of the device)
 */
function _getLuminance(attributeValue, name, deviceId) {
  Logger.log("==> _getLuminance ***");
  
  var type = "LUM";
  
  // Apply a filter on LUMINANCE              
  if (attributeValue['definition']['name'] == "LUMINANCE") {
    CollectValuesForFeeds("meters", deviceId, name, attributeValue);
    
    // Insert a record in the spreadsheet
    if(attributeValue.canCreateSheet) {
      var sheetName = createSheetName(name, type);            
      var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();    
      
      if (!spreadSheet.getSheetByName(sheetName)) {
        spreadSheet.insertSheet(sheetName);
        spreadSheet.getSheetByName(sheetName).appendRow(["Luminance", deviceId]);
        spreadSheet.getSheetByName(sheetName).appendRow(["Timestamp", "Value"]);
        spreadSheet.getSheetByName(sheetName).getRange("A2:B2").setBackground("purple");
        spreadSheet.getSheetByName(sheetName).getRange("A2:B2").setFontColor("white");
      }
      
      var value = parseFloat(attributeValue['value']);
      writelog("Insert values in spreadsheet for device ["+name+"] / Value: "+value);
      _insertRecord(sheetName, value);
    }
  }    
}


/**
 * Get the current consumption
 * @param attributeValue : attribute in JSON returned by the zipabox
 * @param name : name of the device
 * @param deviceId : uuid of the device (check logs to see the uuid of the device)
 */
function _getCurrentConsumption(attributeValue, name, deviceId) {
  Logger.log("==> _getCurrentConsumption ***");
  
  var type = "CCONS";
  
  // Apply a filter on CURRENT_CONSUMPTION              
  if (attributeValue['definition']['name'] == "CURRENT_CONSUMPTION") {
    CollectValuesForFeeds("meters", deviceId, name, attributeValue);
    
    // Insert a record in the spreadsheet
    if(attributeValue.canCreateSheet) {
      var sheetName = createSheetName(name, type);            
      var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();    
      
      if (!spreadSheet.getSheetByName(sheetName)) {
        spreadSheet.insertSheet(sheetName);
        spreadSheet.getSheetByName(sheetName).appendRow(["Current Consumption", deviceId]);
        spreadSheet.getSheetByName(sheetName).appendRow(["Timestamp", "Value"]);
        spreadSheet.getSheetByName(sheetName).getRange("A2:B2").setBackground("purple");
        spreadSheet.getSheetByName(sheetName).getRange("A2:B2").setFontColor("white");
      }
      
      var value = parseFloat(attributeValue['value']);
      writelog("Insert values in spreadsheet for device ["+name+"] / Value: "+value);
      _insertRecord(sheetName, value);
    }
  }    
}


/**
 * Get the cumulative consumption
 * @param attributeValue : attribute in JSON returned by the zipabox
 * @param name : name of the device
 * @param deviceId : uuid of the device (check logs to see the uuid of the device)
 */
function _getCumulativeConsumption(attributeValue, name, deviceId) {
  Logger.log("==> _getCumulativeConsumption ***");
  
  var type = "CUMCONS";
  
  // Apply a filter on CUMULATIVE_CONSUMPTION              
  if (attributeValue['definition']['name'] == "CUMULATIVE_CONSUMPTION") {
    CollectValuesForFeeds("meters", deviceId, name, attributeValue);
    
    // Insert a record in the spreadsheet 
    if(attributeValue.canCreateSheet) {
       var sheetName = createSheetName(name, type);
       var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();    
      
       if (!spreadSheet.getSheetByName(sheetName)) {
         spreadSheet.insertSheet(sheetName);
         spreadSheet.getSheetByName(sheetName).appendRow(["Cumulative Consumption", deviceId]);
         spreadSheet.getSheetByName(sheetName).appendRow(["Timestamp", "Value"]);
         spreadSheet.getSheetByName(sheetName).getRange("A2:B2").setBackground("purple");
         spreadSheet.getSheetByName(sheetName).getRange("A2:B2").setFontColor("white");
       }
       
       var value = parseFloat(attributeValue['value']); // French format with 2 decimals
       writelog("Insert values in spreadsheet for device ["+name+"] / Value: "+value);
       _insertRecord(sheetName, value);
     }
  }
}


/**
 * Get the sensor state
 * @param attributeValue : attribute in JSON returned by the zipabox
 * @param name : name of the device
 * @param deviceId : uuid of the device (check logs to see the uuid of the device)
 */
function _getSensorState(attributeValue, name, deviceId) {
  writelog("==> _getSensorState ***");
  
  var type = "SENSOR";
  var value = attributeValue['value'];
  
  // Get the semantic of the sensor value  
  if((typeof attributeValue['definition'] != "undefined") && (typeof attributeValue['definition']['enumValues'] != "undefined")) {
    value = attributeValue['definition']['enumValues'][attributeValue['value']];
  }
  
  CollectValuesForFeeds("sensors", deviceId, name, attributeValue, value);  
  
  // Insert a record in the spreadsheet
  if(attributeValue.canCreateSheet) {
    var sheetName = createSheetName(name, type);
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (!spreadSheet.getSheetByName(sheetName)) {
      spreadSheet.insertSheet(sheetName);
      spreadSheet.getSheetByName(sheetName).appendRow(["Timestamp", "Value"]);
      spreadSheet.getSheetByName(sheetName).getRange("A1:B1").setBackground("purple");
      spreadSheet.getSheetByName(sheetName).getRange("A1:B1").setFontColor("white");
    }
    
    writelog("Insert values in spreadsheet for device ["+name+"] / Value: "+value);
    _insertRecord(sheetName, value); 
  }
 
  writelog("*** _getSensorState <==");
}


/**
 * Collect all data from Zipabox
 * Data are stored in zipabox.devices
 */
function collectDataFromZipabox() {
  zipabox.showlog = true;
  
  /**************************
   * ### CALLBACK EVENTS ###
   **************************/
  
  // Callback event: Init properties before connect
  zipabox.events.OnBeforeConnect = _initLogin;
  
  // Callback event: Check if a sheet should be generated for the device
  zipabox.events.OnAfterLoadDevice = function(device) {
    writelog("OnAfterLoadDevice");
  }
  
  // Callback event: Load all devices from Zipabox after connect
  zipabox.events.OnAfterConnect = function(){
    writelog("OnAfterConnect");
    zipabox.LoadDevices();	
  };
  
  // Callback event: Get logs for debug after disconnect
  zipabox.events.OnAfterDisconnect = function() {
    //var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
    //sheet.getRange(1, 1).setValue(Logger.getLog());
  };
  
  /****************************
   * ### FUNCTION MAIN PART ###
   ****************************/
  
  // Connecting to Zipabox  
  zipabox.Connect();
}


/**
 * Send prepared feeds to sense
 */
function sendToSense() {
  /**************************
   * ### CALLBACK EVENTS ###
   **************************/
  
  // Callback event: Retrieving data from "SENSE" section in the spreadsheet before sending feeds
  opensense.events.OnBeforeSendFeeds = function(){
    var senseSection = getPropertiesRangeByName("SENSE").getValues();
    
    for(var i=0; i< senseSection.length; i++) {
      switch(senseSection[i][0]) {
        case 'apikey':
          opensense.sense_key = senseSection[i][1];
          break;
          
        case 'url':
          opensense.url = senseSection[i][1];
          break;
      }
    }    
  };
  
  // Callback event: Insert a record in spreadsheet before sending feed
  opensense.events.OnBeforeSendFeed = function(feed){
    writelog("Sending Feed : " + (""+JSON.stringify(feed)));
        
    // Insert a record in the spreadsheet
    /*
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadSheet.getSheetByName(feed.deviceName)) {
      spreadSheet.insertSheet(feed.deviceName);
    }
    
    writelog("Insert values in spreadsheet for device ["+feed.deviceName+"]");
    var valueFR = feed.value.replace(".", ","); // French format with 2 decimals
    _insertRecord(feed.deviceName, valueFR);
    */
  };
  
  opensense.events.OnAfterSendFeed = function(feed){
    writelog("Feed sended : " + (""+JSON.stringify(feed)));	    
  };
  
  // Callback event : disconnect zipabox after sending all feeds
  opensense.events.OnAfterSendFeeds = function() {
    zipabox.Disconnect();
  };
  
  /****************************
   * ### FUNCTION MAIN PART ###
   ****************************/
  
  opensense.sendfeeds();
}


/**
 * Preparing lights data before sending to sense
 * TODO: everything
 */
function processLights() {
  writelog("*** processLights ***");
  
  
  /****************************
   * ### FUNCTION MAIN PART ###
   ****************************/
  /*
  // Reinit feeds
  opensense.emptyFeeds();
  
  // Retrieve data from Zipabox
  if(!zipabox.connected)
    collectDataFromZipabox();    
    
  // Collecting feeds for all devices
  for (var iddevice in zipabox.devices) {
    var device = zipabox.devices[iddevice];    
    
    for (var uuid in device.json){				
      var devicejson = device.json[uuid];
      var name = devicejson.name;
      var endpoint = devicejson.endpoint;
      
      for (var attr in devicejson.attributes){				
        var attribute = devicejson.attributes[attr];
        var value = attribute.value;
        
        //writelog("Attribute["+attr+"]: "+JSON.stringify(attribute));
        
        // Processing enum values
        if(attribute.definition && attribute.definition.type && attribute.definition.type == "Enum"){
          value = attribute.definition.enumValues[attribute.value];
        }
        
        CollectValuesForFeeds(endpoint,name,value);
      }
    }	
  }*/
}
  

/**
 * Process sensors data for:
 * - Sending to sense
 * - Insert value in sheet
 */
function processSensors() {
  writelog("*** processSensors ***");
  
  /****************************
   * ### FUNCTION MAIN PART ###
   ****************************/
  
  // Retrieve data from Zipabox
  if(!zipabox.connected)
    collectDataFromZipabox();
  
  // Get meters devices in spreadsheet
  var deviceMeters = getPropertiesRangeByName("sensors");  
  var search = deviceMeters.getValues();
    
  for (var i=2; i<search.length; i++) {
    var endointUUID = search[i][1];
    var devicejson = zipabox.GetDeviceByEndpointUUID(endointUUID);    
    
    if(!devicejson) break;
    
    var name = devicejson.name;
    
    writelog("Device["+endointUUID+"]: "+JSON.stringify(devicejson));
    
    for (var attr in devicejson.attributes) {
      var attribute = devicejson.attributes[attr];      
      
      _getSensorState(attribute, name, endointUUID);
      writelog("Attribute["+attr+"]: "+JSON.stringify(attribute));
    }     
  }
}
  
/**
 * Preparing meters data before sending to Sense
 * @param typeToProcess: "ALL" by default
 * - TEMPERATURE for temperature only
 * - HUMIDITY for humidity only
 * - LUMINANCE for luminance only
 * - CURRENT_CONSUMPTION for current consumption only
 * - ALL for all of above
 */
function processMeters(typeToProcess) {  
  writelog("*** processMeters ***");
  
  /****************************
   * ### FUNCTION MAIN PART ###
   ****************************/  
  
  // Check function argument
  if (typeof typeToProcess != "string")
    typeToProcess = "ALL";
  
  // Retrieve data from Zipabox
  if(!zipabox.connected)
    collectDataFromZipabox();
  
  // Get meters devices in spreadsheet
  var deviceMeters = getPropertiesRangeByName("meters");  
  var search = deviceMeters.getValues();
    
  for (var i=2; i<search.length; i++) {
    var endointUUID = search[i][1];
    var devicejson = zipabox.GetDeviceByEndpointUUID(endointUUID);
    
    if(!devicejson) break;
    
    var name = devicejson.name;
    writelog("Device["+endointUUID+"]: "+JSON.stringify(devicejson));
    
    for (var attr in devicejson.attributes) {
      var attribute = devicejson.attributes[attr];
      var attributeName = (typeof attribute.definition != "undefined") ? attribute.definition.name : attribute.name;
      
      if(attributeName != search[i][2]) break;
      
      attribute.canCreateSheet = search[i][5];
      
      writelog("Search: "+search.join("/"));
      writelog("Attribute["+attr+"]: "+JSON.stringify(attribute));
    
      //switch on typetoprocess     
      switch(typeToProcess) {
        case "TEMPERATURE":
          //Get temperature
          _getTemperature(attribute, name, devicejson.endpoint);            
          break;
          
        case "HUMIDITY":
          _getHumidity(attribute, name, devicejson.endpoint);          
          break;
          
        case "LUMINANCE":
          _getLuminance(attribute, name, devicejson.endpoint);         
          break;
          
        case "CURRENT_CONSUMPTION":
          _getCurrentConsumption(attribute, name, devicejson.endpoint);
          break;
          
        case "CUMULATIVE_CONSUMPTION":
          _getCumulativeConsumption(attribute, name, devicejson.endpoint);
          break;
          
        case "ALL":
          _getTemperature(attribute, name, devicejson.endpoint);
          _getHumidity(attribute, name, devicejson.endpoint);
          _getLuminance(attribute, name, devicejson.endpoint);
          _getCurrentConsumption(attribute, name, devicejson.endpoint);
          _getCumulativeConsumption(attribute, name, devicejson.endpoint);
          break;          
      } // end switch      
    } // end attributes
  }	// end devices
}


/**
 * Initialize the param sheet with devices
 */
function initParamSheet() {
  writelog("*** initParamSheet ***");
  
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var paramName = CacheService.getPrivateCache().get("paramSheet");
  var sheet = spreadSheet.getSheetByName("Paramètres");
  var columnsDevice = ["Name", "UUID Endpoint", "Attribute", "Feed ID", "Send to Sense", "Generate sheet"];
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(["FALSE", "TRUE"], true).build();
  
  // Clear data
  sheet.getRange("A13:F900").clear();
  
  /***************************
   * ### Processing Meters ###
   ***************************/  
  
  var lastRow = sheet.getLastRow()+2;
  var activeRange = sheet.getRange(lastRow, 1, 1, columnsDevice.length);
  
  activeRange.getCell(1, 1).setValue("meters");
  sheet.getRange(lastRow++, 1, 1, columnsDevice.length).setFontColor("white").setBackground("#003366");
  sheet.appendRow(columnsDevice);
  sheet.getRange(lastRow++, 1, 1, columnsDevice.length).setFontColor("white").setBackground("#004C99");
  
  // Retrieve data from Zipabox
  if(!zipabox.connected)
    collectDataFromZipabox();
  
  // Get only meters devices
  var deviceMeters = null;    
  for (var iddevice in zipabox.devices) {
    var device = zipabox.devices[iddevice];    
    
    if(device.name == "meters"){ 
      deviceMeters = device;    
      break;
    }
  }   
    
  for (var uuid in deviceMeters.json) {
    var devicejson = deviceMeters.json[uuid];
    var name = devicejson.name;      
    var endpoint = devicejson.endpoint;
    
    writelog("Device["+uuid+"]: "+JSON.stringify(devicejson));
    
    for (var attr in devicejson.attributes) {
      var attribute = devicejson.attributes[attr];      
      var nameAttribute = (typeof attribute.definition != "undefined") ? attribute.definition.name : attribute.name;
      
      writelog("Attribute["+attr+"]: "+JSON.stringify(attribute));
            
      var result = sheet.appendRow([name, endpoint, nameAttribute, '', false, false]);
      sheet.getRange(result.getLastRow(), 1, 1, 3).setBackground("#C9C7C5");      
      result.getDataRange().getCell(result.getLastRow(), 5).setDataValidation(rule);
      result.getDataRange().getCell(result.getLastRow(), 6).setDataValidation(rule);
      /* named range doesn't work for now... (bug in google script ?) */
      //spreadSheet.setNamedRange(devicejson.endpoint, sheet.getRange(result.getLastRow(), 1, 1, sheet.getLastColumn()));
    }
  }
  
  
  /****************************
   * ### Processing Sensors ###
   ****************************/  
  
  lastRow = sheet.getLastRow()+2;
  activeRange = sheet.getRange(lastRow, 1, 1, columnsDevice.length);
    
  activeRange.getCell(1, 1).setValue("sensors");
  sheet.getRange(lastRow++, 1, 1, columnsDevice.length).setFontColor("white").setBackground("#003366");
  sheet.appendRow(columnsDevice);
  sheet.getRange(lastRow++, 1, 1, columnsDevice.length).setFontColor("white").setBackground("#004C99");
  
  // Get only sensors devices
  var deviceMeters = null;    
  for (var iddevice in zipabox.devices) {
    var device = zipabox.devices[iddevice];    
    
    if(device.name == "sensors"){ 
      deviceMeters = device;    
      break;
    }
  }
  
  for (var uuid in deviceMeters.json) {
    var devicejson = deviceMeters.json[uuid];
    var name = devicejson.name;
    var endpoint = devicejson.endpoint;
    
    writelog("Device["+uuid+"]: "+JSON.stringify(devicejson));    
    
    for (var attr in devicejson.attributes) {
      var attribute = devicejson.attributes[attr];
      var nameAttribute = (typeof attribute.definition != "undefined") ? attribute.definition.name : attribute.name;
      
      writelog("Attribute["+attr+"]: "+JSON.stringify(attribute));
            
      var result = sheet.appendRow([name, endpoint, nameAttribute, '', false, false]);
      sheet.getRange(result.getLastRow(), 1, 1, 3).setBackground("#C9C7C5");
      result.getDataRange().getCell(result.getLastRow(), 5).setDataValidation(rule);
      result.getDataRange().getCell(result.getLastRow(), 6).setDataValidation(rule);  
      /* named range doesn't work for now... (bug in google script ?) */
      //spreadSheet.setNamedRange(devicejson.endpoint, sheet.getRange(result.getLastRow(), 1, 1, sheet.getLastColumn()));
    }
  }
  
  /****************************
   * ### Processing Lights ###
   ****************************/  
  
  lastRow = sheet.getLastRow()+2;
  activeRange = sheet.getRange(lastRow, 1, 1, columnsDevice.length);
    
  activeRange.getCell(1, 1).setValue("lights");
  sheet.getRange(lastRow++, 1, 1, columnsDevice.length).setFontColor("white").setBackground("#003366");
  sheet.appendRow(columnsDevice);
  sheet.getRange(lastRow++, 1, 1, columnsDevice.length).setFontColor("white").setBackground("#004C99");
  
  // Get only sensors devices
  var deviceMeters = null;    
  for (var iddevice in zipabox.devices) {
    var device = zipabox.devices[iddevice];    
    
    if(device.name == "lights"){ 
      deviceMeters = device;    
      break;
    }
  }
  
  for (var uuid in deviceMeters.json) {
    var devicejson = deviceMeters.json[uuid];
    var name = devicejson.name;
    var endpoint = devicejson.endpoint;
    
    writelog("Device["+uuid+"]: "+JSON.stringify(devicejson));    
    
    for (var attr in devicejson.attributes) {
      var attribute = devicejson.attributes[attr];
      var nameAttribute = (typeof attribute.definition != "undefined") ? attribute.definition.name : attribute.name;
      
      writelog("Attribute["+attr+"]: "+JSON.stringify(attribute));
            
      var result = sheet.appendRow([name, endpoint, nameAttribute, '', false, false]);
      sheet.getRange(result.getLastRow(), 1, 1, 3).setBackground("#C9C7C5");
      result.getDataRange().getCell(result.getLastRow(), 5).setDataValidation(rule);
      result.getDataRange().getCell(result.getLastRow(), 6).setDataValidation(rule);  
      /* named range doesn't work for now... (bug in google script ?) */
      //spreadSheet.setNamedRange(devicejson.endpoint, sheet.getRange(result.getLastRow(), 1, 1, sheet.getLastColumn()));
    }
  }
  
}


/**
 * Check if a sheet can be created for the device
 * @param uuid : endpoint of the device
 * @param tabName : name of the params tab
 */
/*
function _canCreateSheetForDeviceUUID(uuid, tabName) {
  writelog("*** _canCreateSheetForDeviceUUID ***");
  
  // Check if sheet should be created for the device
  var range = getRowByUUID(uuid, tabName);
  
  if(range == -1) { 
    writelog("No row found for device: "+uuid);
    return false;
  }
  
  if(range.getCell(1, 5).getValue()) return true;
  
  writelog("Creating sheet disabled for device: "+uuid);
  return false;  
}
*/


/**
 * Custom function for periodical execution by google script
 */
function main(){
  //Logger.clear();
  processMeters("ALL");
  processSensors();
  processLights();
  sendToSense();
}
