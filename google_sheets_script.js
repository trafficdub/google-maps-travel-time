/*
Google Maps Current Travel Times Recorder
Version: v1.0
Created By: Yuzhu Huang @Jacobs (based on work of Kyle Yasumiishi, @chicagocomputerclasses, and Google Online Guide)
Date: 11/6/2018
Description: This program automatically calls the Google Directions API, and retrieve the current travel time (default: driving) for a route with origin, destination and waypoint specified on the Spreadsheet.
**Note**: Please use obtain and use your own API key.
**Changes Log:
11/6/2018 12:00 pm: Changed from calling in-google G Suite Maps service to get "base" travel time from DirectionFinder, to calling Directions API and use "duration_in_traffic" to get real-time travel time.
11/6/2018 03:00 pm: Added getGeoCode method to allow using either text address or lat/lng coordinates in "Start" and "End" columns.
11/6/2018 10:45 pm: Added testing of content in "StartGeoCode" and "EndGeoCode" columns, to avoid calling Geocoding API repeatedly.
11/6/2018 11:45 pm: Used getValues to replace for looping in getRouteArray (removed).
*/
function writeCurrentVehTT() {
  var app = SpreadsheetApp;
  var activeSheet = app.getActiveSpreadsheet().getActiveSheet();
  
  var lastColNum = activeSheet.getLastColumn();
  var headerArray = getFields(activeSheet, 1, 1, 1, lastColNum); 
  
  var routeIdColNum = headerArray.indexOf("Route ID") + 1;
  var startColNum = headerArray.indexOf("Start") + 1;
  var endColNum = headerArray.indexOf("End") + 1;
  var wayPointColNum = headerArray.indexOf("Must Pass") + 1;
  var startGeoColNum = headerArray.indexOf("StartGeoCode") + 1;
  var endGeoColNum = headerArray.indexOf("EndGeoCode") + 1;
  var routeLastRowNum = activeSheet.getRange(1,routeIdColNum,activeSheet.getMaxRows(),1).getValues().filter(String).length;
  var routeCount = routeLastRowNum - 1;

  
  var routesArray = activeSheet.getRange(2, startColNum, routeLastRowNum - 1, endGeoColNum - startColNum + 1).setNumberFormat("@STRING@").getValues();

  
  var VehRouteRecordedColNum = headerArray.indexOf("Veh Route Recorded") + 1
  var VehTtLastRowNum = activeSheet.getRange(1,VehRouteRecordedColNum,activeSheet.getMaxRows(),1).getValues().filter(String).length;
  var VehTimeStampColNum = headerArray.indexOf("Time Recorded") + 1;
  var VehDurSecColNum = headerArray.indexOf("Duration (sec)") + 1;
  var VehDurMinColNum = headerArray.indexOf("Duration (min)") + 1;
  var VehDistColNum = headerArray.indexOf("Distance (mile)") + 1;
  var VehViaColNum = headerArray.indexOf("via") + 1;
  
  //Get Route information for every route in the array
  for (var i = 0; i < routeCount; i++) {
    //API URL Parameters - origin, destination, wayPoint, Mode, Alternatives, Units, Departure_Time, Traffic_Mode
    
    if (routesArray[i][3] == "" || routesArray[i][4] == "") {
      var originGeo = getGeoCode(routesArray[i][0]);
      var destGeo = getGeoCode(routesArray[i][1]);
      activeSheet.getRange(2 + i, startGeoColNum).setValue(originGeo);
      activeSheet.getRange(2 + i, endGeoColNum).setValue(destGeo);  
    } else {
      var originGeo = routesArray[i][3].replace(" ","");
      var destGeo = routesArray[i][4].replace(" ","");
    }
    
  }
  
}

function getApiJsonData(activeSheet, originGeo, destGeo, wayPointGeo) {
  //Directions API Info
  var endPoint = "https://maps.googleapis.com/maps/api/directions/";
  var outputFormat = "json";
  var directionsApiKey = activeSheet.getRange("B17").setNumberFormat("@STRING@").getValue();
  
  var mode = "driving"; //options: "driving", "walking", "bicycling", "transit"
  var alternatives = "false"; //options: "false", "true"
  var units = "imperial"; //options: "metric, "imperial"
  var departureTime = "now";
  //var trafficModel = "best_guess"; //options: "best_guess", "pessimistic", "optimistic"
  
  //Parse API URL
  var urlCall = endPoint + outputFormat
  + "?origin=" + originGeo
  + "&destination=" + destGeo
  + "&mode=" + mode
  + "&waypoints=via:" + wayPointGeo
  + "&alternatives=" + alternatives
  + "&units=" + units
  + "&departure_time=" + departureTime
  //+ "&traffic_model=" + trafficModel
  + "&key=" + directionsApiKey;
  // Call API and pase JSON to native object
  var response = UrlFetchApp.fetch(urlCall, {'muteHttpExceptions': true});
  var json = response.getContentText();
  var data = JSON.parse(json);
  Logger.log(data);
  return data;
}


function getGeoCode(locAddress) {
  if (locAddress.indexOf(".") <= -1) {
    var response = Maps.newGeocoder().geocode(locAddress);
    for (var j = 0; j < response.results.length; j++) {
      var result = response.results[j];
      var resultLat = result.geometry.location.lat;
      var resultLong = result.geometry.location.lng;
      var resultCord = resultLat.toString() + "," + resultLong.toString();
    }
  } else {
    var resultCord = locAddress.replace(" ","");
  }
  return resultCord;
}


function getFields(activeSheet, row, col, numRows, numCols) {
  var fieldRange = activeSheet.getRange(row, col, numRows, numCols).setNumberFormat("@STRING@").getValues();
  var fieldArray = [];
  for (var i = 0; i < fieldRange[0].length; i++) {
    fieldArray.push(fieldRange[0][i]);
  }
  return fieldArray;
} 

function getDataFromJson(routesJson, return_type) {

  var route = routesJson.routes[0];
  var getTheLeg1 = route["legs"][0];
  var meters1 = getTheLeg1["distance"]["value"];
  var meters = meters1;
  var duration1 = getTheLeg1["duration_in_traffic"]["value"];
  var duration = duration1;
  var summary = route["summary"];
  
  switch(return_type){
    case "miles":
      return meters * 0.000621371;
      break;
    case "minutes":
        //convert to minutes and return
        return duration / 60;
      break;    
    case "seconds":
      return duration;
      break;
    case "summary":
      return summary;
      break;
    default:
      return "Error: Wrong Unit Type";
   }
}

function saveData(){
  //var newSheetName = "Sheet2";
  var ssId = "1_bfzHTEkwcLAnlFnZ1TTojtqJFJfM9ez62ZRmvHAwss";
  var dataSheet = SpreadsheetApp.openById(ssId).getSheetByName("Sheet1");
  var dataLastRow = dataSheet.getRange("O:O").getValues().filter(String).length.toString();
  var dataRng = dataSheet.getRange("N1:S" + dataLastRow).getValues();
  var cfaSsId = "1K1EGka90hWTQ6Y7AjkoY-MynEB4DCEa0hx7ZQ7_R23s";
  var cfaSs = SpreadsheetApp.openById(cfaSsId);
  var today = new Date();
  var newSheetName = (today.getMonth() + 1) + '-' + today.getDate() + ' ' + today.getHours() + ":" + today.getMinutes();
  Logger.log(newSheetName);
  var newSheet = cfaSs.insertSheet(newSheetName);
  newSheet.getRange(1, 1, Number(dataLastRow), 6).setValues(dataRng);
  newSheet.getRange(1, 2, Number(dataLastRow), 1).setNumberFormat("m/d/yyyy hh:mm:ss AM/PM")
  
  if (dataLastRow < 2) {
   dataLastRow = 2; 
  }
  var dataRng = dataSheet.getRange("N2:S" + dataLastRow).clearContent();
}

function startMasterTrigger(){
  ScriptApp.newTrigger("startMasterTrigger")
  .timeBased()
  .atHour(16)
  .everyHours(1)
  .create();
  
  FiveMinTimeTrigger();
  
  ScriptApp.newTrigger("deleteTriggers")
  .timeBased()
  .after(0.8*60*60*1000)
  .create();
  saveData();
}

function deleteTriggers(){
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log(triggers.length);
  try {
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() == "writeCurrentVehTT" 
          || triggers[i].getHandlerFunction() == "deleteTriggers") {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  }
  catch(err) {
    try {
      for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() == "writeCurrentVehTT" 
            || triggers[i].getHandlerFunction() == "deleteTriggers") {
          ScriptApp.deleteTrigger(triggers[i]);
        }
      }
    }
    catch(err) {
      for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() == "writeCurrentVehTT" 
            || triggers[i].getHandlerFunction() == "deleteTriggers") {
          ScriptApp.deleteTrigger(triggers[i]); 
        }
      }
    }
  }
}

function FiveMinTimeTrigger(){
  ScriptApp.newTrigger("writeCurrentVehTT")
   .timeBased()
   .everyMinutes(5)
   .create();
}

/*Resources:
API Billing: https://developers.google.com/maps/previous-pricing
External URL: https://developers.google.com/apps-script/guides/services/external
Directions API Guide: https://developers.google.com/maps/documentation/directions/intro
Trigger Quota: https://developers.google.com/apps-script/guides/services/quotas
Currrent Quota: https://script.google.com/dashboard
Geocoding: https://developers.google.com/apps-script/reference/maps/geocoder#geocode(String)

Archived: (only able to get Directions, not current travel time - duration_in_traffic)
https://www.chicagocomputerclasses.com/google-sheets-google-maps-function-distance-time/
https://developers.google.com/apps-script/reference/maps/mode
https://developers.google.com/apps-script/reference/maps/direction-finder
https://developers.google.com/apps-script/guides/triggers/installable#limitations
https://stackoverflow.com/questions/3018875/is-it-possible-to-automate-google-spreadsheets-scripts-e-g-without-an-event-to
https://webapps.stackexchange.com/questions/16009/display-sheet-name-in-google-spreadsheets

*/
