var api_key = "keyXXXXXXXXXXXX"; //ADD YOUR API KEY FROM AIRTABLE HERE
var baseID = "appXXXXXXXXXXXX"; //ADD YOUR BASE ID HERE
var tablesToSync_fromSheetRange = "A14:B16"; //UPDATE CELL RANGE HERE (for tables that you want to sync)

////////// add items to UI menu ///////////
function onOpen(e) {
   SpreadsheetApp.getUi()
       .createMenu('Airtable to google sheets sync')
       .addItem('Manually sync all data', 'syncData')
       .addToUi();
 }

////////// function to trigger the entire data syncing operation ///////////

function syncData(){
  //fetch table names from the control panel of the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tablesToSync = ss.getSheetByName("Control panel").getRange(tablesToSync_fromSheetRange).getValues();
    
  //sync each table
  for (var i = 0; i<tablesToSync.length; i++){
    var tableName = tablesToSync[i][0];
    var viewID = tablesToSync[i][1];
    var airtableData = fetchDataFromAirtable(tableName, viewID);
    pasteDataToSheet(tableName, airtableData);
    
    //wait for a bit so we don't get rate limited by Airtable
    Utilities.sleep(201);
  }
}

/////////////////////////////

function saveFormulas(dataSheets){
  //add the control panel to our list of data-related sheets
  dataSheets.push("Control panel");
  
  //initialise the object which will hold all formulas
  var formulas = {};
  
  //get all sheets in spreadsheet
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
  //iterate through the sheets - if they're not used for Airtable data stuff, then save all their formulas
  for (var i in sheets){
    var sheetName = sheets[i].getSheetName();
    if (dataSheets.indexOf(sheetName) == -1){
      formulas[sheetName] = sheets[i].getRange("A:Z").getFormulas();
    }
  }
  return formulas;
}

////////// take airtable data and paste it into a sheet ///////////////////////////

function pasteDataToSheet(sheetName, airtableData){
  
  //define field schema, which will be added to the row headers
  var fieldNames = ["Record ID"];
  //add every single field name to the array
  for (var i = 0; i<airtableData.length; i++){
    for (var field in airtableData[i].fields){
      fieldNames.push(field);
    }
  }
  //remove duplicates from field names array
  fieldNames = fieldNames.filter(function(item, pos){
    return fieldNames.indexOf(item)== pos;
  });
  
  //select the sheet we want to update, or create it if it doesn't exist yet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet; 
  if (ss.getSheetByName(sheetName) == null){
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet = ss.getSheetByName(sheetName);
  }

  //clear data from sheet
  sheet.clear();
  
  //add field names to sheet as row headers, and format the headers
  var headerRow = sheet.getRange(1,1,1,fieldNames.length);
  headerRow.setValues([fieldNames]).setFontWeight("bold").setWrap(true);
  sheet.setFrozenRows(1);
  
  //add Airtable record IDs to the first column of each row
  for (var i = 0; i<airtableData.length; i++){
    sheet.getRange(i+2,1).setValue(airtableData[i].id);
  }
  
  //// add other data to rows ////
  //for each record in our Airtable data...
  for (var i = 0; i<airtableData.length; i++){
    //iterate through each field in the record
    for (var field in airtableData[i].fields){
      sheet.getRange(i+2,fieldNames.indexOf(field)+1) //find the cell we want to update
        .setValue(airtableData[i].fields[field]); //update the cell 
    }
  }  
}

////////////// query the Airtable API to get raw data ///////////////////////

function fetchDataFromAirtable(tableName, viewID) {
  
  // Initialize the offset.
  var offset = 0;

  // Initialize the result set.		
  var records = [];

  // Make calls to Airtable, until all of the data has been retrieved...
  while (offset !== null){	

    // Specify the URL to call.
    var url = [
      "https://api.airtable.com/v0/", 
      baseID, 
      "/",
      encodeURIComponent(tableName),
      "?",
      "api_key=", 
      api_key,
      "&view=",
      viewID,
      "&offset=",
      offset
      ].join('');
    var options =
        {
          "method"  : "GET"
        };
    
    //call the URL and add results to to our result set
    response = JSON.parse(UrlFetchApp.fetch(url,options));
    records.push.apply(records, response.records);
    
    //wait for a bit so we don't get rate limited by Airtable
    Utilities.sleep(201);

    // Adjust the offset.
	// Airtable returns NULL when the final batch of records has been returned.
    if (response.offset){
      offset = response.offset;
    } else {
      offset = null;
    }
      
  }
  return records;
}

////////////////////////////////////
