function batchCreateDashboard() {
  var Count = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('demographics').getRange(2, 19).getValues();

  // Get all of the rows cached into arrays
  var SIDs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('demographics').getRange(0, 1, Count).getValues();
  var schools = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('demographics').getRange(0, 6, Count).getValues();
  var hrTeachs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('demographics').getRange(0, 19, Count).getValues();
  var first_names = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('demographics').getRange(0, 4, Count).getValues();
  var last_names = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('demographics').getRange(0, 3, Count).getValues();


  for(var i = 2; i<= Count; i+=1){
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('printData').getRange(3, 11).setValue(SIDs[i][0]);
    var studentName = first_names[i][0] + " " + last_names[i][0];
    var fileName = schools[i][0] + " - " + hrTeachs[i][0] + " - " + SIDs[i][0] + " - " + studentName + " - " + "Grade 2 - Spring 22 Data Dashboard.pdf";
    var fileNameGenesis = SIDs[i][0] + ".pdf"

    switch(schools[i][0]){
      case "BWD":
      var driveFolder = "1KOHZoejZZBqR4YfWP4X1VzHPqw-nimwU";
      break;

      case "CAD":
      var driveFolder = "17YdmNwG6PN1JspMeZeRTbFOOwF6PLdI2";
      break;
      
      case "DBO":
      var driveFolder = "1DIEEukKK4ahBJNoYqxYU22k1oAq1D-CZ";
      break;
      
      case "KDM":
      var driveFolder = "1FT0NUXsZD8N17MLoIu9W9eQNcqiascJD";
      break;
      
      case "SB":
      var driveFolder = "1IXxIOutlJjKRgTglraRYw--5O9kCyogz";
      break;                  
    }
    
    // Get the currently active spreadsheet URL (link)
    // Or use SpreadsheetApp.openByUrl("<>");
    // const ss = SpreadsheetApp.getActiveSpreadsheet(); unused

    // URL of Form Letter
    const url = 'https://docs.google.com/spreadsheets/d/10oFhuM4QVHaKLtjCmwT1KAx899P_6VCAKmd_cwvw8wA/export?'+
        'exportFormat=pdf'+      //Exports as PDF
        '&format=pdf'+           //PDF
        '&size=letter'+          //Letter sized paper
        '&portrait=true'+        //Portrait mode false means landscape
        '&sheetnames=false'+     //Include sheet names true/false
        '&printtitle=false'+     //Include title true/false
        '&pagenumbers=false'+    //Include Page numbers true/false
        '&gridlines=false'+      //Include gridlines true/false
        '&fzr=false'+            //dont know
        '&gid=1518815625'+       //gid of sheet 
        '&ir=false'+             //must be false
        '&ic=false'+             //must be false
        '&scale=4'               //1 Normal 2 Fit to Width 3 Fit to Height 4 Fit to Page
        '&top_margin=0&left_margin=0&right_margin=0&bottom_margin=0';  //set margins to 0

    var token = ScriptApp.getOAuthToken();
    // var sheets = ss.getSheets(); unused

    var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
    var response = UrlFetchApp.fetch(url,params).getBlob().setName(fileName);

    // Saves the file to the /Data Dasboards/Winter 22/Kindergarten/School folder
    var folder = DriveApp.getFolderById(driveFolder);
    var ff = folder.createFile(response);

    // Saves the file to the /Data Dasboards/Winter 22/Kindergarten/Genesis
    var genesisDriveFolder = DriveApp.getFolderById("1417xEJZK-U_bk2_5VNyYqIASIGKFtk2V");
    var copy=ff.makeCopy(fileNameGenesis,genesisDriveFolder);

    Utilities.sleep(1000);
  }
}

function createSingleDashboard() {

    var SID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('printData').getRange(3,11).getValue();
    var school = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('printData').getRange(6,11).getValue();
    var hrTeach = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('printData').getRange(6,12).getValue();
    var studentName = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('printData').getRange(3,12).getValue();
    var fileName = school + " - " + hrTeach + " - " + SID + " - " + studentName + " - " + "Grade 2 - Spring 22 Data Dashboard.pdf";
    var fileNameGenesis = SID + ".pdf"
    switch(school){
      case "BWD":
      var driveFolder = "1KOHZoejZZBqR4YfWP4X1VzHPqw-nimwU";
      break;

      case "CAD":
      var driveFolder = "17YdmNwG6PN1JspMeZeRTbFOOwF6PLdI2";
      break;
      
      case "DBO":
      var driveFolder = "1DIEEukKK4ahBJNoYqxYU22k1oAq1D-CZ";
      break;
      
      case "KDM":
      var driveFolder = "1FT0NUXsZD8N17MLoIu9W9eQNcqiascJD";
      break;
      
      case "SB":
      var driveFolder = "1IXxIOutlJjKRgTglraRYw--5O9kCyogz";
      break;                  
    }
    
    // Get the currently active spreadsheet URL (link)
    // Or use SpreadsheetApp.openByUrl("<>");
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // URL of Form Letter
    const url = 'https://docs.google.com/spreadsheets/d/10oFhuM4QVHaKLtjCmwT1KAx899P_6VCAKmd_cwvw8wA/export?'+
        'exportFormat=pdf'+      //Exports as PDF
        '&format=pdf'+           //PDF
        '&size=letter'+          //Letter sized paper
        '&portrait=true'+        //Portrait mode false means landscape
        '&sheetnames=false'+     //Include sheet names true/false
        '&printtitle=false'+     //Include title true/false
        '&pagenumbers=false'+    //Include Page numbers true/false
        '&gridlines=false'+      //Include gridlines true/false
        '&fzr=false'+            //dont know
        '&gid=1518815625'+       //gid of sheet 
        '&ir=false'+             //must be false
        '&ic=false'+             //must be false
        '&scale=4'               //1 Normal 2 Fit to Width 3 Fit to Height 4 Fit to Page
        '&top_margin=0&left_margin=0&right_margin=0&bottom_margin=0';  //set margins to 0

    var token = ScriptApp.getOAuthToken();
    var sheets = ss.getSheets();

    var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
    var response = UrlFetchApp.fetch(url,params).getBlob().setName(fileName);

    // Saves the file to the /Data Dasboards/Winter 22/Grade 1/School folder
    var folder = DriveApp.getFolderById(driveFolder);
    var ff = folder.createFile(response);

    // Saves the file to the /Data Dasboards/Winter 22/Grade 1/Genesis
    var genesisDriveFolder = DriveApp.getFolderById("1417xEJZK-U_bk2_5VNyYqIASIGKFtk2V");
    var copy=ff.makeCopy(fileNameGenesis,genesisDriveFolder);
  }



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Data Dashboards')
      .addSubMenu(ui.createMenu('Create')
        .addItem('◷ Create Single', 'createSingleDashboard')
        .addItem('◷ Create Whole Grade Level', 'batchCreateDashboard'))
      .addSeparator()
      //.addSubMenu(ui.createMenu('Email')
        //.addItem('◷ Send Payment Request', 'emailPaymentRequest'))
      .addToUi();
}
