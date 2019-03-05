{\rtf1\ansi\ansicpg1252\cocoartf1561\cocoasubrtf600
{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 var LOG_SHEET_NAME = 'Unsampled Report Logs';\
var ss = SpreadsheetApp.getActive();\
\
//var ui = SpreadsheetApp.getUi();\
\
function requestReport() \{\
  var ss = SpreadsheetApp.getActive();\
  var reportConfigSheet = ss.getSheetByName('Report Configuration');\
  var t =2;\
  var numberOfColumns = reportConfigSheet.getRange("C1").getValue();\
 \
  //adding more columns \
  for (var i=1; i<=numberOfColumns; i++)\{\
  var startDateRange = reportConfigSheet.getRange(5,t); \
  var startDate = startDateRange.getValue();\
  var endDateRange = reportConfigSheet.getRange(6,t); \
  var endDate = endDateRange.getValue();\
  var metricsRange = reportConfigSheet.getRange(8,t); \
  var metrics = metricsRange.getValue();\
  var dimensionsRange = reportConfigSheet.getRange(9,t); \
  var dimensions = dimensionsRange.getValue();\
  var filtersRange = reportConfigSheet.getRange(11,t); \
  var filters = filtersRange.getValue();\
  var titleRange = reportConfigSheet.getRange(2,t); \
  var title = titleRange.getValue();\
  var sortRange = reportConfigSheet.getRange(10,t); \
  var sort = sortRange.getValue();\
  \
   if (title != "" | startDate != "" | endDate != "" | metrics != "" | dimensions != "" )\{\
   \
  var resource = \{\
        'title': title,\
        'start-date': startDate,\
        'end-date': endDate,\
        'metrics': metrics,\
        'dimensions': dimensions,\
        'sort': sort,\
        'filters': filters\
      \};\
   var ss = SpreadsheetApp.getActive();\
   var reportConfigSheet = ss.getSheetByName('Report Configuration');\
   var accountId = reportConfigSheet.getRange(18,t).getValue();\
   var webPropertyId = reportConfigSheet.getRange(19,t).getValue();\
   var profileId = reportConfigSheet.getRange(20,t).getValue();\
\
\
  try \{\
    var request = Analytics.Management.UnsampledReports.insert(resource, accountId, webPropertyId, profileId);\
    \
  \} catch (error) \{\
    ui.alert('Error Performing Unsampled Report Query', error.message, ui.ButtonSet.OK);\
    return;\
  \}\
\
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);\
\
  if (!sheet) \{\
    sheet = ss.insertSheet(LOG_SHEET_NAME);\
    sheet.appendRow(['User', 'Account', 'Web Property', 'View', 'Title', 'Inserted Time', 'Updated Time', 'Status', 'Id', 'File']);\
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');\
  \}\
  sheet.appendRow([\
    Session.getEffectiveUser().getEmail(),\
    request.accountId,\
    request.webPropertyId,\
    request.profileId,\
    request.title,\
    request.created,\
    request.updated,\
    request.status,\
    request.id\
  ]);\
 t++;\
  \}else\{\
    var failRecords = "Some reports were not run. Please make sure all the fields are correct."\
    Browser.msgBox(failRecords);\
  \}\
  \
\}\
  \
\}\
 \
// Scans LOG_SHEET_NAME and tries to update any report that is PENDING\
function updateAllReports() \{\
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);\
\
  var lastRow = sheet.getLastRow();\
\
  var dataRange = sheet.getRange(2,1, lastRow, 10);\
\
  var data = dataRange.getValues();\
\
  for (var i=0; i<data.length; i++) \{\
    // If data is PENDING let's try to update it's status. Hopefully it's complete now\
    // but it may take up to 24h to process an Unsampled Reprot\
    if (data[i][0] == Session.getEffectiveUser().getEmail() && data[i][7] == 'PENDING') \{\
      try \{\
      var request = Analytics.Management.UnsampledReports.get(data[i][1], data[i][2], data[i][3], data[i][8]);\
      \} catch (error) \{\
        ui.alert('Error Performing Unsampled Report Query', error.message, ui.ButtonSet.OK);\
        return;\
      \}\
\
      data[i] = [\
        Session.getEffectiveUser().getEmail(),\
        request.accountId,\
        request.webPropertyId,\
        request.profileId,\
        request.title,\
        request.created,\
        request.updated,\
        request.status,\
        request.id,\
        request.status == 'COMPLETED' ? DriveApp.getFileById(request.driveDownloadDetails.documentId).getUrl() : ''\
      ];\
\
\
      // If data is Complete let's import it into a new sheet\
      if (request.status == 'COMPLETED') \{\
        importReportFromDrive(request.title, request.driveDownloadDetails.documentId);\
      \}\
    \}\
  \}\
\
  // Write only once to the spreadsheet this is faster\
  dataRange.setValues(data);\
\
\}\
\
function importReportFromDrive(title, fileId) \{\
  var file = DriveApp.getFileById(fileId);\
  var csvString = file.getBlob().getDataAsString();\
  var data = Utilities.parseCsv(csvString);\
  var ss = SpreadsheetApp.getActiveSpreadsheet();\
\
  // Check if the sheet already exists\
  \
  if (ss.getSheetByName(title)==null)\{\
    var i=1;\
    var sheetName = title;\
    while (ss.getSheetByName(sheetName)) \{\
      sheetName = title + ' ('+ i++ +')';\
    \}\
    \
    var sheet = ss.insertSheet(sheetName);\
    var range = sheet.getRange(2, 1, data.length, data[0].length);\
    range.setValues(data);\
    //sheet.insertRows(1);\
    var date = Utilities.formatDate(new Date(), "PST", "MM/dd/yyyy '@' hh:mm a");\
    sheet.getRange("A1").setValue("Last updated: " +date);\
    \
  \}else\{\
 \
    var dataSheet = ss.getSheetByName(title);\
    dataSheet.clear();\
    var range = dataSheet.getRange(2, 1, data.length, data[0].length);\
    range.setValues(data);\
    //dataSheet.insertRows(1);\
    var date = Utilities.formatDate(new Date(), "PST", "MM/dd/yyyy '@' hh:mm a");\
    dataSheet.getRange("A1").setValue("Last updated: " +date);\
  \}\
  \
\}\
\
  \
\
\
\
}

// testing commits
//testing 3rd commit