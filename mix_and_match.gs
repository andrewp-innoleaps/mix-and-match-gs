function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mix and match')
    .addItem('Create new sheets', 'CreateSheet')
    .addItem('Generate Match', 'savetoDatabase')
    .addToUi();
}

function CreateSheet() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const insertSheet = activeSpreadsheet.insertSheet();
  insertSheet.setName("Startups");
  const insertAttendeesSheet = activeSpreadsheet.insertSheet();
  insertAttendeesSheet.setName("Attendees");
}

function savetoDatabase() {
  const startups = compileStartup();
  const attendees = compileAttendees();
  const activeSheet = SpreadsheetApp.getActive().getName();
  const formBody = {
    startups: startups,
    attendees:attendees,
    event_name: activeSheet
  }
    var options = {
    'method' : 'post',
    'payload' : JSON.stringify(formBody),
    'headers': {
      'Content-Type' : 'application/json',
      'Accept' : 'application/json',
    }
  }
    var response = UrlFetchApp.fetch(`https://mix-and-match-zknig.ondigitalocean.app/event`, options);
    const responseContext = JSON.parse(response.getContentText())
    renderToSheet(responseContext["suggestions"])
}

function renderToSheet(data) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(newSuggestionsSheet());
  for (var i in data) {
    sheet.appendRow(data[i]);
  }
}

function compileStartup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Startups');
  const range = sheet.getDataRange(); // determine the range of populated data
  const data = range.getValues(); // get the actual data in an array data[row][column]
  const lastCol = range.getLastColumn();
  let obj = {};
  let clone
  let sessions = []
    for(let j=1; j<lastCol; j++){
      const newData = data[0]
      for(let i=1; i<data[0].length; i++){
        data.forEach(function(row) {
          obj[`${row[0]}`]=row[j];
          delete obj.Name;
          clone = {...obj};
        });
      }
      sessions.push({[newData[j]]:clone})
    }
    return sessions
}

function compileAttendees(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Attendees');
  const range = sheet.getDataRange(); // determine the range of populated data
  const data = range.getValues(); // get the actual data in an array data[row][column]

  const lastRow = range.getLastRow();
  const numColumns = range.getNumColumns();

  let attendees = []
  for(let row=1; row<lastRow; row++){
    let startupPreference = {}
    for(let col=1; col<numColumns; col++) {
      startupPreference[`${data[0][col]}`] = isPreferenceValid(`${data[row][col]}`)
    }
    const tempObj = {};
    tempObj[`${data[row][0]}`] = startupPreference;
    attendees.push(tempObj);
  }
  return attendees
}

function isPreferenceValid(preference) {
  const validPreferences =["yes","maybe", "no"]
  const isIncluded = validPreferences.includes(preference.toLowerCase())
  if(isIncluded) {
    return preference
  } else {
     SpreadsheetApp.getUi().alert("Invalid Preference");
  }
 
}
function newSuggestionsSheet(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  let sheetNameArray = [];
   for (var i = 0; i < sheets.length; i++) {
     if(sheets[i].getName().toLowerCase().includes('suggestions')) {
      sheetNameArray.push(sheets[i].getName());
     }
     sheetNameArray.push('Suggestions')
  }
  const sortarray = sheetNameArray.sort()
  return incrementString(sortarray[sortarray.length-1])
}

function incrementString(str) {
  var count = str.match(/\d*$/);

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const insertAttendeesSheet = activeSpreadsheet.insertSheet();
  const newSheet =  str.substr(0, count.index) + (++count[0]);
  insertAttendeesSheet.setName(newSheet);
  console.log({newSheet})
  return newSheet
};
