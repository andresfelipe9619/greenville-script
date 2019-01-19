// import "google-apps-script";
var ACTIVE_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
var SELECTED_HOUSE = { data: null, row: null };

function onEdit(e) {
  var range = e.range;

  checkEditedCell(range);
}

function checkEditedCell(range) {
  if (range.getColumn() == 10 && range.getValue() == "OK") {
    handleOnGenerateFields(range);
  }
}

function handleOnGenerateFields(range) {
  var sheet = ACTIVE_SPREADSHEET.getActiveSheet();
  var selectedRow = range.getRow();
  var nextRow = selectedRow + 1;

  var nextRowValues = sheet.getSheetValues(
    nextRow,
    1,
    1,
    sheet.getLastColumn()
  )[0];

  var isEmpty = isNextRowEmpty(nextRowValues);
  if (!isEmpty) return;

  var rawHouse = sheet.getSheetValues(selectedRow, 1, 1, sheet.getLastColumn());
  var headers = sheet.getSheetValues(1, 1, 1, sheet.getLastColumn());
  var houseObj = sheetValuesToObject(rawHouse, headers)[0];
  var nextId = houseObj["id house"] + 1;
  var nextIdRange = sheet.getRange(nextRow, 1, 1, 1);
  nextIdRange.setValues([[nextId + ""]]);
  SELECTED_HOUSE.data = houseObj;
  SELECTED_HOUSE.row = selectedRow;
  addHouseToSheets();
}

function isNextRowEmpty(nextRowValues) {
  var empties = nextRowValues.filter(function(value) {
    return String(value) == "" || String(value) == " ";
  });
  Logger.log("isEmpty");
  Logger.log(empties);
  if (empties.length < 4) {
    return false;
  }
  return true;
}

function addHouseToSheets() {
  var sheets = ACTIVE_SPREADSHEET.getSheets();

  for (var sheet in sheets) {
    var range = [];
    if (sheet.getName.includes("BALANCE")) {
      range = sheet.getRange(SELECTED_HOUSE.row - 1, 1);
    } else {
      range = sheet.getRange(
        SELECTED_HOUSE.row - 2,
        1,
        1,
        sheet.getLastColumn()
      );
    }
    range.setValues([
      [SELECTED_HOUSE.data["id_house"], SELECTED_HOUSE["address"]]
    ]);
  }
}

function sheetValuesToObject(sheetValues, headers) {
  var headings =
    headers[0].map(String.toLowerCase) ||
    sheetValues[0].map(String.toLowerCase);
  if (sheetValues)
    var people = sheetValues.length > 1 ? sheetValues.slice(1) : sheetValues;

  var peopleWithHeadings = addHeadings(people, headings);

  function addHeadings(people, headings) {
    return people.map(function(personAsArray) {
      var personAsObj = {};

      headings.forEach(function(heading, i) {
        personAsObj[heading] = personAsArray[i];
      });

      return personAsObj;
    });
  }
  return peopleWithHeadings;
}

function objectToSheetValues(object, headers) {
  var arrayValues = new Array(headers.length);
  var lowerHeaders = headers.map(function(item) {
    return item.toLowerCase();
  });

  for (var item in object) {
    for (var header in lowerHeaders) {
      if (String(object[item].name) == String(lowerHeaders[header])) {
        arrayValues[header] = object[item].value;
        Logger.log(arrayValues);
      }
    }
  }
  return arrayValues;
}
function getRawDataFromSheet(sheetName) {
  var mSheet = ACTIVE_SPREADSHEET.getSheetByName(sheetName);
  if (mSheet)
    return mSheet.getSheetValues(
      1,
      1,
      mSheet.getLastRow(),
      mSheet.getLastColumn()
    );
}
