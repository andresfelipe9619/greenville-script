// import "google-apps-script";
let ACTIVE_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
let SELECTED_HOUSE = { data: null, row: null };

function onEdit(e) {
  let range = e.range;
  checkEditedCell(range);
}

function checkEditedCell(range) {
  const sheetName = range.getSheet().getName();
  if (sheetName !== "HOUSES") return;
  if (range.getColumn() == 10 && range.getValue() == "OK") {
    handleOnGenerateFields(range);
  } else if (range.getColumn() == 10 && range.getValue() != "OK") {
    console.log("NOTHING TO DO");
  }
}

function handleOnGenerateFields(range) {
  let sheet = ACTIVE_SPREADSHEET.getActiveSheet();
  let selectedRow = range.getRow();
  let nextRow = selectedRow + 1;

  let nextRowValues = sheet.getSheetValues(
    nextRow,
    1,
    1,
    sheet.getLastColumn()
  )[0];

  let rawHouse = sheet.getSheetValues(selectedRow, 1, 1, sheet.getLastColumn());
  let headers = sheet.getSheetValues(1, 1, 1, sheet.getLastColumn());
  let houseObj = sheetValuesToObject(rawHouse, headers)[0];
  console.log("HEADERS----->");
  console.log(headers);
  console.log("HOUSE OBJECT----->");
  console.log(houseObj);
  let isEmpty = isNextRowEmpty(nextRowValues);
  if (isEmpty) addIncrementalId({ sheet, houseObj, nextRow });

  SELECTED_HOUSE.data = houseObj;
  SELECTED_HOUSE.row = selectedRow;
  addHouseToSheets();
}

function addIncrementalId({ sheet, houseObj, nextRow }) {
  let nextId = houseObj["id house"] + 1;
  let nextIdRange = sheet.getRange(nextRow, 1, 1, 1);
  nextIdRange.setValues([[nextId + ""]]);
}

function isNextRowEmpty(nextRowValues) {
  let empties = nextRowValues.filter(function(value) {
    return String(value) == "" || String(value) == " ";
  });
  if (empties.length < 4) return false;
  return true;
}

function addHouseToSheets() {
  let sheets2NotSearch = ["HOUSES", "BALANCE", "LISTADOS"];
  let sheets = ACTIVE_SPREADSHEET.getSheets();
  let house = SELECTED_HOUSE.data;
  let houseIndex = null;
  for (let i in sheets) {
    let sheet = sheets[i];
    if (!sheets2NotSearch.includes(sheet.getName()) && !houseIndex) {
      let sheetName = sheet.getName();
      console.log(`--------------SHEET(${i}) NAME: ${sheetName}-------------`);
      const { index } = findText({ sheet, text: house["address"] });
      if (index) houseIndex = { index, sheetName };
      console.log("------------------------------------------------");
    }
  }
  if (houseIndex) return alreadyCreatedMessage({ house, ...houseIndex });
  let canCreate = true;
  for (let i in sheets) {
    let sheet = sheets[i];
    if (!sheets2NotSearch.includes(sheet.getName()) && canCreate) {
      canCreate = createHouse({ sheet, house });
    }
  }

  if (canCreate) {
    showMessage("Success", "House " + house["address"] + " copied to sheets");
  } else {
    showMessage(
      "Warning",
      `Can not copy house id ${house["id house"]} until id ${house["id house"] - 1} has been copied`
    );
  }
}

function findText({ sheet, text }) {
  let index = undefined;
  let textFinder = sheet.createTextFinder(text);
  let textFound = textFinder.findNext();
  if (textFound) index = textFound.getRow();
  let data = textFound || null;
  console.log("{data, index}", { data, index });
  return { index, data };
}

function getLastRow(sheet) {
  let columnAvalues = sheet.getRange("A1:A").getValues();
  let lastRow = columnAvalues.filter(String).length;
  return lastRow;
}

function createHouse({ sheet, house }) {
  const { index } = findText({ text: "ID HOUSE", sheet });
  let numberOfRowsBeforeData = index + 1;
  let lastRowWithValues = getLastRow(sheet);
  let rowToPlaceData = lastRowWithValues + numberOfRowsBeforeData;
  let backRowRangeId = sheet.getRange(rowToPlaceData - 1, 1, 1, 1);
  if (backRowRangeId.getValue() != house["id house"] - 1) return false;
  let range = sheet.getRange(rowToPlaceData, 1, 1, 2);
  range.setValues([[house["id house"], house["address"]]]);
  return true;
}

function alreadyCreatedMessage({ house, index, sheetName }) {
  showMessage(
    "Warning",
    `House ${house["address"]} already created on sheets. Row ${index} on ${sheetName}`
  );
}

function showErrorMessage(body) {
  showMessage("Error", body);
}

function showMessage(title, body) {
  console.log(title, body);
  Browser.msgBox(title, body, Browser.Buttons.OK);
}

function sheetValuesToObject(sheetValues, headers) {
  let headings =
    headers[0].map(v => v.toLowerCase()) ||
    sheetValues[0].map(v => v.toLowerCase());
  let people = sheetValues;
  if (sheetValues.length > 1) people = sheetValues.slice(1);

  let peopleWithHeadings = addHeadings(people, headings);

  function addHeadings(people, headings) {
    return people.map(function(personAsArray) {
      let personAsObj = {};

      headings.forEach(function(heading, i) {
        personAsObj[heading] = personAsArray[i];
      });

      return personAsObj;
    });
  }
  return peopleWithHeadings;
}

function objectToSheetValues(object, headers) {
  let arrayValues = new Array(headers.length);
  let lowerHeaders = headers.map(function(item) {
    return item.toLowerCase();
  });

  for (let item in object) {
    for (let header in lowerHeaders) {
      if (String(object[item].name) == String(lowerHeaders[header])) {
        arrayValues[header] = object[item].value;
        console.log(arrayValues);
      }
    }
  }
  return arrayValues;
}
function getRawDataFromSheet(sheetName) {
  let mSheet = ACTIVE_SPREADSHEET.getSheetByName(sheetName);
  if (mSheet)
    return mSheet.getSheetValues(
      1,
      1,
      mSheet.getLastRow(),
      mSheet.getLastColumn()
    );
}
