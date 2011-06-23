// Global variables
var ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = ss.getActiveSheet(),
    activeRange = ss.getActiveRange(),
    settings = {};
    headersRaw = getHeaders(sheet, activeRange,1);

// Add menu for MapBox functions
function onOpen() {
  var menuEntries = [ {name: "Export GeoJSON", functionName: "toGeoJSON"},
                      {name: "Help", functionName: "helpSite"} ];
  ss.addMenu("MapBox", menuEntries);
}

// Export selected range to GeoJSON
function toGeoJSON() {
  // Create a new UI
  var app = UiApp.createApplication().setTitle('Export GeoJSON').setStyleAttribute('width', '460');
  
  // Create a grid to hold the form
  var grid = app.createGrid(4, 2);
  
  // Add form elements to the grid
  grid.setWidget(0, 0, app.createLabel('ID:'));
  grid.setWidget(0, 1, app.createListBox().setName('idBox').setId('idBox')); 
  grid.setWidget(1, 0, app.createLabel('Longitude:'));
  grid.setWidget(1, 1, app.createListBox().setName('longBox').setId('longBox')); 
  grid.setWidget(2, 0, app.createLabel('Latitude:'));
  grid.setWidget(2, 1, app.createListBox().setName('latBox').setId('latBox')); 
  
  // Set the list boxes to the header values
  for (var i = 0; i < headersRaw.length; ++i) {
    app.getElementById("idBox").addItem(headersRaw[i]);
    app.getElementById("longBox").addItem(headersRaw[i]);
    app.getElementById("latBox").addItem(headersRaw[i]);
  }
  
  // Create a vertical panel...
  var panel = app.createVerticalPanel().setId('settingsPanel');

  var description = app.createLabel(
    'To format your spreadsheet as GeoJSON file, select the following columns:'
  ).setStyleAttribute('margin-bottom','20');
  panel.add(description);
  
  // ...and add the grid to the panel
  panel.add(grid);
  
  // Create a button and click handler; pass in the grid object as a 
  // callback element and the handler as a click handler
  // Identify the function b as the server click handler
  var button = app.createButton('Export')
                  .setStyleAttribute('margin-top','10')
                  .setId('export');
  var handler = app.createServerClickHandler('b');
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);
  
  // Add the button to the panel and the panel to the application, 
  // then display the application app in the Spreadsheet doc
  grid.setWidget(3, 1, button);
  app.add(panel);
  app.setStyleAttribute('padding','20');
  ss.show(app);
}

// Handle submits by updating the settings variable, calling the
// export function, closing the UI window
function b(e) {
  settings = {
    'id': e.parameter.idBox,
    'long': e.parameter.longBox,
    'lat': e.parameter.latBox
  };
  
  // Get the UI object
  var app = UiApp.getActiveApplication();
  
  // Update the settings button (this is not firing in before the call to exportGeoJSON())
  app.getElementById('export')
    .setText("Exporting...")
    .setEnabled(false);
  
  // Hide the settings panel
  var settingsPanel = app.getElementById('settingsPanel');
  settingsPanel.setVisible(false);
  
  // Create GeoJSON file and pass back it's filepath
  var filePath = exportGeoJSON();
  
  // Notify the user that the file is done and in their Google Docs list
  var description = app.createLabel('The GeoJSON file has been saved in your Google Docs List.')
                       .setStyleAttribute('margin-bottom','10');
  app.add(description);
  
  // And provide a link to it
  var fileLink = app.createAnchor('Download GeoJSON File',filePath)
                    .setStyleAttribute('font-size','150%');
  app.add(fileLink);
  
  // Update the UI.
  return app;
}

function getHeaders(sheet, range, columnHeadersRowIndex) {
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return headers;
}

function exportGeoJSON() {
  
  // Set up the response object
  var response = {
    "type": "FeatureCollection",
    "features": []
  };

  // For every row active range, make feature object
  var features = getRowsData(sheet, activeRange,1);
  response.features = features;

  // Output GeoJSON
  var outputFileName = normalizeHeader(ss.getName())+"-"+Date.now()+".geojson",
      stringToDisplay = JSON.stringify(response, null, 2),
      file = DocsList.createFile(outputFileName, stringToDisplay);
  return file.getUrl();
}


function helpSite() {
  Browser.msgBox('Support available here: https://github.com/mapbox/MapBox-for-Google-Docs');
}


/*
 * The following is taken from the Google Apps Script example for 
 * [reading docs](http://goo.gl/TigQZ). It's modified to build a 
 * [GeoJSON](http://geojson.org/) object.
 *
*/


// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  if (range.getRowIndex() == 1) {
    range = range.offset(1,0);
  }
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}


// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  var settingsIndex = [
    headersRaw.indexOf(settings.id),
    headersRaw.indexOf(settings.long),
    headersRaw.indexOf(settings.lat)
   ];
  
  // For each row
  for (var i = 0; i < data.length; ++i) {
    // If we have an id, long, and lat
    if (data[i][settingsIndex[0]] && data[i][settingsIndex[1]] && data[i][settingsIndex[2]]) {
      // Define a new GeoJSON feature object
      var feature = {
        "type": "Feature",
        // Get ID from UI
        "id": data[i][settingsIndex[0]],
        "geometry": {
          "type": 'Point',
          // Get coordinates from UIr
          "coordinates": [data[i][settingsIndex[1]], data[i][settingsIndex[2]]]
        },
        // Place holder for properties object
        "properties": {}
      };
      
      var hasData = false;
      
      // for each feild in row [i]
      for (var j = 0; j < data[i].length; ++j) {
        var cellData = data[i][j];
        if (isCellEmpty(cellData)) {
          continue;
            }
        // Populate the properties object with each feild
        feature.properties[keys[j]] = cellData;
        hasData = true;
      }
      if (hasData) {
        // Add feature to objects array
        objects.push(feature);
      }
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}




// Exercise:


function runExercise() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];

  // Get the range of cells that store employee data.
  var activeRange = sheet.getRange("B1:F5");

  // For every row of employee data, generate an employee object.
  var features = getColumnsData(sheet, activeRange);

  var thirdEmployee = features[2];
  var stringToDisplay = "The third column is: " + thirdEmployee.firstName + " " + thirdEmployee.lastName;
  stringToDisplay += " (id #" + thirdEmployee.employeeId + ") working in the ";
  stringToDisplay += thirdEmployee.department + " department and with phone number ";
  stringToDisplay += thirdEmployee.phoneNumber;
  ss.msgBox(stringToDisplay);
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}

// getColumnsData iterates column by column in the input range and returns an array of objects.
// Each object contains all the data for a given column, indexed by its normalized row name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - rowHeadersColumnIndex: specifies the column number where the row names are stored.
//       This argument is optional and it defaults to the column immediately left of the range; 
// Returns an Array of objects.
function getColumnsData(sheet, range, rowHeadersColumnIndex) {
  rowHeadersColumnIndex = rowHeadersColumnIndex || range.getColumnIndex() - 1;
  var headersTmp = sheet.getRange(range.getRow(), rowHeadersColumnIndex, range.getNumRows(), 1).getValues();
  var headers = normalizeHeaders(arrayTranspose(headersTmp)[0]);
  return getObjects(arrayTranspose(range.getValues()), headers);
}