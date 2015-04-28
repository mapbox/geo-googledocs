// Global variables
var ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = ss.getActiveSheet(),
    activeRange = ss.getActiveRange(),
    settings = {};

var geocoders = {
    yahoo: {
      query: function(query, key) {
        return 'http://where.yahooapis.com/geocode?appid=' +
          key + '&flags=JC&q=' + query;
      },
      parse: function(r) {
        try {
          return {
            longitude: r.ResultSet.Results[0].longitude,
            latitude: r.ResultSet.Results[0].latitude,
            accuracy: r.ResultSet.Results[0].quality
          }
        } catch(e) {
          return { longitude: '', latitude: '', accuracy: '' };
        }
      }
    },
    mapquest: {
      query: function(query, key) {
        return 'http://open.mapquestapi.com/nominatim/v1/search?format=json&limit=1&q=' + query;
      },
      parse: function(r) {
        try {
          return {
            longitude: r[0].lon,
            latitude: r[0].lat,
            accuracy: r[0].type
          }
        } catch(e) {
          return { longitude: '', latitude: '', accuracy: '' };
        }
      }
    },
    cicero: {
      query: function(query, key) {
        return 'https://cicero.azavea.com/v3.1/legislative_district?format=json&key=' +
          key + '&search_loc=' + query;
      },
      parse: function(r) {
        try {
          return {
            longitude: r.response.results.candidates[0].x,
            latitude: r.response.results.candidates[0].y,
            accuracy: r.response.results.candidates[0].score
          }
        } catch(e) {
          return { longitude: '', latitude: '', accuracy: '' };
        }
      }
    }
};


// Parts of following is taken from a Google Apps Script example for
// [reading docs](http://goo.gl/TigQZ). It's modified to build a
// [GeoJSON](http://geojson.org/) object.

// Add menu for Geo functions
function onOpen() {
  ss.addMenu('Geo', [{
      name: 'Geocode Addresses',
      functionName: 'gcDialog'
  }, {
      name: 'Export GeoJSON',
      functionName: 'gjDialog'
  }, {
      name: 'Help',
      functionName: 'helpSite'
  }]);
}

// UI to set up GeoJSON export
function gjDialog() {
  var headersRaw = getHeaders(sheet, activeRange, 1);

  // Create a new UI
  var app = UiApp.createApplication()
    .setTitle('Export GeoJSON')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Create a grid to hold the form
  var grid = app.createGrid(4, 2);

  // Add form elements to the grid
  grid.setWidget(0, 0, app.createLabel('Unique ID:'));
  grid.setWidget(0, 1, app.createListBox().setName('idBox').setId('idBox'));
  grid.setWidget(1, 0, app.createLabel('Longitude:'));
  grid.setWidget(1, 1, app.createListBox().setName('lonBox').setId('lonBox'));
  grid.setWidget(2, 0, app.createLabel('Latitude:'));
  grid.setWidget(2, 1, app.createListBox().setName('latBox').setId('latBox'));

  // Set the list boxes to the header values
  for (var i = 0; i < headersRaw.length; i++) {
    app.getElementById('idBox').addItem(headersRaw[i]);
    app.getElementById('lonBox').addItem(headersRaw[i]);
    app.getElementById('latBox').addItem(headersRaw[i]);
  }

  // Create a vertical panel...
  var panel = app.createVerticalPanel().setId('settingsPanel');

  panel.add(app.createLabel(
    'To format your spreadsheet as GeoJSON file, select the following columns:'
  ).setStyleAttribute('margin-bottom', '20'));

  // ...and add the grid to the panel
  panel.add(grid);

  // Create a button and click handler; pass in the grid object as a
  // callback element and the handler as a click handler
  // Identify the function b as the server click handler
  var button = app.createButton('Export')
      .setStyleAttribute('margin-top', '10')
      .setId('export');
  var handler = app.createServerClickHandler('exportGJ');
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);

  // Add the button to the panel and the panel to the application,
  // then display the application app in the spreadsheet doc
  grid.setWidget(3, 1, button);
  app.add(panel);
  ss.show(app);
}

// Handle submits by updating the settings object, calling the
// export function, updates the UI
function exportGJ(e) {
  settings = {
    id: e.parameter.idBox,
    lon: e.parameter.lonBox,
    lat: e.parameter.latBox
  };

  // Update ui to show status
  updateUi();

  // Create GeoJSON file and pass back it's filepath
  var file = createGJFile();

  // Update ui to deliver file
  displayFile(file);
}

function updateUi() {
  // Create a new UI instance
  var app = UiApp.createApplication()
    .setTitle('Export GeoJSON')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Add a status message to the UI
  app.add(app.createLabel(
    'Exporting your file...')
    .setStyleAttribute('margin-bottom', '10')
    .setId('exportingLabel'));

  // Show the new UI
  ss.show(app);
}

function displayFile(file) {
  // Create a new UI instance
  var app = UiApp.createApplication()
    .setTitle('Export GeoJSON')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Notify the user that the file is done and in the Google Docs list
  app.add(
    app.createLabel('The GeoJSON file has been saved in your Google Docs List.')
    .setStyleAttribute('margin-bottom', '10')
  );

  // And provide a link to it
  app.add(
    app.createAnchor('Download GeoJSON File', file.getUrl())
    .setStyleAttribute('font-size', '150%')
  );

  // Show the new UI
  ss.show(app);
}

// Get headers within a sheet and range
function getHeaders(sheet, range, columnHeadersRowIndex) {
    var numColumns = range.getEndColumn() - range.getColumn() + 1;
    var headersRange = sheet.getRange(columnHeadersRowIndex,
        range.getColumn(), 1, numColumns);
    return headersRange.getValues()[0];
}

// Create the GeoJSON file and returns its filepath
function createGJFile() {
    return DriveApp.createFile(
        (cleanCamel(ss.getName()) || 'unsaved') + '-' + Date.now() + '.geojson',
        Utilities.jsonStringify({
            type: 'FeatureCollection',
            features: getRowsData(sheet, activeRange, 1)
        })
    );
}

// Help menu
function helpSite() {
  Browser.msgBox('Support available here: https://github.com/mapbox/geo-googledocs');
}

// Geocoding UI to select API and enter key
function gcDialog() {
  // Create a new UI
  var app = UiApp.createApplication()
    .setTitle('Geocode Addresses')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Create a grid to hold the form
  var grid = app.createGrid(3, 2);

  // Add form elements to the grid
  grid.setWidget(0, 0, app.createLabel('Geocoding service:'));
  grid.setWidget(0, 1, app.createListBox()
    .setName('apiBox')
    .setId('apiBox')
    .addItem('mapquest')
    .addItem('yahoo')
    .addItem('cicero'));
  grid.setWidget(1, 0, app.createLabel('API key:'));
  grid.setWidget(1, 1, app.createTextBox().setName('keyBox').setId('keyBox'));

  // Create a vertical panel...
  var panel = app.createVerticalPanel().setId('geocodePanel');

  panel.add(app.createLabel(
    'The selected cells will be joined together and sent to a geocoding service. '
    +'New columns will be added for longitude, latitude, and accuracy score. '
    +'Select a geocoding API and enter your API key if required:'
  ).setStyleAttribute('margin-bottom', '20'));

  // ...and add the grid to the panel
  panel.add(grid);

  // Create a button and click handler; pass in the grid object as a
  // callback element and the handler as a click handler
  // Identify the function b as the server click handler
  var button = app.createButton('Geocode')
      .setStyleAttribute('margin-top', '10')
      .setId('geocode');
  var handler = app.createServerClickHandler('geocode');
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);

  // Add the button to the panel and the panel to the application,
  // then display the application app in the Spreadsheet doc
  grid.setWidget(2, 1, button);
  app.add(panel);
  ss.show(app);
}

// Geocode selected range with user-selected api and key
function geocode(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(),
      sheet = ss.getActiveSheet(),
      activeRange = ss.getActiveRange(),
      address = '',
      api = e.parameter.apiBox,
      key = e.parameter.keyBox,
      response = {},
      rowData = activeRange.getValues(),
      topRow = activeRange.getRow(),
      lastCol = activeRange.getLastColumn();


  // update UI
  updateUiGc();

  // Check to see if destination columns already exist

  var res = getDestCols();

  if (res.long >= 0 && res.lat  >= 0 && res.acc >= 0) {
    var longCol = (res.long+1),
        latCol = (res.lat+1),
        accCol = (res.acc+1);
  } else {
   // Add new columns
    sheet.insertColumnsAfter(lastCol, 3);

    // Set new column headers
    sheet.getRange(1, lastCol + 1, 1, 1).setValue('geo_longitude');
    sheet.getRange(1, lastCol + 2, 1, 1).setValue('geo_latitude');
    sheet.getRange(1, lastCol + 3, 1, 1).setValue('geo_accuracy');

    // Set destination columns
    var longCol = (lastCol + 1),
        latCol = (lastCol + 2),
        accCol = (lastCol + 3);
  }

  // Don't geocode the first row!
  if (activeRange.getRow() == 1) {
    rowData.shift();
    topRow = topRow + 1;
  }

  // For each row, query the API and update the spreadsheet
  for (var i = 0; i < rowData.length; i++) {
    // Join all fields in selected row with a space
    address = rowData[i].join(' ');

    // Concatenate all geo columns
    if (longCol && latCol&& accCol) {
      var testString = sheet.getRange(i + topRow, longCol, 1, 1).getValues()
          + sheet.getRange(i + topRow, latCol, 1, 1).getValues()
          + sheet.getRange(i + topRow, accCol, 1, 1).getValues();
    }
    // Test to see that all geo columns are empty
    Logger.log(testString);
    if (testString == '') {
      // Send address to query the geocoding api
      response = getApiResponse(address, api, key);

      // Add responses to columns in the active spreadsheet
      try {
        sheet.getRange(i + topRow, longCol, 1, 1).setValue(response.longitude);
        sheet.getRange(i + topRow, latCol, 1, 1).setValue(response.latitude);
        sheet.getRange(i + topRow, accCol, 1, 1).setValue(response.accuracy);
      } catch(e) {
        Logger.log(e);
      }
    }
  }

  // Update UI to notify user the geocoding is done
  closeUiGc();
}

// Check the spreadsheet to see if geo columns exist
function getDestCols() {
  // Get all headers of the active spreadsheet
  var headers = getHeaders(sheet, sheet.getRange(1,1,1,sheet.getLastRow()), 1);

  // Search through array for geo cols
  var output = {
    'long': include(headers,'geo_longitude'),
    'lat': include(headers,'geo_latitude'),
    'acc': include(headers,'geo_accuracy')
  };

  Logger.log(output.long);
  return output;
}

// Find item in array, return its index
function include(arr,obj) {
    Logger.log(arr.indexOf(obj));
    return arr.indexOf(obj);
}

// Update the UI to show geocoding status
function updateUiGc() {
  // Create new UI
  var app = UiApp.createApplication()
    .setTitle('Geocode Addresses')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Show working message
  app.add(app
    .createLabel('Geocoding these addresses...')
    .setStyleAttribute('margin-bottom', '10')
    .setId('geocodingLabel')
  );

  // Show the new ui
  ss.show(app);
}

// Update UI to show that geocoding is done
function closeUiGc() {
  Logger.log('starting updateUiGc');
  var app = UiApp.createApplication()
    .setTitle('Geocode Addresses')
    .setStyleAttribute('width', '460')
    .setStyleAttribute('padding', '20');

  // Exporting message
  app.add(app.createLabel(
    'Geocoding is done! You may close this window.')
    .setStyleAttribute('margin-bottom', '10')
    .setStyleAttribute('font-size', '150%')
    .setId('geocodingLabel'));

  ss.show(app);
}

// Send address to api
function getApiResponse(address, api, key) {
  var geocoder = geocoders[api],
      url = geocoder.query(encodeURI(address), encodeURI(key));

  // If the geocoder returns a response, parse it and return components
  // If the geocoder responds poorly or doesn't response, try again
  for (var i = 0; i < 5; i++) {
    try {
      var response = UrlFetchApp.fetch(url, {method:'get'});
    } catch(e) {
      Logger.log(e);
    }
    if (response && response.getResponseCode() == 200) {
      Logger.log(response.getResponseCode());
      return geocoder.parse(Utilities.jsonParse(response.getContentText()));
    } else {
      Logger.log('The geocoder service being used may be offline.');
    }
    // If no or bad response, sleep for 5 * i seconds and try again
    Logger.log('Something bad happened; retrying. Round: '+(i+1));
    for (var x = 0; x <= i; x++) {
      if (x < 3) { wait(5) };
      if (x = 3) { wait(60) };
      if (x = 4) { wait(120) };
    }
  }
  Logger.log('Tried 5 times, giving up.');
}

function wait(ms) {
  for (var i = 0; i < ms; i++) {
    Logger.log('Sleeping for '+(i+1)+' seconds.');
    Utilities.sleep(1000);
  }
}

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  if (range.getRowIndex() === 1) {
    range = range.offset(1, 0);
  }
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), headers.map(cleanCamel));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  var headers = getHeaders(sheet, activeRange, 1);

  // Zip an array of keys and an array of data into a single-level
  // object of `key[i]: data[i]`
  var zip = function(keys, data) {
    var obj = {};
    for (var i = 0; i < keys.length; i++) {
        obj[keys[i]] = data[i];
    }
    return obj;
  };

  // For each row
  for (var i = 0; i < data.length; i++) {
    var obj = zip(headers, data[i]);

    var lat = parseFloat(obj[settings.lat]),
      lon = parseFloat(obj[settings.lon]);

    var coordinates = (lat && lon) ? [lon, lat] : false;

    // If we have an id, lon, and lat
    if (obj[settings.id] && coordinates) {
      // Define a new GeoJSON feature object
      var feature = {
        type: 'Feature',
        // Get ID from UI
        id: obj[settings.id],
        geometry: {
          type: 'Point',
          // Get coordinates from UIr
          coordinates: coordinates
        },
        // Place holder for properties object
        properties: obj
      };
      objects.push(feature);
    }
  }
  return objects;
}

// Normalizes a string, by removing all non alphanumeric characters and using mixed case
// to separate words.
function cleanCamel(str) {
  return str
         .replace(/\s(.)/g, function($1) { return $1.toUpperCase(); })
         .replace(/\s/g, '')
         .replace(/[^\w]/g, '')
         .replace(/^(.)/, function($1) { return $1.toLowerCase(); });
}
