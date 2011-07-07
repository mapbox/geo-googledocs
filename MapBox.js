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
            if (r.ResultSet && r.ResultSet.Results && r.ResultSet.Results.length) {
                return {
                    longitude: r.ResultSet.Results[0].longitude,
                    latitude: r.ResultSet.Results[0].latitude,
                    accuracy: r.ResultSet.Results[0].quality
                }
            } else {
                return { longitude: '', latitude: '', accuracy: '' };
            }
        }
    }
};


// The following is taken from the Google Apps Script example for
// [reading docs](http://goo.gl/TigQZ). It's modified to build a
// [GeoJSON](http://geojson.org/) object.

// Add menu for Geo functions
function onOpen() {
  ss.addMenu('Geo', [{
      name: 'Export selected range to GeoJSON',
      functionName: 'gjDialog'
  }, {
      name: 'Geocode selected range',
      functionName: 'geocode'
  }, {
      name: 'Help',
      functionName: 'helpSite'
  }]);
}

// Export selected range to GeoJSON
function gjDialog() {
    var headersRaw = getHeaders(sheet, activeRange, 1);

    // Create a new UI
    var app = UiApp.createApplication()
      .setTitle('Export GeoJSON')
      .setStyleAttribute('width', '460');

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
    // then display the application app in the Spreadsheet doc
    grid.setWidget(3, 1, button);
    app.add(panel);
    app.setStyleAttribute('padding', '20');
    ss.show(app);
}

// Handle submits by updating the settings variable, calling the
// export function, closing the UI window
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
  Logger.log(file);
  
  return true;
}

function updateUi() {
  Logger.log('starting updateUi');
  var app = UiApp.createApplication()
    .setTitle('Export GeoJSON')
    .setStyleAttribute('width', '460');

  // Exporting message  
  app.add(app.createLabel(
    'Exporting your file...')
    .setStyleAttribute('margin-bottom', '10')
    .setId('exportingLabel'));

  app.setStyleAttribute('padding', '20');
  ss.show(app);
}

function displayFile(file) {
  Logger.log('starting displayFile');
  var app = UiApp.createApplication()
    .setTitle('Export GeoJSON')
    .setStyleAttribute('width', '460');
  
  // Notify the user that the file is done and in their Google Docs list
  app.add(app.createLabel(
      'The GeoJSON file has been saved in your Google Docs List.')
      .setStyleAttribute('margin-bottom', '10'));

  // And provide a link to it
  app.add(app.createAnchor('Download GeoJSON File', file.getUrl())
      .setStyleAttribute('font-size', '150%'));
  
  // Update the UI.
  app.setStyleAttribute('padding', '20');
  ss.show(app);
}

// Helper function to get headers within a sheet and range.
function getHeaders(sheet, range, columnHeadersRowIndex) {
    var numColumns = range.getEndColumn() - range.getColumn() + 1;
    var headersRange = sheet.getRange(columnHeadersRowIndex,
        range.getColumn(), 1, numColumns);
    return headersRange.getValues()[0];
}

function createGJFile() {
    return DocsList.createFile(
        (cleanCamel(ss.getName()) || 'unsaved') + '-' + Date.now() + '.geojson',
        JSON.stringify({
        type: 'FeatureCollection',
        features: getRowsData(sheet, activeRange, 1)
    }, null, 2));
}

function helpSite() {
    Browser.msgBox('Support available here: https://github.com/mapbox/geo-googledocs');
}

// crazy expirmental geocoding feature!

// UI to select API and enter key
function gcDialog() {

  // Create a new UI
  var app = UiApp.createApplication()
    .setTitle('Geocode Addresses')
    .setStyleAttribute('width', '460');

  // Create a grid to hold the form
  var grid = app.createGrid(3, 2);

  // Add form elements to the grid
  grid.setWidget(0, 0, app.createLabel('Geocoding service:'));
  grid.setWidget(0, 1, app.createListBox()
    .setName('apiBox')
    .setId('apiBox')
    .addItem('yahoo'));
  grid.setWidget(1, 0, app.createLabel('API key:'));
  grid.setWidget(1, 1, app.createTextBox().setName('keyBox').setId('keyBox'));

  // Create a vertical panel...
  var panel = app.createVerticalPanel().setId('geocodePanel');

  panel.add(app.createLabel(
    'The selected cells will be joined together and sent to a geocoding service. '
    +'New columns will be added for longitude, latitude, and accuracy score. '
    +'Select a geocoding API and enter your API key if required'
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
  app.setStyleAttribute('padding', '20');
  ss.show(app);
}

function geocode(e) {
    var address = '',
        api = 'yahoo',
        key = '0m1ivXjV34FJTFL7uW2WL5CbNIJrL14loXYnp2bqE3baaED9xpb_g2T9Puli2qhMdCUXtBbqPprTXqpa5d.o3Q--',
        response = {},
        rowData = activeRange.getValues(),
        topRow = activeRange.getRow();
        lastCol = activeRange.getLastColumn();
    
    // update UI
    var app = updateUiGc();
  
    // Add new columns
    sheet.insertColumnsAfter(lastCol, 3);
    sheet.getRange(1, lastCol + 1, 1, 1).setValue('longitude');
    sheet.getRange(1, lastCol + 2, 1, 1).setValue('latitude');
    sheet.getRange(1, lastCol + 3, 1, 1).setValue('accuracy');

    if (activeRange.getRow() == 1) {
        rowData.shift();
        topRow = topRow + 1;
    }
    for (var i = 0; i < rowData.length; i++) {
        address = rowData[i].join(' ');

        // Send address to geocoding service
        response = getApiResponse(address, api, key);

        // Add responses to columns in the active spreadsheet
        try {
            sheet.getRange(i + topRow, lastCol + 1, 1, 1).setValue(response.longitude);
            sheet.getRange(i + topRow, lastCol + 2, 1, 1).setValue(response.latitude);
            sheet.getRange(i + topRow, lastCol + 3, 1, 1).setValue(response.accuracy);
        } catch (e) {
            return null;
        }
    }
    closeUiGc();
}


function updateUiGc() {
  Logger.log('starting updateUiGc');
  var app = UiApp.createApplication()
    .setTitle('Geocode Addresses')
    .setStyleAttribute('width', '460');

  // Exporting message  
  app.add(app.createLabel(
    'Geocoding these addresses...')
    .setStyleAttribute('margin-bottom', '10')
    .setId('geocodingLabel'));

  app.setStyleAttribute('padding', '20');
  ss.show(app);
  return app;
}

function closeUiGc() {
  Logger.log('starting updateUiGc');
  var app = UiApp.createApplication()
    .setTitle('Geocode Addresses')
    .setStyleAttribute('width', '460');

  // Exporting message  
  app.add(app.createLabel(
    'Geocoding is done! You may close this window.')
    .setStyleAttribute('margin-bottom', '10')
    .setStyleAttribute('font-size', '150%')
    .setId('geocodingLabel'));

  app.setStyleAttribute('padding', '20');
  ss.show(app);
  return app;
}

function getApiResponse(address, api, key) {
    var geocoder = geocoders[api],
        url = geocoder.query(escape(address), key),
        response = UrlFetchApp.fetch(url, {
            method:'get'
        });

    if (response.getResponseCode() == 200) {
        return geocoder.parse(JSON.parse(response.getContentText()));
    } else {
        throw new Error('The geocoder service being used may be offline.');
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