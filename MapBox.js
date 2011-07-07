// Global variables
var ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = ss.getActiveSheet(),
    activeRange = ss.getActiveRange(),
    settings = {};

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

    // Get the UI object
    var app = UiApp.getActiveApplication();

    // Update the settings button (this is not firing in before the call to exportGeoJSON())
    app.getElementById('export')
      .setText('Exporting...')
      .setEnabled(false);

    // Hide the settings panel
    var settingsPanel = app.getElementById('settingsPanel');
    settingsPanel.setVisible(false);

    // Create GeoJSON file and pass back it's filepath
    var file = createGJFile();

    // Notify the user that the file is done and in their Google Docs list
    app.add(app.createLabel(
        'The GeoJSON file has been saved in your Google Docs List.')
        .setStyleAttribute('margin-bottom', '10'));

    // And provide a link to it
    app.add(app.createAnchor('Download GeoJSON File', file.getUrl())
        .setStyleAttribute('font-size', '150%'));

    // Update the UI.
    return app;
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

function geocode() {
  var address = '',
      api = 'yahoo',
      key = '0m1ivXjV34FJTFL7uW2WL5CbNIJrL14loXYnp2bqE3baaED9xpb_g2T9Puli2qhMdCUXtBbqPprTXqpa5d.o3Q--',
      response = {},
      rowData = activeRange.getValues(),
      topRow = activeRange.getRow();
  
  if (activeRange.getRow() == 1) {
    rowData.shift();
    topRow = topRow + 1;
  }
  for (var i = 0; i < rowData.length; i++) {
    var addressParts = [];
    // Get the current address
    for (var x = 0; x < rowData[i].length; x++) {
      addressParts.push(rowData[i][x]);
    }
    address = addressParts.join(' ');
    
    // Send address to geocoding service
    response = getApiResponse(address, api, key);
    // Add responses to columns in the active spreadsheet
    try {
      sheet.getRange(i+topRow, 6, 1, 1).setValue(response.longitude);
      sheet.getRange(i+topRow, 7, 1, 1).setValue(response.latitude);
      sheet.getRange(i+topRow, 8, 1, 1).setValue(response.accuracy);
      sheet.getRange(i+topRow, 9, 1, 1).setValue(address);
    } catch(e) {
      return null;
    }
  }
}

function getApiResponse(address, api, key) {
  switch(api) {
    case 'yahoo':
    apiModel = {
      baseUrl: 'http://where.yahooapis.com/geocode?appid='+key+'&flags=JC&q=',
      longitude: 'ResultSet.Results[0].longitude',
      latitude: 'ResultSet.Results[0].latitude',
      accuracy: 'ResultSet.Results[0].quality',
    }
    break;
    //TODO add more geocoders
  }
  var url = apiModel.baseUrl + escape(address),
      response = UrlFetchApp.fetch(url, {
        method:'get'
      }),
      responseArray = [];

  if (response.getResponseCode() == 200) {
    var responseObject = JSON.parse(response.getContentText());
    responseArray = responseObject["ResultSet"]["Results"];
    try {
      Logger.log('longitude: ' + eval('responseObject.'+apiModel.longitude));
      Logger.log('latitude: ' + eval('responseObject.'+apiModel.latitude));
      Logger.log('accuracy: ' + eval('responseObject.'+apiModel.accuracy));
    } catch (e) {
      return null;
    }
    var output = {
      longitude: eval('responseObject.'+apiModel.longitude),
      latitude: eval('responseObject.'+apiModel.latitude),
      accuracy: eval('responseObject.'+apiModel.accuracy)
    };
    Logger.log(output.longitude);
    return output;
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

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words.
function cleanCamel(str) {
    return str
        .replace(/\s(.)/g, function($1) { return $1.toUpperCase(); })
        .replace(/\s/g, '')
        .replace(/[^\w]/g, '')
        .replace(/^(.)/, function($1) { return $1.toLowerCase(); });
}
