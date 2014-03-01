# Geo for Google Docs

Geo for Google Docs is a set of tools that make it easy to use data from
Google Docs Spreadsheets in [TileMill](http://tilemill.com), an open source
map design studio.

## Uses

- **Export spreadsheet data to [GeoJSON](http://geojson.org/)**
  Any spreadsheet with geocodes (longitude and latitude
  coordinates) can be exported to a TileMill-ready
  file. After exporting, just copy the GeoJSON file to your
  `TileMill/files/data` directory. 
- **Geocode arbitrary addresses** If your spreadsheet does not have
  geocodes, you can add them using a geocoding service like those provided by
  [Yahoo PlaceFinder](http://developer.yahoo.com/geo/placefinder/) or [MapQuest Nominatim](http://developer.mapquest.com/web/products/open/nominatim) or [Cicero API](https://cicero.azavea.com/docs/). 
  Consult these services for their terms of use.

## Installation

- Copy the [source of mapbox.js](https://raw.github.com/mapbox/geo-googledocs/master/MapBox.js)
- Open your spreadsheet and goto `tools` > `script editor`
- Replace the content in the text box with the copied source from mapbox.js
- Set the name of this script to `geo`
- Go to `file` > `save` and close the popup window
- Refresh your spreadsheet and you will see a new menu called `Geo` added after `Help` on the menu bar

## Authors

* Dave Cole [[dhcole](https://github.com/dhcole)]
* Tom MacWright [[tmcw](https://github.com/tmcw)]

------

## Usage walkthrough

Geo for Google Docs is an add-on script for Google Docs Spreadsheets that lets you geocode arbitrary addresses and export spreadsheets as GeoJSON.

## Getting set up

First [copy the source code](https://raw.github.com/mapbox/geo-googledocs/master/MapBox.js) for the Geo for Google Docs script.

Next, open your spreadsheet in Google Docs and select `Tools` > `Script Editor`. Erase the default `function myFunction() {}` in the code window, and paste in the source code you copied. Name the project 'Geo', `File` > `Save`, and close the window. Refresh your spreadsheet and you will see a new `Geo` menu next to `Help`.

Now you're ready to use the Geo add-on. It has two simple features: geocode addresses and export as GeoJSON.

## Geocoding arbitrary addresses

Say you've got a spreadsheet with a column for addresses. It could be a street address, a ZIP Code, a county or a state name - even a country code. To map this you need to find the geocodes, or longitude and latitude coordinates for these addresses. For locations that are more general, like "Canada", geocoding APIs usually return the coordinates of the centroid - the area's center point - or capital.

After you've installed the Geo script in your spreadsheet, you can use the [MapQuest Nominatim](http://developer.mapquest.com/web/products/open/nominatim) or [Yahoo PlaceFinder](http://developer.yahoo.com/geo/placefinder/) services to get geocodes for your addresses. First select the cells with your address data. This could be one column or several. If you select multiple columns (such as 'Address', 'City', 'State') the columns with be joined together from left to right and sent to the geocoding service. You can also select just a few cells to geocode instead of a whole column. This is useful when you add new rows or update existing addresses and only want to geocode the new data.

Once you've selected your address data, go to `Geo` > `Geocode Addresses`. If this is the first time you're using the script, you'll need to authorize it to access your spreadsheet and then go to `Geocode Addresses` again. Now all you need to do is set your preferred API (and enter an API key for Yahoo - an API key is not required for MapQuest), and click `Geocode`.

![Geo menu with data in it](https://farm7.static.flickr.com/6044/6237508379_58021d70c5.jpg)

You should now see three new columns added to your spreadsheet and coordinate data being entered into those. The third column is a response back from the geocoding service about the resolution of the geocode - something like 'residential' for street address, 'administrative' for state level match if your using MapQuest, or a numeric score if you're using Yahoo. This is particularly useful for checking the accuracy of your geocodes. For instance if you submit a country name and get back a 'residential' response, its likely that the geocoder could not figure out what to do with that address and that you should try to make it more specific, perhaps by using a three-digit code like 'USA' instead of 'United States'. If you find your geocodes are still inaccurate or slightly off, you can manually enter coordinates. Try using a service like [Get Lat Lon](http://www.getlatlon.com/). Also, we do not recommend you use this script for more than 1,000 geocodes at a time. For large datasets, [try Google Refine](http://support.mapbox.com/kb/tilemill/converting-addresses-in-spreadsheets-to-custom-maps-in-tilemill), which we have successfully used for up to 30,000 addresses.

Many web services exist for geocoding addresses. We use MapQuest's Nominatum search most often, as it is based on Open Street Map data and does not have a defined rate limit. We also use Yahoo PlaceFinder occasionally because it too has a generous rate limit and terms of service. We tend to avoid geocoders with expensive licensing, like the Google Maps API. Before using any geocoder, make sure to review its terms of service. If you'd like to add a new geocoder to this script, please [fork it on GitHub](https://github.com/mapbox/geo-googledocs) and submit a pull request. General issues should be [reported on GitHub](https://github.com/mapbox/geo-googledocs/issues) as well.

Now that your spreadsheet has geospatial data (coordinates), it's ready to be mapped.

## Exporting to GeoJSON

To load your spreadsheet data into [TileMill](http://mapbox.com/tilemill) or another GIS application like [Quantum GIS](http://www.qgis.org/), save the data as a [GeoJSON](http://geojson.org/) file. To do this with the Geo script, select all the data you want to export and go to `Geo` > `Export GeoJSON`. Select a column with unique content for each row, and specify the columns with your coordinates. Then click export and download the file. You'll need to add `.geojson` to the end of the filename so that TileMill recognizes it as GeoJSON.

To use this in TileMill, copy the file to your [Mapbox data directory](http://mapbox.com/tilemill/docs/manual/files-directories/). Then open TileMill, add a new layer, and specify this file. Now you're ready to start designing your map. For more on that, see the [TileMill Manual](http://mapbox.com/tilemill/docs/manual/) and [Support Forum](http://support.mapbox.com/discussions/tilemill).
