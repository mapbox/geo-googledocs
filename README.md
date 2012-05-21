# Geo for Google Docs

[Read the blog post & instructions for using this system](http://mapbox.com/blog/mapping-google-doc-spreadsheet/).

## A [MapBox](http://mapbox.com/) Project

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
  [Yahoo PlaceFinder](http://developer.yahoo.com/geo/placefinder/) or [MapQuest Nominatim](http://developer.mapquest.com/web/products/open/nominatim). Consult these services for their terms of use.

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
