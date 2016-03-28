
# Creating Excel-Maps using Shape Files

This document describes how to generate excel-maps using shape files.

#### Points to highlight
* What is a Shape File 
* What is a TopoJSON file 
* How to create TopoJSON file
* How to create Excel-Maps using TopoJSON file

### What is a Shape File ?

The **shapefile** format is a digital vector storage format for storing geometric location and associated attribute information. This format lacks the capacity to store topological information. It is possible to read and write geographical datasets using the shapefile format with a wide variety of software.

The shapefile format is simple because it can store the primitive geometric data types of points, lines, and polygons. Shapes (points/lines/polygons) together with data attributes can create infinitely many representations about geographic data. Representation provides the ability for powerful and accurate computations.

The three mandatory files have filename extensions .shp, .shx, and .dbf. The actual shapefile relates specifically to the .shp file, but alone is incomplete for distribution as the other supporting files are required. Legacy GIS software may expect that the filename prefix be limited to eight characters to conform to the DOS 8.3 filename convention, though modern software applications accept files with longer names.


### What is a TopoJSON File ?
**TopoJSON**  is an extension of **GeoJSON**. **TopoJSON** introduces a new type of "Topology", that contains GeoJSON objects. A topology has an objects map which indexes geometry objects by name. These are standard GeoJSON objects, such as polygons, multi-polygons and geometry collections. However, the coordinates for these geometries are stored in the topology's arcs array, rather than on each object separately. An arc is a sequence of points, similar to a line string; the arcs are stitched together to form the geometry. Lastly, the topology has a transform which specifies how to convert delta-encoded integer coordinates to their native values (such as longitude & latitude).
Please follow the below link for more information on **TopoJSON** 

Introduction · mbostock/topojson Wiki ---   https://github.com/mbostock/topojson/wiki/Introduction

### How to create TopoJSON File
This code repository is creating topojson files using shapefiles. The file **getshapefiles** is a python script which downloads the shape files of world countries from **gadm** resource and converts them into **TopoJSON** format internally. As explained earlier **TopoJSON** is a topology to describe figures using geomerical objects.

The python script can easily be run using **python** command line interface, It will automatically **download**, **store** and **create** topojson files for the shape files.

### How to create Excel-Maps using TopoJSON files
The python script **getshapefiles** implements a method which sends the earlier created topojson files to external python script named **shape.py** which reads the topojson and draw the curves over excel sheets. It automatically saves excel maps which can be accessed any later time.




