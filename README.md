# Excel Maps

This application generates Excel maps from shape files.

## Usage

Use [topojson](https://github.com/mbostock/topojson) to convert your shape file into a topoJSON file. Then run:

    python shape.py path/to/topo.json

This creates `topo.xlsm` in the current directory with the map.

This takes a number of options that can be used independently:

    shape.py topo.json --out outfile.xlsm        # Saves to outfile.xlsm
    shape.py topo.json --col STATE,DISTRICT      # Only add STATE and DISTRICT columns to Excel
    shape.py topo.json --key STATE,DISTRICT      # Uses STATE:DISTRICT columns as key
    shape.py topo.json --filters ST=AP|TN,CN=IN  # Only draw features where CN is IN, and ST is AP or TN
    shape.py topo.json --show                    # Show Excel while drawing (slow, useful to debug)
    shape.py topo.json --enc cp1252              # Switch encoding of the TopoJSON file
    shape.py topo.json --license license-key     # Generate protected Excel file with specified license key

To display the properties, use:

    shape.py topo.json --prop prop.csv          # Saves all properties in prop.csv
    shape.py topo.json --prop -                 # Summarises properties on screen

```
python getshapefiles.py --help
usage: getshapefiles.py [-h] [-d DIRECTORY]

optional arguments:
  -h, --help            show this help message and exit
  -d DIRECTORY, --directory DIRECTORY
                        directory path inside which zipfiles should be
                        downloaded

 ```

The python script ```getshapefiles.py``` implements a method which sends the earlier created topojson files to external python script named ``shape.py`` which reads the topojson and draw the curves over excel sheets. It automatically saves excel maps which can be accessed any later time.
