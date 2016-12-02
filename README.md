# Excel Maps

This application generates Excel maps from shape files.

## Usage

Use [topojson 1.0](https://github.com/topojson/topojson-1.x-api-reference/blob/master/Command-Line-Reference.md)
to convert your shape file into a topoJSON file. For example:

    topojson -p -o topo.json input.shp

Then run:

    python shape.py path/to/topo.json

This creates `topo.xlsm` in the current directory with the map.

This takes a number of options that can be used independently:

    shape.py topo.json --out outfile            # Saves to outfile.xlsm
    shape.py topo.json --col STATE,DISTRICT     # Only add STATE and DISTRICT columns to Excel
    shape.py topo.json --key STATE,DISTRICT     # Uses STATE:DISTRICT columns as key
    shape.py topo.json --filters ST=AP|TN,C=IN  # Only draw features where C is IN, and ST is AP or TN
    shape.py topo.json --show                   # Show Excel while drawing (slow, useful to debug)
    shape.py topo.json --enc cp1252             # Switch encoding of the TopoJSON file
    shape.py topo.json --license license-key    # Generate protected Excel file with specified license key

To display the properties, use:

    shape.py topo.json --prop prop.csv          # Saves all properties in prop.csv
    shape.py topo.json --prop -                 # Summarises properties on screen

## Distribution

- Generate the Excel file using a command line script
- Open the generated Excel file
- Press Alt-F11 to go to the Visual Basic Editor
- On the left pane, right-click on VBAProject (for your filename) and select VBAProject Properties...
- Go to the Protection tab and
    - Check "Lock project for viewing"
    - Select a password
- Save the Excel sheet on share.gramener.com
- Note the command line script,


    python shape.py maps/india-districts.json --filter "STATE_NAME=ANDHRA PRADESH" --key DISTRICT --col DISTRICT --license W1KS51M11F6 --out Airtel-AP-TS
    python shape.py Hyderabad-PIN-codes.json --key id --out Hyderabad-PIN-codes

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
