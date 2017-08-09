# Excel Maps

This application generates Excel maps from TopoJSON files.

Use [mapshaper](https://github.com/mbloch/mapshaper) to convert shapefiles to
topojson:

    npm install -g mapshaper
    topojson -p -o topo.json input.shp

Then follow the [commnand line usage](#command-line-usage) or [batch usage](#batch-usage) below.

## Command Line Usage

Run:

    shape.py --topo path/to/topo.json --out topo

This creates `topo.xlsm` in the current directory with the map.

This takes a number of options that can be used independently:

    shape.py topo.json --col STATE,DISTRICT     # Only add STATE and DISTRICT columns to Excel
    shape.py topo.json --key STATE,DISTRICT     # Uses STATE:DISTRICT columns as key
    shape.py topo.json --filters ST=AP|TN,C=IN  # Only draw features where C is IN, and ST is AP or TN
    shape.py topo.json --view                   # View Excel while drawing (slow, useful to debug)
    shape.py topo.json --enc cp1252             # Switch encoding of the TopoJSON file
    shape.py topo.json --license license-key    # Generate protected Excel file with specified license key

To display the properties, use:

    shape.py topo.json --prop prop.csv          # Saves all properties in prop.csv
    shape.py topo.json --prop -                 # Summarises properties on screen

## Troubleshooting

- If you get a "Programmatic access to Visual Basic Project is not trusted"
  error, open Excel > File > Options > Trust Center > Trust Center Setttings >
  Macro Settings > Trust Access to the VBA Project object model.
  [Ref](https://stackoverflow.com/a/25638419/100904).

## Batch Usage

Create a `config.yaml` with this structure:

```yaml
# These config files consolidate multiple shape.py commands into a single file.
# The file has 2 sections: common and maps.
# The maps section is a list of command line arguments passed to shape.py.
# The common section has default arguments passed to all maps.
common:
    topo: maps/india-districts.json         # --topo=...
    key: ...                                # --key=...
    # The --csv option generates a CSV file with specified attributes
    # This is used to create output for Shopify uploads, for example.
    csv: out/shopify.csv
    # attr: holds columns added to the CSV file.
    # Keys beginning with "_" are ignored, but can be used as {template} variables.
    attr:
        _name: 'India'          # An internal column that will not be saved
        text: 'map of {_name}'  # Picks up {_name} from the _name attribute
        desc: '{table}'         # table is a pre-defined attr that holds the HTML table of all properties.
maps:
    -
        out: "out/India-Districts-2011-AP"      # Save the first map file here
        filters: "STATE_NAME=ANDHRA PRADESH"
        attr: {_name: Andhra Pradesh}
    -
        ...
```

Sample usage:

    shape.py -y config.yaml

## Notes

This application uses Windows automation. It "manually" opens Excel and draws the map. So:

- It cannot run on OSs other than Windows
- It requires MS Excel to be installed
- You mustn't copy-paste things while it's running. It uses the clipboard


## Protection

- Generate the Excel file using a `--license` key
- Open the generated Excel file
- Press Alt-F11 to go to the Visual Basic Editor
- On the left pane, right-click on VBAProject (for your filename) and select VBAProject Properties...
- Go to the Protection tab and
    - Check "Lock project for viewing"
    - Select a password
- Save the Excel sheet on share.gramener.com
- Note the command line script

Here is a sample usage of the `--license` key. Only users with a machine ID
matching `A1B2C3D4` will be able to open it. For others, it will pop up their
machine ID. You can copy-paste that and regenerate a license for their machine.

    python shape.py maps/india-districts.json --filter "STATE_NAME=ANDHRA PRADESH" --key DISTRICT --col DISTRICT --license A1B2C3D4 --out Airtel-AP-TS

## Sourcing shapefiles

`getshapefiles.py` downloads Shapefiles and converts them into topojson. To be documented.


    python getshapefiles.py --help
    usage: getshapefiles.py [-h] [-d DIRECTORY]

    optional arguments:
      -h, --help            show this help message and exit
      -d DIRECTORY, --directory DIRECTORY
                            directory path inside which zipfiles should be
                            downloaded
