# Excel Maps

This application generates Excel maps from TopoJSON files.

Use [mapshaper](https://github.com/mbloch/mapshaper) to convert shapefiles to
topojson:

    npm install -g mapshaper
    mapshaper-gui

On the GUI, import your map. Then export as TopoJSON.

Now follow the [command line usage](#command-line-usage) or [batch usage](#batch-usage) below.

## Command Line Usage

Clone this repository. Then, from that directory, run:

    python shape.py -t path/to/topo.json --out output

This creates `output.xlsm` in the current directory with the map.

This takes a number of options that can be used independently:

    python shape.py -t topo.json --col STATE,DISTRICT     # Only add STATE and DISTRICT columns to Excel
    python shape.py -t topo.json --key STATE,DISTRICT     # Uses STATE:DISTRICT columns as key
    python shape.py -t topo.json --filters ST=AP|TN,C=IN  # Only draw features where C is IN, and ST is AP or TN
    python shape.py -t topo.json --view                   # View Excel while drawing (slow, useful to debug)
    python shape.py -t topo.json --enc cp1252             # Switch encoding of the TopoJSON file
    python shape.py -t topo.json --license license-key    # Generate protected Excel file with specified license key

If you don't know the columns (called properties) in the JSON file, use:

    python shape.py -t topo.json --prop prop.csv          # Saves all columns in prop.csv
    python shape.py -t topo.json --prop -                 # Summarises columns on screen

## Troubleshooting

- If you get a "Programmatic access to Visual Basic Project is not trusted"
  error, open Excel > File > Options > Trust Center > Trust Center Setttings >
  Macro Settings > Trust Access to the VBA Project object model.
  [Ref](https://stackoverflow.com/a/25638419/100904).

## Notes

This application uses Windows' COM for automation. It "manually" opens Excel and
draws the map. So:

- It only runs Windows
- MS Excel must be installed
- You mustn't copy-paste things while it's running. It uses the clipboard

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

    python shape.py -y config.yaml

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
