# Convert shape files into other formats

SOURCE=/d/site/gramener.com/viz/maps/data/india-constituencies/
TOPOJSON=D:/node_modules/topojson/bin/topojson

# The below paths are no longer used.
# OGR2OGR=/c/Program\ Files\ \(x86\)/Quantum\ GIS\ Lisboa/bin/ogr2ogr.exe
# GDAL_DATA='C:\Program Files (x86)\Quantum GIS Lisboa\share\gdal'
# INKSCAPE=/d/Apps/Inkscape/App/Inkscape/inkscape.com

# Create topojson files for Parliamentary constituences under maps/
function topojson_pc {
    for shape in `find $SOURCE -name '*_PC.shp'`
    do
        SRC=`cygpath -aw $shape`
        OUT="maps/`basename $shape`"
        OUT="${OUT%.*}".json
        node $TOPOJSON \
            -p ST_CODE \
            -p PC_NO \
            -p ST_NAME \
            -p PC_NAME \
            -p PC_TYPE \
            -p AREA \
            --simplify-proportion 0.15 \
            --quantization 2000 \
            --out $OUT \
            $SRC
    done
}

# Create topojson files for Assembly constituences under maps/
function topojson_ac {
    for shape in `find $SOURCE -name '*_AC.shp'`
    do
        SRC=`cygpath -aw $shape`
        OUT="maps/`basename $shape`"
        OUT="${OUT%.*}".json
        node $TOPOJSON \
            -p ST_CODE \
            -p PC_NO \
            -p AC_NO \
            -p AC_NAME \
            -p AC_TYPE \
            -p AREA \
            --simplify-proportion 0.15 \
            --quantization 2000 \
            --out $OUT \
            $SRC
    done
}

# Create CSV metadata file for all topojson files under maps/
function topojson_meta {
    python <<EOF
import os
import json
import glob

file_state = {}
state_name = {}
for filename in glob.glob('maps/*_PC.json') + glob.glob('maps/*_AC.json'):
    data = json.load(open(filename))
    p = data['objects'].values()[0]['geometries'][0]['properties']
    if 'ST_NAME' in p:
        state_name[p['ST_CODE']] = p['ST_NAME']
    file_state[os.path.split(filename)[-1]] = p['ST_CODE']

for filename in file_state:
    file_state[filename] = state_name[file_state[filename]]

with open('maps/metadata.json', 'w') as fp:
    fp.write(json.dumps(sorted(file_state.items())))

EOF
}


topojson_pc
topojson_ac
topojson_meta
