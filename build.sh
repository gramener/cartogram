# Convert shape files into other formats

SOURCE=/d/site/gramener.com/viz/maps/data/
TOPOJSON=D:/node_modules/topojson/bin/topojson

# Huge SHP files require more memory from node.js
# OPTIONS=--max_old_space_size=900

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
        node $OPTIONS $TOPOJSON \
            -p ST_CODE \
            -p PC_NO \
            -p ST_NAME \
            -p PC_NAME \
            -p PC_TYPE \
            -p AREA \
            --simplify-proportion 0.15 \
            --quantization 10000 \
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
        node $OPTIONS $TOPOJSON \
            -p ST_CODE \
            -p PC_NO \
            -p AC_NO \
            -p AC_NAME \
            -p AC_TYPE \
            -p AREA \
            --simplify-proportion 0.15 \
            --quantization 10000 \
            --out $OUT \
            $SRC
    done
}

# Create topojson files for administrative boundaries under maps/
function topojson_adm {
    for shape in `find $SOURCE -name 'IND_adm*.shp'`
    do
        SRC=`cygpath -aw $shape`
        OUT="maps/`basename $shape`"
        OUT="${OUT%.*}".json
        node $OPTIONS $TOPOJSON \
            -p \
            --simplify-proportion 0.15 \
            --quantization 10000 \
            --out $OUT \
            $SRC
    done
}

# Create CSV metadata file for all topojson files under maps/
function topojson_meta {
    python <<EOF
import os
import csv
import json
import glob

state_name = {}
properties = []
for filename in glob.glob('maps/*_PC.json') + glob.glob('maps/*_AC.json'):
    data = json.load(open(filename))

    for g in data['objects'].values()[0]['geometries']:
        properties.append(g['properties'])
        properties[-1]['filename'] = os.path.split(filename)[-1]

    # State names are not always present. Create a lookup table.
    p = g['properties']
    if 'ST_NAME' in p:
        state_name[p['ST_CODE']] = p['ST_NAME']


fields = ['filename', 'ST_CODE', 'PC_NO', 'AC_NO', 'PC_TYPE', 'PC_NAME', 'AC_TYPE', 'AC_NAME', 'ST_NAME', 'AREA']
with open('maps/metadata.csv', 'w') as fp:
    out = csv.DictWriter(fp, fields, lineterminator='\n')
    out.writerow({field:field for field in fields})
    for p in properties:
        p['ST_NAME'] = state_name[p['ST_CODE']]
        out.writerow(p)

EOF
}


topojson_pc
topojson_ac
topojson_adm
topojson_meta
