#!/usr/bin/env python
'''
Create an Excel map from topojson files.
'''
from __future__ import print_function, unicode_literals

import io
import os
import re
import math
import json
import yaml
import tornado.template
import win32com.client
from tqdm import tqdm
from collections import Counter, OrderedDict

# Define MS Office and Excel constants to make the code VB-like
msoEditingAuto = 0
msoSegmentLine = 0
msoFalse = 0
msoTrue = -1
ppLayoutBlank = 0xc
msoThemeColorText1 = 13
msoThemeColorBackground1 = 14
msoThemeColorBackground2 = 16
vbext_ct_StdModule = 1
xlOpenXMLWorkbookMacroEnabled = 52
xlLocationAsObject = 2

# Chart size and position
WIDTH, HEIGHT = 400, 400
LEFT, TOP = 400, 50
SIZE = {'width': 0, 'height': 0}            # Store the computed size of the last map

folder = os.path.dirname(os.path.abspath(__file__))


def rgb(r=0, g=0, b=0, r_factor=1, g_factor=256, b_factor=65536):
    return r * r_factor + g * g_factor + b * b_factor


# Map Colours
map_colors = [
    rgb(r=255), rgb(r=255, g=128),
    rgb(r=255, g=255),
    rgb(r=128, g=255),
    rgb(g=255),
]

# Keep count of how many times each shape key has been used
count = Counter()


def delete(path):
    if os.path.exists(path):
        os.unlink(path)


def projection(lon, lat):
    '''
    Albers: http://mathworld.wolfram.com/AlbersEqual-AreaConicProjection.html
    '''
    pi_in_degrees = 180
    deg2rad = math.pi / pi_in_degrees

    lon, lat = lon * deg2rad, lat * deg2rad
    # Origin of Cartesian coordinates
    india_lat, india_lon = 24, 80
    phi0, lambda0 = india_lat * deg2rad, india_lon * deg2rad
    # Standard parallels
    india_lat_min, india_lat_max = 8, 37
    phi1, phi2 = india_lat_min * deg2rad, india_lat_max * deg2rad

    n = .5 * (math.sin(phi1) + math.sin(phi2))
    theta = n * (lon - lambda0)
    c = math.cos(phi1) ** 2 + 2 * n * math.sin(phi1)
    rho = ((c - 2 * n * math.sin(lat)) / n) ** .5
    rho0 = ((c - 2 * n * math.sin(phi0)) / n) ** .5
    x, y = rho * math.sin(theta), rho0 - rho * math.cos(theta)
    return x, -y


def load_topojson(path, encoding='utf-8'):
    '''Loads a topojson file specified in the command line'''
    with io.open(path, encoding=encoding) as handle:
        data = json.load(handle, object_hook=OrderedDict)
    return data


def parse_filters(filters):
    '''
    Parses command line filters.
    Converts `X=a,Y=b|c` to `{'X': {'a'}, 'Y': {'b', 'c'}}`
    '''
    result = {}
    for item in filters.split(','):
        if '=' in item:
            parts = item.split('=', 2)
            result[parts[0]] = set(parts[1].split('|'))
    return result


def apply_filters(data, filters):
    '''
    Removes geometries not matching the command line filters.
    '''
    filters = parse_filters(filters)

    def cond(properties):
        '''Return True if properties meet condition'''
        for key, val in filters.items():
            if str(properties.get(key, '')) not in val:     # noqa: E911
                return False
        return True

    # Remove properties that do not match the filter
    used_arcs = set()
    for shape in data['objects'].values():
        shape['geometries'] = [
            geom for geom in shape['geometries']
            if cond(geom.get('properties', {}))
        ]
        # Identify used arcs
        for geom in shape['geometries']:
            for arcgroup in geom['arcs']:
                # TODO: factor in holes
                if isinstance(arcgroup[0], list):
                    arcgroup = arcgroup[0]
                used_arcs |= set(arcgroup)

    data['used_arcs'] = used_arcs


def add_cols(data, cols):
    unid_val = [0]
    keys = [k for k in args.key.split(',') if k]

    def key(properties):
        return ':'.join(str(properties.get(k, '')) for k in keys)   # noqa: E911

    def unid():
        unid_val[0] += 1
        return 'M%d' % unid_val[0]

    # Restrict to pre-defined columns, and add an ID column
    cols = set(col for col in cols if col)
    for shape in data['objects'].values():
        for geom in shape['geometries']:
            properties = geom.get('properties', {})
            if len(cols):
                result = {key: properties[key] for key in properties if key in cols}
            else:
                result = properties
            result['ID'] = key(properties) if len(keys) else unid()
            geom['properties'] = result


def draw(sheet, topo):
    '''
    Draw into a sheet
    the topo (JSON) object
    '''
    # Convert arcs into absolute positions
    sx, sy = topo['transform']['scale']
    tx, ty = topo['transform']['translate']

    coords = []
    for arc in topo['arcs']:
        # Convert into absolute integer coordinates
        x, y = arc[0]
        points = [(x, y)]
        for relative in arc[1:]:
            x, y = x + relative[0], y + relative[1]
            points.append((x, y))

        # Convert into lat-long, then project it
        coords.append([
            projection(px * sx + tx, py * sy + ty)
            for px, py in points
        ])

    # Get bounds used the used arcs, ignoring arcs unused by filters
    used_arcs = topo['used_arcs']
    vx = [px for arc, pointlist in enumerate(coords) for px, py in pointlist if arc in used_arcs]
    vy = [py for arc, pointlist in enumerate(coords) for px, py in pointlist if arc in used_arcs]
    minx, miny = min(vx), min(vy)
    maxx, maxy = max(vx), max(vy)
    dx, dy = maxx - minx, maxy - miny

    # We want the map centered in a WIDTH x HEIGHT bounding box from TOP, LEFT
    scale = min(WIDTH / dx, HEIGHT / dy)
    x0, y0 = LEFT, TOP
    SIZE['width'], SIZE['height'] = scale * dx, scale * dy

    for i, points in enumerate(coords):
        coords[i] = [(x0 + (px - minx) * scale, y0 + (py - miny) * scale)
                     for px, py in points]
    geoms = sum((shape['geometries'] for shape in topo['objects'].values()), [])
    map_color_index = 0

    for geom in tqdm(geoms):
        properties = geom['properties']
        name = properties['ID']
        names = []

        # Convert arcs of a geometry into array of points
        for i, arcgroup in enumerate(geom['arcs']):
            # Consolidate shapes into a point list. TODO: factor in holes
            points = []
            if isinstance(arcgroup[0], list):
                arcgroup = arcgroup[0]

            for arc in arcgroup:
                # arc is an index into point coords. +ve values go clockwise.
                # Else, it's two's complement (~) goes anti- clockwise.
                points += coords[arc] if arc >= 0 else coords[~arc][::-1]

            # Draw the points
            shape = sheet.Shapes.BuildFreeform(msoEditingAuto, *points[0])
            for point in points[1:]:
                shape.AddNodes(msoSegmentLine, msoEditingAuto, *point)
            shape = shape.ConvertToShape()
            shape.Line.Weight = 0.25
            shape.Line.ForeColor.ObjectThemeColor = msoThemeColorBackground1
            shape.Fill.ForeColor.RGB = map_colors[map_color_index]
            map_color_index = (map_color_index + 1) % len(map_colors)

            shapename = shape.Name = 'ID_{:d}'.format(count['ID'])
            names.append(shapename)
            count['ID'] += 1

        # Group shapes if it has more than 1 arc
        if len(geom['arcs']) > 1:
            shape = sheet.Shapes.Range(names).Group()

        shapename = shape.Name = name
        yield properties, shapename


def screenshot(sheet, img_file):
    '''Export all shapes on this sheet as an image'''
    xl = sheet.Application
    # Resize chart to picture size
    sheet.Shapes.SelectAll()
    xl.Selection.Copy()

    # Create chart as a canvas for saving this picture
    chart = xl.Charts.Add()
    chart = chart.Location(Where=xlLocationAsObject, Name=sheet.Name)
    chart_padding = {'width': 10, 'height': 30}     # Padding added by Excel charts
    chart.ChartArea.Width = WIDTH + chart_padding['width']
    chart.ChartArea.Height = HEIGHT + chart_padding['height']
    chart.Parent.Border.LineStyle = 0
    chart.ChartArea.ClearContents()
    chart.ChartArea.Select()
    chart.Paste()
    # Center the image
    xl.Selection.ShapeRange.IncrementLeft((WIDTH - SIZE['width']) / 2)
    xl.Selection.ShapeRange.IncrementTop((HEIGHT - SIZE['height']) / 2)

    # Save chart as image and delete it
    chart.Export(Filename=img_file)
    sheet.ChartObjects(1).Delete()


def main(xl, args):
    # Launch Excel
    workbook = xl.Workbooks.Add()

    if args.view:
        xl.Visible = msoTrue
    # output file defaults to the base name of the TopoJSON file
    if not args.out:
        args.out = os.path.splitext(args.topo)[0]

    # In Excel 2007 / 2010, Excel files have multiple sheets. Retain only first
    for sheet in range(1, len(workbook.Sheets)):
        workbook.Sheets[1].Delete()
    sheet = workbook.Sheets[0]

    propcol = {}

    data = load_topojson(args.topo, args.enc)
    apply_filters(data, args.filters)
    add_cols(data, args.col.split(','))

    row = start_row = 4
    props = []
    for prop, shapename in draw(sheet, data):
        sheet.Cells(row, 1).Value = 0
        sheet.Cells(row, 2).Value = shapename
        for attr, val in prop.items():
            if attr == 'ID':
                continue
            if attr not in propcol:
                propcol[attr] = len(propcol)
            sheet.Cells(row, propcol[attr] + 3).Value = val
            props.append(prop)
        row += 1

    sheet.Cells(start_row - 1, 1).Value = 'Value'
    sheet.Cells(start_row - 1, 2).Value = 'ID'
    for attr, column in propcol.items():
        sheet.Cells(start_row - 1, column + 3).Value = attr

    # Set the default gradient
    sheet.Cells(1, 1).Value = 'Colors'
    sheet.Cells(1, 2).Value = 0.0
    sheet.Cells(1, 3).Value = 0.5
    sheet.Cells(1, 4).Value = 1.0
    sheet.Cells(1, 2).Interior.Color = 255      # Red
    sheet.Cells(1, 3).Interior.Color = 65535    # Yellow
    sheet.Cells(1, 4).Interior.Color = 5296274  # Green

    # Take a screenshot
    filename = os.path.abspath(args.out + '.png')
    delete(filename)
    screenshot(sheet, filename)

    # Save CSV data
    if args.csv:
        import pandas as pd
        from six import StringIO, string_types

        info = {
            'Handle': os.path.split(args.out)[-1],
            'Title': args.out,
        }
        if args.attr:
            info.update(parse_filters(args.attr) if isinstance(args.attr, string_types)
                        else args.attr)
        buf = StringIO()
        pd.DataFrame(props).drop('ID', axis=1).to_html(buf, index=False, classes=None)
        table = re.sub(r'\s+', ' ', buf.getvalue())
        info['Body (HTML)'] = (info['Body (HTML)'] or '{table}').format(table=table, **info)
        # Keys starting with _ are ignored. These are just meant as formatting variables
        info = {key: val for key, val in info.items() if not key.startswith('_')}

        if os.path.exists(args.csv):
            # Load data, match columns and update / add record
            data = pd.read_csv(args.csv, encoding='utf-8').set_index('Handle')
            handle = info.pop('Handle')
            info.update({col: '' for col in data.columns if col not in info})
            data.loc[handle] = {key: val for key, val in info.items() if key in data.columns}
        else:
            data = pd.DataFrame([info]).set_index('Handle')
        data.to_csv(args.csv, encoding='utf-8')

    # Color all shapes in grey
    sheet.Shapes.SelectAll()
    xl.Selection.ShapeRange.Fill.ForeColor.RGB = rgb(r=224, g=224, b=224)

    # Add visual basic code. http://www.cpearson.com/excel/vbe.aspx
    # Requires Excel modification: http://support.microsoft.com/kb/282830
    # to resolve error 'Programmatic Access to Visual Basic Project is not
    # trusted'
    # Note: workbook.VBProject.VBComponents('Sheet1') works. But after renaming
    # the sheet, it still stays Sheet1. So rename sheet AFTER updating VBScript.
    vbscript = os.path.join(folder, 'shape.bas')
    with io.open(vbscript, encoding='utf-8') as handle:
        source = tornado.template.Template(handle.read()).generate(license=args.license)
        codemod = workbook.VBProject.VBComponents(sheet.Name).CodeModule
        for line, row in enumerate(source.decode('utf-8').split('\n')):
            codemod.InsertLines(line + 1, row)

    sheet.Name = os.path.split(args.out)[-1]
    filename = os.path.abspath(args.out + '.xlsm')
    delete(filename)
    print('Saving as', filename)
    workbook.SaveAs(filename, xlOpenXMLWorkbookMacroEnabled)
    workbook.Close()


def prop(args):
    import pandas as pd

    data = load_topojson(args.topo, args.enc)
    apply_filters(data, args.filters)
    add_cols(data, args.col.split(','))

    values = []
    for shape in data['objects'].values():
        for geom in shape['geometries']:
            values.append(geom.get('properties', {}))
    properties = pd.DataFrame(values)
    for col in properties.columns:
        top = properties[col].value_counts().head(5)
        print('%-16s %s' % (col, ', '.join(top.index.astype(str))[:60]))
    if args.prop == '-':
        print('')
        print(properties.head().to_string(index=False, justify='right'))
    else:
        properties.to_csv(args.prop, encoding='cp1252', index=False)
        print('Saved properties into', args.prop)


def batch(args):
    with io.open(args.yaml, encoding='utf-8') as handle:
        config = yaml.load(handle)
    common = config.get('common', {})
    for row in tqdm(config.get('maps', [])):
        arg = parser.parse_args([])
        for props in [common, row]:
            for key, val in props.items():
                if isinstance(val, dict):
                    original = getattr(arg, key, {}) or {}
                    original.update(val)
                    val = original
                setattr(arg, key, val)
        # If the generated file is newer than the topoJSON
        if os.path.exists(arg.out + '.xlsm'):
            if os.stat(arg.out + '.xlsm').st_mtime > os.stat(arg.topo).st_mtime:
                continue
        main(xl, arg)


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(
        description=__doc__.strip(),
        formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('-y', '--yaml', help='Load configuration from a YAML file')
    parser.add_argument('-t', '--topo', help='TopoJSON file')
    parser.add_argument('-o', '--out', help='File name to save .xlsm file as')
    parser.add_argument('-k', '--key', help='Columns to use as keys (comma-separated)', default='')
    parser.add_argument('-c', '--col', help='Columns to include (comma-separated)', default='')
    parser.add_argument('-f', '--filters', help='Filters (col=VAL,col=VAL,...)', default='')
    parser.add_argument('-l', '--license', help='License key for Excel')
    parser.add_argument('-v', '--view', help='View Excel while rendering', action='store_true')
    parser.add_argument('-p', '--prop', help='Save properties as CSV file (or "-" to print )')
    parser.add_argument('-e', '--enc', help='Topojson encoding', default='utf-8')
    parser.add_argument('--csv', help='Generate summary CSV file')
    parser.add_argument('-a', '--attr', help='CSV file attrs (col=VAL,col=VAL,...)', default='')
    args = parser.parse_args()
    if not args.topo and not args.yaml:
        parser.exit(status=2, message='One of --topo or --yaml is required\n')

    if args.prop:
        prop(args)
    else:
        xl = win32com.client.Dispatch('Excel.Application')
        try:
            if args.yaml:
                batch(args)
            else:
                main(xl, args)
        finally:
            xl.Quit()
