'''
Create an Excel map from topojson files.
'''
from __future__ import print_function, unicode_literals

import io
import os
import sys
import glob
import math
import json
import logging
import tornado.template
import win32com.client
from tqdm import tqdm
from collections import Counter, OrderedDict

# Define MS Office and Excel constants
msoEditingAuto = 0x0
msoSegmentLine = 0x0
msoFalse = 0
msoTrue = -1
ppLayoutBlank = 0xc
msoThemeColorText1 = 13
msoThemeColorBackground1 = 14
msoThemeColorBackground2 = 16
vbext_ct_StdModule = 1
xlOpenXMLWorkbookMacroEnabled = 52

# Keep count of how many times each shape key has been used
count = Counter()


def projection(lon, lat):
    '''
    Albers: http://mathworld.wolfram.com/AlbersEqual-AreaConicProjection.html
    '''
    deg2rad = math.pi / 180

    lon, lat = lon * deg2rad, lat * deg2rad
    # Origin of Cartesian coordinates
    phi0, lambda0 = 24 * deg2rad, 80 * deg2rad
    # Standard parallels
    phi1, phi2 = 8 * deg2rad, 37 * deg2rad

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
            if str(properties.get(key, '')) not in val:
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


def add_cols(data, cols, key=None):
    keys = [key for key in args.key.split(',') if key]
    unid_val = [0]

    def key(properties):
        return ':'.join(str(properties.get(k, '')) for k in keys)

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

    # We want the map in a 400x400 bounding box at 400, 20
    width = 400
    x0, y0 = width, 20
    size = min(width / dx, width / dy)

    for i, points in enumerate(coords):
        coords[i] = [(x0 + (px - minx) * size, y0 + (py - miny) * size)
                     for px, py in points]
    geoms = sum((shape['geometries'] for shape in topo['objects'].values()), [])

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

            shapename = shape.Name = 'ID_{:d}'.format(count['ID'])
            names.append(shapename)
            count['ID'] += 1

        # Group shapes if it has more than 1 arc
        if len(geom['arcs']) > 1:
            shape = sheet.Shapes.Range(names).Group()

        shapename = shape.Name = name
        yield properties, shapename


def main(args):
    # Launch Excel
    xl = win32com.client.Dispatch('Excel.Application')
    workbook = xl.Workbooks.Add()

    if args.show:
        xl.Visible = msoTrue

    # In Excel 2007 / 2010, Excel files have multiple sheets. Retain only first
    for sheet in range(1, len(workbook.Sheets)):
        workbook.Sheets[1].Delete()
    sheet = workbook.Sheets[0]

    propcol = {}

    data = load_topojson(args.file, args.enc)
    apply_filters(data, args.filters)
    add_cols(data, args.col.split(','))

    row = start_row = 4
    for prop, shapename in draw(sheet, data):
        sheet.Cells(row, 1).Value = 0
        sheet.Cells(row, 2).Value = shapename
        for attr, val in prop.items():
            if attr == 'ID':
                continue
            if attr not in propcol:
                propcol[attr] = len(propcol)
            sheet.Cells(row, propcol[attr] + 3).Value = val
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

    # Add visual basic code. http://www.cpearson.com/excel/vbe.aspx
    # Requires Excel modification: http://support.microsoft.com/kb/282830
    # to resolve error 'Programmatic Access to Visual Basic Project is not
    # trusted'
    # Note: workbook.VBProject.VBComponents('Sheet1') works. But after renaming
    # the sheet, it still stays Sheet1. So rename sheet AFTER updating VBScript.
    vbscript = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'shape.bas')
    with io.open(vbscript, encoding='utf-8') as handle:
        source = tornado.template.Template(handle.read()).generate(license=args.license)
        codemod = workbook.VBProject.VBComponents(sheet.Name).CodeModule
        for line, row in enumerate(source.decode('utf-8').split('\n')):
            codemod.InsertLines(line + 1, row)

    sheet.Name = os.path.splitext(os.path.basename(args.file))[0]
    filename = os.path.abspath(sheet.Name + '.xlsm')
    if os.path.exists(filename):
        os.unlink(filename)
    print('Saving as', filename)
    workbook.SaveAs(filename, xlOpenXMLWorkbookMacroEnabled)
    workbook.Close()
    xl.Quit()


def prop(args):
    import pandas as pd

    data = load_topojson(args.file, args.enc)
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
        print()
        print(properties.head().to_string(index=False, justify='right'))
    else:
        properties.to_csv(args.prop, encoding='cp1252', index=False)
        print('Saved properties into', args.prop)


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(
        description=__doc__.strip(),
        formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('-o', '--out', help='File name to save .xlsm file as')
    parser.add_argument('-k', '--key', help='Columns to use as keys (comma-separated)', default='')
    parser.add_argument('-c', '--col', help='Columns to include (comma-separated)', default='')
    parser.add_argument('-f', '--filters', help='Filters (col=VALUE,col=VALUE,...)', default='')
    parser.add_argument('-l', '--license', help='License key for Excel')
    parser.add_argument('-s', '--show', help='Show Excel while rendering', action='store_true')
    parser.add_argument('--prop', help='Save properties as CSV file (or "-" to print )')
    parser.add_argument('-e', '--enc', help='Topojson encoding', default='utf-8')
    parser.add_argument('file', help='TopoJSON file')
    args = parser.parse_args()

    if args.prop:
        prop(args)
    else:
        main(args)
