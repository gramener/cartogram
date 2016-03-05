"""
Create an Excel map from topojson files.
"""
from __future__ import print_function

import io
import os
import sys
import glob
import math
import json
import win32com.client
from collections import Counter

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

# Keep count of how many times each shape key has been used
count = Counter()


def projection(lon, lat):
    """Albers: http://mathworld.wolfram.com/AlbersEqual-AreaConicProjection.html"""

    lon, lat = lon * math.pi / 180, lat * math.pi / 180
    # Origin of Cartesian coordinates
    phi0, lambda0 = 24 * math.pi / 180, 80 * math.pi / 180
    # Standard parallels
    phi1, phi2 = 8 * math.pi / 180, 37 * math.pi / 180

    n = .5 * (math.sin(phi1) + math.sin(phi2))
    theta = n * (lon - lambda0)
    c = math.cos(phi1) ** 2 + 2 * n * math.sin(phi1)
    rho = ((c - 2 * n * math.sin(lat)) / n) ** .5
    rho0 = ((c - 2 * n * math.sin(phi0)) / n) ** .5
    x, y = rho * math.sin(theta), rho0 - rho * math.cos(theta)
    return x, -y




def draw(base, topo, key):
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

    # The following are from trial and error, and work only for India
    vx = [px for pointlist in coords for px, py in pointlist]
    vy = [py for pointlist in coords for px, py in pointlist]
    minx, miny = min(vx), min(vy)
    maxx, maxy = max(vx), max(vy)
    dx, dy = maxx - minx, maxy - miny

    # We want the map in a 400x400 bounding box at 400, 20
    x0, y0 = 400, 20
    size = min(400 / dx, 400 / dy)

    for i, points in enumerate(coords):
        coords[i] = [(x0 + (px - minx) * size, y0 + (py - miny) * size) for px, py in points]

    for shape in topo['objects'].values():
        for geom in shape['geometries']:
            n_arcs = len(geom['arcs'])
            properties = geom.get('properties', {})
            name = key(properties)
            names = []

            # Convert arcs of a geometry into array of points
            for i, arcgroup in enumerate(geom['arcs']):
                # Consolidate shapes into a point list. TODO: factor in holes
                points = []
                for arc in arcgroup:
                    # arc is an index into point coords. Positive values go
                    # clockwise. Else, it's two's complement (~) goes anti-
                    # clockwise.
                    points += coords[arc] if arc >= 0 else coords[~arc][::-1]

                # Draw the points
                shape = base.Shapes.BuildFreeform(msoEditingAuto, *points[0])
                for point in points[1:]:
                    shape.AddNodes(msoSegmentLine, msoEditingAuto, *point)
                shape = shape.ConvertToShape()
                shape.Line.Weight = 0.25
                shape.Line.ForeColor.ObjectThemeColor = msoThemeColorBackground1

                shapename = shape.Name = '{:s}_{:d}'.format(name, count[name])
                names.append(shapename)
                count[name] += 1

            # Group shapes if required
            if n_arcs > 1:
                shape = base.Shapes.Range(names).Group()
                shape.Name = name or 'NA'

            yield properties


def main(args):
    # Launch Excel
    xl = win32com.client.Dispatch("Excel.Application")
    workbook = xl.Workbooks.Add()
    xl.Visible = msoTrue

    # In Excel 2007 / 2010, Excel files have multiple sheets. Remove all but first
    for sheet in range(1, len(workbook.Sheets)):
        workbook.Sheets[1].Delete()

    single_sheet = True
    propcol = {key: i for i, key in enumerate(args.key)}
    row = start_row = 4

    def key(properties):
        'Create a key by joining key columns from properties'
        return ':'.join(properties.get(k, '') for k in args.key).title()

    for pathspec in args.file:
        for filename in glob.glob(pathspec):
            print(filename)
            with io.open(filename, encoding=args.encoding) as handle:
                data = json.load(handle)
            if single_sheet:
                sheet = workbook.ActiveSheet
            else:
                sheet = workbook.Sheets.Add()
                row = start_row
            for prop in draw(sheet, data, key):
                sheet.Cells(row, 1).Value = 0
                for attr, val in prop.iteritems():
                    if attr not in propcol:
                        propcol[attr] = len(propcol)
                    sheet.Cells(row, propcol[attr] + 2).Value = val
                row += 1
                sys.stdout.write('.')

            sheet.Cells(start_row - 1, 1).Value = 'Value'
            for attr, column in propcol.iteritems():
                sheet.Cells(start_row - 1, column + 2).Value = attr

            if not single_sheet:
                sheet.Name = os.path.splitext(os.path.basename(filename))[0]

    # Add visual basic code. http://www.cpearson.com/excel/vbe.aspx
    # Requires Excel modification: http://support.microsoft.com/kb/282830
    # to resolve error 'Programmatic Access to Visual Basic Project is not trusted'
    vbproj = workbook.VBProject
    for sheet in workbook.Worksheets:
        codemod = vbproj.VBComponents(sheet.Name).CodeModule
        for line, row in enumerate(open('shape.bas')):
            codemod.InsertLines(line + 1, row)

        # Set the gradient
        sheet.Cells(1, 1).Value = 'Colors'
        sheet.Cells(1, 2).Value = 0.0
        sheet.Cells(1, 3).Value = 0.5
        sheet.Cells(1, 4).Value = 1.0
        sheet.Cells(1, 2).Interior.Color = 255      # Red
        sheet.Cells(1, 3).Interior.Color = 65535    # Yellow
        sheet.Cells(1, 4).Interior.Color = 5296274  # Green


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(
        description=__doc__.strip(),
        formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--key', nargs='*', default=[], help='Properties to be used as keys')
    parser.add_argument('--encoding', help='Input topojson encoding', default='utf-8')
    parser.add_argument('file', help='TopoJSON files', nargs='+')
    args = parser.parse_args()

    main(args)
