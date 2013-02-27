"""
Create an Excel map from topojson files.

Usage:

    python shape.py file1.json file2.json ...

TODO:
- Account for shapes with holes (Outer Manipur in S14_PC.json)
"""

import sys
import glob
import math
import json
import win32com.client

msoEditingAuto = 0x0
msoSegmentLine = 0x0
msoFalse = 0
msoTrue = -1
ppLayoutBlank = 0xc
msoThemeColorText1 = 13
msoThemeColorBackground1 = 14
msoThemeColorBackground2 = 16
vbext_ct_StdModule = 1


def projection(lon, lat):
    """Albers: http://mathworld.wolfram.com/AlbersEqual-AreaConicProjection.html"""

    # The following are from trial and error, and work only for India
    x0, y0 = 600, 230
    size = 1500

    lon, lat = lon * math.pi / 180, lat * math.pi / 180
    # Origin of Cartesian coordinates
    phi0, lambda0 = 24 * math.pi / 180, 80 * math.pi / 180
    # Standard parallels
    phi1, phi2 = 8 * math.pi / 180, 37 * math.pi / 180

    n = .5 * (math.sin(phi1) + math.sin(phi2))
    theta = n * (lon - lambda0)
    C = math.cos(phi1) ** 2 + 2 * n * math.sin(phi1)
    rho = ((C - 2 * n * math.sin(lat)) / n) ** .5
    rho0 = ((C - 2 * n * math.sin(phi0)) / n) ** .5
    x, y = rho * math.sin(theta), rho0 - rho * math.cos(theta)
    return x0 + x * size, y0 - y * size

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
            projection(x * sx + tx, y * sy + ty)
            for x, y in points
        ])

    for shape in topo['objects'].values():
        for geom in shape['geometries']:
            n_arcs = len(geom['arcs'])
            name = key(geom['properties'])

            # Convert arcs of a geometry into array of points
            for i, arcgroup in enumerate(geom['arcs']):
                # Ignoring holes at the moment
                points = []
                for arc in arcgroup:
                    points += coords[arc] if arc >= 0 else coords[~arc][::-1]

                # Draw the points
                shape = base.Shapes.BuildFreeform(msoEditingAuto, *points[0])
                for point in points[1:]:
                    shape.AddNodes(msoSegmentLine, msoEditingAuto, *point)
                shape = shape.ConvertToShape()
                shape.Line.Weight = 0.25
                shape.Line.ForeColor.ObjectThemeColor = msoThemeColorBackground1
                shape.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground2

                shape.Name = name if n_arcs == 1 else name + str(i)

            # Group shapes if required
            if n_arcs > 1:
                shape = base.Shapes.Range([name + str(i) for i in range(n_arcs)]).Group()
                shape.Name = name

            yield geom['properties']


Application = win32com.client.Dispatch("Excel.Application")
Workbook = Application.Workbooks.Add()
single_sheet = True
propcol = {}
row = 2

for pathspec in sys.argv[1:]:
    for filename in glob.glob(pathspec):
        print filename
        key = lambda v: v['ST_CODE'] + ':' + v['PC_NAME'].title()
        data = json.load(open(filename))
        if single_sheet:
            sheet = Workbook.ActiveSheet
        else:
            sheet = Workbook.Sheets.Add()
            row = 2

        for prop in draw(sheet, data, key):
            sheet.Cells(row, 1).Value = 0
            sheet.Cells(1, 1).Value = 'Value'
            for attr, val in prop.iteritems():
                if attr not in propcol:
                    propcol[attr] = len(propcol)
                sheet.Cells(row, propcol[attr] + 2).Value = val
            row += 1

        for attr, column in propcol.iteritems():
            sheet.Cells(1, column + 2).Value = attr

        if not single_sheet:
            sheet.Name = key

# Add visual basic code. http://www.cpearson.com/excel/vbe.aspx
# Requires Excel modification: http://support.microsoft.com/kb/282830
vbproj = Workbook.VBProject
for sheet in Workbook.Worksheets:
    codemod = vbproj.VBComponents(sheet.Name).CodeModule
    for line, row in enumerate(open('shape.bas')):
        codemod.InsertLines(line + 1, row)


Application.Visible = msoTrue

