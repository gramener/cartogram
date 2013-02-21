# -*- coding: utf8 -*-
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


def projection(lon, lat):
    """Albers: http://mathworld.wolfram.com/AlbersEqual-AreaConicProjection.html"""
    x0, y0 = 300, 230
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
            # Convert arcs of a geometry into array of points
            points = []
            for arcgroup in geom['arcs']:
                for arc in arcgroup:
                    points += coords[arc] if arc >= 0 else coords[~arc][::-1]

            # Draw the points
            shape = base.Shapes.BuildFreeform(msoEditingAuto, *points[0])
            for point in points[1:]:
                shape.AddNodes(msoSegmentLine, msoEditingAuto, *point)
            shape = shape.ConvertToShape()
            shape.Name = key(geom['properties'])
            shape.Fill.Visible = msoFalse
            shape.Line.Weight = 0.25
            shape.Line.ForeColor.ObjectThemeColor = msoThemeColorText1

            yield geom['properties']


Application = win32com.client.Dispatch("Excel.Application")
Application.Visible = msoTrue
Workbook = Application.Workbooks.Add()
single_sheet = True
row = 1

for filename in glob.glob('maps/S*_PC.json'):
    key = lambda v: v['PC_NAME']
    data = json.load(open(filename))
    if single_sheet:
        sheet = Workbook.ActiveSheet
    else:
        sheet = Workbook.Sheets.Add()
        row = 1

    for prop in draw(sheet, data, key):
        sheet.Cells(row, 1).Value = 0
        sheet.Cells(row, 2).Value = prop['ST_NAME']
        sheet.Cells(row, 3).Value = key(prop)
        row += 1

    if not single_sheet:
        sheet.Name = prop['ST_NAME']


# Below is the Visual Basic code to be added to the Worksheet
"""
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range

    ' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    Set KeyCells = Range("A:A")

    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then

        For Each Cell In Target.Cells
            Set Shape = ActiveSheet.Shapes(Cell.Offset(0, 2).value)
            Shape.Fill.BackColor.RGB = Gradient(Cell.value)
        Next

    End If
End Sub

Public Function Gradient(value)
    ' We'll always use a 3 point scale: 0, .5, 1
    g = Array(Array(1, 0, 0), Array(1, 1, 0), Array(0, 1, 0))
    Dim result As Variant

    If value < 0 Then
        result = g(0)
    ElseIf value < 0.5 Then
        a = g(0)
        b = g(1)
        q = 2 * value
        p = 1 - q
        result = Array(a(0) * p + b(0) * q, a(1) * p + b(1) * q, a(2) * p * b(2) * q)
    ElseIf value <= 1# Then
        a = g(1)
        b = g(2)
        q = 2 * (value - 0.5)
        p = 1 - q
        result = Array(a(0) * p + b(0) * q, a(1) * p + b(1) * q, a(2) * p * b(2) * q)
    Else
        result = g(2)
    End If

    Gradient = RGB(result(0) * 255, result(1) * 255, result(2) * 255)

End Function
"""
