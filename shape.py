import json
import win32com.client

msoEditingAuto = 0x0
msoSegmentLine = 0x0
msoTrue = -1
ppLayoutBlank = 0xc


def projection(lon, lat):
    return lon * 3, lat * 3

def draw(Base, topo):
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
            shape = Base.Shapes.BuildFreeform(msoEditingAuto, *points[0])
            for point in points[1:]:
                shape.AddNodes(msoSegmentLine, msoEditingAuto, *point)
            shape.ConvertToShape().Name = geom['properties']['PC_NAME']


Application = win32com.client.Dispatch("Excel.Application")
Application.Visible = msoTrue
Workbook = Application.Workbooks.Add()
Base = Workbook.ActiveSheet
draw(Base, json.load(open('maps/S10_PC.json')))
draw(Base, json.load(open('maps/S22_PC.json')))
draw(Base, json.load(open('maps/S01_PC.json')))
