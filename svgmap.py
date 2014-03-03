"""
Creates an Excel map application given an SVG map.
"""

import argparse
import color as _color
import win32com.client
from MSO import *
from svg2mso import svg2mso

parser = argparse.ArgumentParser(description=__doc__.strip())
parser.add_argument('svgfile')
parser.add_argument('-l', '--license', help='motherboard id')
parser.add_argument('-e', '--expiry', help='mm/dd/yyyy')
parser.add_argument('-a', '--attr', help='attribute to take ID from', default='title')
args = parser.parse_args()

Application = win32com.client.Dispatch("Excel.Application")
Application.Visible = msoTrue
Workbook = Application.Workbooks.Add()
Workbook.Sheets('Sheet1').Name = 'Map'
#Workbook.Sheets('Sheet2').Delete()
#Workbook.Sheets('Sheet3').Delete()

Base = Workbook.Sheets('Map')

def titles(e):
    while True:
        title = e.get(args.attr)
        if title is not None:
            yield title
            e = e.getparent()
        else:
            break

shapes = []
def callback(e, shape):
    shape.Fill.ForeColor.RGB = _color.msrgb('#ccc')
    shape.Fill.Visible = msoTrue
    shapes.append(shape)
    name = ':'.join(reversed(list(titles(e))))
    n = len(shapes)
    if not name:
        name = 'Shape%04d' % n
    Base.Cells(2 + n, 1).Value = 0
    Base.Cells(2 + n, 2).Value = shape.Name = name

svg = open(args.svgfile).read()
svg2mso(Base, svg, callback=callback)

# Set the gradient
Base.Cells(1, 1).Value = 'Colors'
Base.Cells(1, 2).Value = 0.0
Base.Cells(1, 3).Value = 0.5
Base.Cells(1, 4).Value = 1.0
Base.Cells(1, 2).Interior.Color = 255      # Red
Base.Cells(1, 3).Interior.Color = 65535    # Yellow
Base.Cells(1, 4).Interior.Color = 5296274  # Green

button = Base.Buttons().Add(332, 0, 48, 14.4)
button.OnAction = "Sheet1.Filter"
button.Characters.Text = "Filter"

button = Base.Buttons().Add(384, 0, 48, 14.4)
button.OnAction = "Sheet1.Refresh"
button.Characters.Text = "Refresh"

vbproj = Workbook.VBProject
codemod = vbproj.VBComponents('Sheet1').CodeModule
source = open('svgmap.bas').read()
if args.license:
    source = source.replace('LICENSEKEY', args.license)
if args.expiry:
    source = source.replace('01/01/2013', args.expiry)

for line, row in enumerate(source.split('\n')):
    codemod.InsertLines(line + 1, row)
