"""
Usage: python32 svgmap.py d:/site/gramener.com/www/src/indiamap/india-states.svg UNID
"""

import sys
import color as _color
import win32com.client
from MSO import *
from svg2mso import svg2mso

Application = win32com.client.Dispatch("Excel.Application")
Application.Visible = msoTrue
Workbook = Application.Workbooks.Add()
Workbook.Sheets('Sheet1').Name = 'Map'
#Workbook.Sheets('Sheet2').Delete()
#Workbook.Sheets('Sheet3').Delete()

Base = Workbook.Sheets('Map')

def titles(e):
    while True:
        title = e.get('title')
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
    Base.Cells(2 + len(shapes), 1).Value = 0
    Base.Cells(2 + len(shapes), 2).Value = shape.name = name

vbscript = open(sys.argv[1]).read()
svg2mso(Base, vbscript, callback=callback)

# Set the gradient
Base.Cells(1, 1).Value = 'Colors'
Base.Cells(1, 2).Value = 0.0
Base.Cells(1, 3).Value = 0.5
Base.Cells(1, 4).Value = 1.0
Base.Cells(1, 2).Interior.Color = 255      # Red
Base.Cells(1, 3).Interior.Color = 65535    # Yellow
Base.Cells(1, 4).Interior.Color = 5296274  # Green

vbproj = Workbook.VBProject
codemod = vbproj.VBComponents('Sheet1').CodeModule
for line, row in enumerate(open('svgmap.bas')):
    codemod.InsertLines(line + 1, row.replace('LICENSEKEY', sys.argv[2]))
    