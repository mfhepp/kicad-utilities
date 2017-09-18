# kicad-utilities
My personal collection of KiCad scripts etc. 

So far only:
    bom.py  A script that creates an MS Excel file for the bill of materials and is able to combine all parts with identical part numbers (MPN), independent of part references.

Usage:

$ python bom.py test.xml

Creates test.xlsx with the BOM.

Requires Python 3.x and the Pandas library.
