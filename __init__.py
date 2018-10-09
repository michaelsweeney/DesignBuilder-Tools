import sys

pydir = "C:/PythonScripts/eQuest"
if pydir not in sys.path:
    sys.path.insert(0,pydir)
        
import pandas as pd
#import objects
#import polygon_areas
#
##args
#inpfile = objects.openinp(r'P:\_Projects\G150000\G150015-000\Calculations\Mech\Energy Model\eQuest\eQuest Model\July 2018 Models\Proposed\CPL.inp')
#spaces = objects.get(inpfile,"SPACE")
#zones = objects.get(inpfile,"ZONE")
#polygons = objects.get(inpfile,"POLYGON")
#floors = objects.get(inpfile,"FLOOR")