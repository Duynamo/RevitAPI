"""Copyright by: vudinhduybm@gmail.com"""
import clr 
from System.Collections.Generic import *

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB.Events import*

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialog

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager

doc = DocumentManager.Instance.CurrentDBDocument
view = doc.ActiveView
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

inList = IN[0]
x=True
dypoint = []
rpointM = []
rpointI = []
counter=0
msg = 'Pick Points on current Workplane in order, hit ESC when finished.'

TaskDialog.Show("Duynamo", msg)

while x == True:
    if len(inList) != 0 :
        pt=uidoc.Selection.PickPoint()
        rpM=Point.ByCoordinates(pt.X*304.8,pt.Y*304.8,pt.Z*304.8)
        rpI=Point.ByCoordinates(pt.X,pt.Y,pt.Z)
        counter=+1
        dypoint.append(pt)
        rpointM.append(rpM)
        rpointI.append(rpI)

    else: 
        x=False

OUT=(dypoint,rpointM,rpointI)









