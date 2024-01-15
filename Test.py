import clr
import sys 
import System   


clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*


clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI import*

clr.AddReference("System") 
from System.Collections.Generic import List

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms.DataVisualization")

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

# # inValue = UnwrapElement(IN[1])
# x = True
# dypoint = []
# rpointM = []
# rpointI = []
# counter=0
# msg = 'Pick Points on current Workplane in order, hit ESC when finished.'

# TaskDialog.Show("Duynamo", msg)

# while x == True :
# 	# msg = 'Pick Points on current Workplane in order, hit ESC when finished.'

# 	# TaskDialog.Show("Duynamo", msg)
# 	try:
# 		pt=uidoc.Selection.PickPoint()
# 		# rpM=Point.ByCoordinates(pt.X*304.8,pt.Y*304.8,pt.Z*304.8)# don vi do mac dinh trong revitAPI la Inch, 1in=304.8mm
# 		# rpI=Point.ByCoordinates(pt.X,pt.Y,pt.Z)
# 		counter=+1
# 		dypoint.append(pt)
# 		# rpointM.append(rpM)
# 		# rpointI.append(rpI)
# 	except:
# 		x=False
	

# OUT=(dypoint,rpointM,rpointI)

# my_list = [1, 2, 3, 4, 5]

# # Remove the first item
# # if my_list:
# #     my_list.pop(0)
    
# newlist = my_list.pop(0)

# print(my_list)



# TransactionManager.Instance.EnsureInTransaction(doc)
# for i,x in enumerate(FirstPoint):
# 	try:
# 		levelid = level[i%ll].Id
# 		systypeid = systemtype[i%stl].Id
# 		pipetypeid = pipetype[i%ptl].Id
# 		diam = diameter[i%dl]
		
# 		pipe = Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,systypeid,pipetypeid,levelid,x.ToXyz(),SecondPoint[i].ToXyz())
		
# 		param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
# 		param.SetValueString(diam.ToString())
	
# 		elements.append(pipe.ToDSType(False))	
# 	except:
# 		elements.append(None)

# TransactionManager.Instance.TransactionTaskDone()

p1 = Point.ByCoordinates(1000, 1000, 1000)
p2 = Point.ByCoordinates(1500, 1500, 1500)
p3 = Point.ByCoordinates(500, 500, 500)
p4 = Point.ByCoordinates(0, 0, 0)

lst1 = []

for l in (p1, p2, p3, p4):
    for e in l:
        lst1.append(e)

print(lst1)