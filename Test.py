# """Copyright by: vudinhduybm@gmail.com"""
# import clr 
# from System.Collections.Generic import *

# clr.AddReference("ProtoGeometry")
# from Autodesk.DesignScript.Geometry import *

# clr.AddReference('RevitAPI')
# from Autodesk.Revit.DB.Events import*

# clr.AddReference('RevitAPIUI')
# from Autodesk.Revit.UI import TaskDialog

# clr.AddReference("RevitServices")
# import RevitServices
# from RevitServices.Persistence import DocumentManager

# doc = DocumentManager.Instance.CurrentDBDocument
# view = doc.ActiveView
# uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

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

my_list = [1, 2, 3, 4, 5]

# Remove the first item
# if my_list:
#     my_list.pop(0)
    
newlist = my_list.pop(0)

print(my_list)