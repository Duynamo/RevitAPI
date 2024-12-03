"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

from math import cos,sin,tan,radians

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI import*
from Autodesk.Revit.UI.Selection import ISelectionFilter
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

# clr.AddReference("System.Windows.Forms")
# clr.AddReference("System.Drawing")
# clr.AddReference("System.Windows.Forms.DataVisualization")

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *
#endregion
#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

#region ___someFunctions
def uwList(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList
#endregion
temp=[]
def highlight_point():
    t=Transaction(doc,"highlight")
    t.Start()
    error=""
    distance=1/304.8
    try:
#region select point cloud
        pointCloudInstance = uidoc.Selection.PickBox(Selection.PickBoxStyle.Enclosing)
        pointCloudInstance = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element)
        pointCloudInstance=doc.GetElement(pointCloudInstance)
        boundingBox=pointCloudInstance.get_BoundingBox(None)
        Min=boundingBox.Min
        Max=boundingBox.Max
        planes=[]
        midpoint=(Min.Add(Max)).Multiply(0.5)
        #X boundaries
        planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisX,Min))
        planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisX.Negate(),midpoint))
        #Y boundaries
        planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisY,Min))
        planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisY.Negate(),midpoint))
        #Z boundaries
        planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisZ,Min))
        planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisZ.Negate(),midpoint))
        filter=PointClouds.PointCloudFilterFactory.CreateMultiPlaneFilter(planes)
        pointCloudInstance.FilterAction=SelectionFilterAction.Highlight
#endregion
#region pick section
        point_box=view.GetSectionBox()
        pick=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.PointOnElement)
        picked_point=pick.GlobalPoint
        Min_Box= point_box.Min
        Max_Box= point_box.Max
        midpoint_section=Min.Add(Max).Multiply(0.5)
        #trans identity
        trans=Transform.Identity
        trans.BasisX=point_box.Transform.BasisX
        trans.BasisY=point_box.Transform.BasisY
        trans.BasisZ=point_box.Transform.BasisZ
        trans.Origin=XYZ(picked_point.X,picked_point.Y,Max_Box.Z)
        temp.append(picked_point)
        temp.append(trans.Origin)
        #box identity
        box=BoundingBoxXYZ()
        box.MaxEnabled=True
        box.MinEnabled=True 
        box_height  =5000/304.8
        box_width   =100/304.8
        box_depth   =5000/304.8
        box.Min=XYZ(-box_height/(2),-box_width/(2),-box_depth/(2))
        box.Max=XYZ(box_height/(2),box_width/(2),box_depth/(2))
        box.Transform=trans
        view.SetSectionBox(box)
        view.IsSectionBoxActive=True   
        doc.ReloadLatest()
#endregion
        return 0
    except Exception as e:
        error=e
        t.Commit()
    return error

OUT=highlight_point()