"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

clr.AddReference("ProtoGeometry")

from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB import Line, XYZ, IntersectionResultArray, SetComparisonResult

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI.Selection import ISelectionFilter
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
#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

class selectionFilter(ISelectionFilter):
    def __init__(self, category):
          self.category = category
    def AllowElement(self, element):
        if element.Category.Name == self.category:
            return True
        else:
            return False
def pickPipe():
    pipeFilter = selectionFilter('Pipes')
    pipeRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter)
    pipe = doc.GetElement(pipeRef.ElementId)
    return pipe

def pickFittings():
    pipeFilter = selectionFilter( 'Pipe Fittings')
    pipeRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter)
    pipe = doc.GetElement(pipeRef.ElementId)
    return pipe



def rotate_tee_and_pipe(pipeA, teeA, angle_degrees):
    """
    Xoay teeA quanh trục của pipeA một góc angle_degrees.
    
    Args:
        pipeA: Pipe chính
        teeA: Tee fitting trên pipeA
        angle_degrees: Góc xoay (độ)
    
    Returns:
        bool: True nếu thành công
    """
    try:
        if not teeA:
            # print("Error: No Tee fitting provided.")
            return False
        
        # Lấy trục xoay từ pipeA
        pipeA_curve = pipeA.Location.Curve
        P1 = pipeA_curve.GetEndPoint(0)
        P2 = pipeA_curve.GetEndPoint(1)
        axis_dir = (P2 - P1).Normalize()
        
        # Tâm xoay là vị trí teeA
        teeA_location = teeA.Location
        if not hasattr(teeA_location, "Point"):
            # print("Error: TeeA has no location point.")
            return False
        center = teeA_location.Point
        
        # Tạo trục xoay
        axis_line = Line.CreateUnbound(center, axis_dir)
        
        # Góc xoay (radian)
        angle_radians = math.radians(angle_degrees)
        
        # Bắt đầu transaction
        with Transaction(doc, "Rotate Tee") as t:
            t.Start()
            
            # Xoay teeA
            ElementTransformUtils.RotateElement(
                doc, teeA.Id, axis_line, angle_radians
            )
            
            t.Commit()
        
        # print(f"Rotated teeA {angle_degrees} degrees around pipeA.")
        return True
    
    except Exception as e:
        # print(f"Error: {str(e)}")
        return False

pipeA = pickPipe()
teeA = pickFittings()
# pipeB = pickPipe()
angle = 22.5    

ro = rotate_tee_and_pipe(pipeA , teeA , -(90-angle))
OUT = ''