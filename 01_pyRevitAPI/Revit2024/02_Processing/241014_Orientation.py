import clr
import sys 
import System   
import math

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB.Plumbing import*
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

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms.DataVisualization")

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

"_______________________________________________________________________________________"

def pickObjects():
    pipeList = []
    pipesRef = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element)
    for ref in pipesRef:
        element = doc.GetElement(ref.ElementId)
        pipeList.append(element)

    return pipesRef,pipeList

def pickObject():
    pipeList = []
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element)
    pipe = doc.GetElement(ref.ElementId)
    return pipe

pipe = pickObject()
pipeDiam1 = pipe.LookupParameter("Diameter").AsDouble()*304.8
pipeDiam2 = pipe.LookupParameter("Diameter").AsValueString()

OUT =  pipeDiam1, pipeDiam2
