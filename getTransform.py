"""Copyright by: vudinhduybm@gmail.com"""
import clr 
import System
import math 
from System.Collections.Generic import *

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import*
from Autodesk.Revit.UI.Selection import*
from Autodesk.Revit.UI.Selection import ObjectType
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

# inlist = IN[0]

def pickObjects():
    refs = uidoc.Selection.PickObjects(ObjectType.Element)
    return  [doc.GetElement(i.ElementId) for i in refs]
inEle = pickObjects()

for j in inEle:
    transform = j.GetTransform()
    transScale = transform.Scale

    basicX_scaled = (transform.BasisX.X / transScale, transform.BasisX.Y / transScale, transform.BasisX.Z / transScale)
    basicY_scaled = (transform.BasisY.X / transScale, transform.BasisY.Y / transScale, transform.BasisY.Z / transScale)
    basicZ_scaled = (transform.BasisZ.X / transScale, transform.BasisZ.Y / transScale, transform.BasisZ.Z / transScale)
    origin_scaled = (transform.Origin.X / transScale, transform.Origin.Y / transScale, transform.Origin.Z / transScale)

    result = {

        "BasicX_scaled": basicX_scaled,
        "BasicY_scaled": basicY_scaled,
        "BasicZ_scaled": basicZ_scaled,
        "Scale": transScale,
        "Origin_scaled": origin_scaled
    }

OUT = [result[key] for key in ["BasicX_scaled", "BasicY_scaled", "BasicZ_scaled", "Scale", "Origin_scaled"]]

