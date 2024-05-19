import clr
import sys 
import System   
import math
from System.Collections.Generic import *
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

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms.DataVisualization")
clr.AddReference("DSCoreNodes")
from DSCore.List import Flatten

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *

"""_____________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = uiapp.ActiveUIDocument
view = doc.ActiveView
"""_________"""

faces = []
dsfaces = []

faces.append(uidoc.Selection.PickObjects(ObjectType.Face,'Select Faces'))
for f in faces:
	for r in f:
		e = uidoc.Document.GetElement(r)
		dsface = e.GetGeometryObjectFromReference(r).ToProtoType(True)
		[f.Tags.AddTag("RevitFaceReference", r) for f in dsface]
		dsfaces.append(dsface)