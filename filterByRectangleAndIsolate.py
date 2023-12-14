"""Copyright by: vudinhduybm@gmail.com"""
import clr 

import System
from System import Array
from System.Collections.Generic import *

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import*
from Autodesk.Revit.UI.Selection import*

#Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

class selectionFilter(ISelectionFilter):
    def __init__(self, category1, category2):
        self.category1 = category1
        self.category2 = category2

    
    def AllowElement(self, element):
        if element.Category.Name == self.category1 or element.Category.Name == self.category2 :
            return True
        else:
            return False
    def AllowReference(reference, point):
        return False
    
ele = selectionFilter('Pipes','Pipe Fittings')
eleAll = uidoc.Selection.PickElementsByRectangle(ele, "Hay chon gia dung")

IDS = List[Autodesk.Revit.DB.ElementId]()
for i in eleAll:
	IDS.Add(i.Id)

TransactionManager.Instance.EnsureInTransaction(doc)
isolateEles = view.IsolateElementsTemporary(IDS)
TransactionManager.Instance.TransactionTaskDone()

OUT = eleAll


	
