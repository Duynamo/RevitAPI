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

#Define list/unwrap list functions
def toList(inList):
    result = inList if isinstance(inList, list) else [inList]
    return result

def dList(inList):
    result = inList if isinstance(inList, list) else [inList]
    return UnwrapElement(result)

# preparing input for dynamo to revit
element   = IN[0]
elements  = toList(IN[0])
d_List    = dList(IN[0]) 

# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)

# Now do something here

TransactionManager.Instance.TransactionTaskDone()

# Output and Changing element to Dynamo for export
# <element>.ToDSType(True), #Not created in script, mark as Revit-owned
# <element>.ToDSType(False) #Created in script, mark as non-Revit-owned

# Output to Dynamo
OUT = element