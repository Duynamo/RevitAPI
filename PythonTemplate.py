"""Copyright by: vudinhduybm@gmail.com"""
import clr 
import sys 
import System   
clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitAPI') 
from Autodesk.Revit.DB import*
from Autodesk.Revit.DB.Structure import*

clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.UI import*

clr.AddReference('System') 
from System.Collections.Generic import List

clr.AddReference('RevitNodes') 
import Revit 
clr.ImportExtensions('Revit.GeometryConversions') 
clr.ImportExtensions('Revit.Elements') 
clr.AddReference('RevitService') 

import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

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

OUT = element