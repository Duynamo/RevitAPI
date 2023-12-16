"""Copyright by: vudinhduybm@gmail.com"""
import clr #This is .NET's Common Language Runtime.
import sys #sys is a fundamental Python library
           #we're using it to load the standard IronPython libraries
import System   #The System namespace at the root of .NET
clr.AddReference("ProtoGeometry")
#A Dynamo library for its proxy geometry classes
#You'll only need it if you're interacting with geometry
from Autodesk.DesignScript.Geometry import *
#Loads everything in Dynamo's geometry library

clr.AddReference('RevitAPI') #Adding reference to Revit's APU .dll
from Autodesk.Revit.DB import*
from Autodesk.Revit.DB.Structure import*

clr.AddReference('RevitAPIUI') #Adding reference to Revit's API.dll
from Autodesk.Revit.UI import*

clr.AddReference('System') 
from System.Collections.Generic import List
#Revit's API sometimes wants hard-types 'generic' lists, called IList

clr.AddReference('RevitNodes') #Dynamo nodes for revit

import Revit #Loads in the Revit namespace in  Revit nodes
clr.ImportExtensions('Revit.GeometryConversions') #You'll only need if you're interacting with geometries 
clr.ImportExtensions('Revit.Elements') #More loading of Dynamo's Revit libraries
clr.AddReference('RevitService') #Dynamo's classes for handling Revit documents

import RevitServices
from RevitServices.Persistence import DocumentManager
#An internal Dynamo class
#That keeps track of the document that Dynamo is currently attched to

from RevitServices.Transactions import TransactionManager
#An Dynamo class for opening and closing transactions to change the Revit document's database

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