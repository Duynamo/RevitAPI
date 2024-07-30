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
'''___'''
def allPipesInActiveView():
	pipesCollector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
	return pipesCollector
def getConnectedElements(pipes):
    connectedElements = []
    des = []
    for pipe in pipes:
        connectedElements = []
        conns = list(pipe.ConnectorManager.Connectors.GetEnumerator())
        # pipeCEParams = pipe.LookupParameter('CE_Length')
        for conn in conns:
            if conn.IsConnected:
                for refConn in conn.AllRefs:
                    connectedElement = refConn.Owner
                    if connectedElement.Id != pipeId:
                        connectedElements.append(connectedElement)
                        for connectedEle in connectedElements:
                             partType = connectedEle.MEPModel.partType
                             if partType == 'Elbow':
                                  des.append(connectedEle)
    return des

pipes = allPipesInActiveView()
connectedEles = getConnectedElements(pipes)

OUT = connectedEles

'''___'''