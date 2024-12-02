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
	pipesCollector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
	return pipesCollector
def getConnectedElements(doc, pipes):
    connected_elements = []
    settedPipesMTOLength = []
    for pipe in pipes:
        conns = pipe.ConnectorManager.Connectors
        for conn in conns:
            if conn.IsConnected:
                for refConn in conn.AllRefs:
                    connectedElement = refConn.Owner
                    if connectedElement.Id != pipe.Id:
                        pipe_MTOLength_param = pipe.LookupParameter('FU_MTO Length')
                        pipeLength = pipe.LookupParameter('Length').AsDouble()
                        connectedEle_MTOLength_param_check = connectedElement.LookupParameter('FU_MTO Length')
                        if connectedEle_MTOLength_param_check != None:
                            connectedEle_MTOLength_param = connectedEle_MTOLength_param_check.AsDouble()
                            TransactionManager.Instance.EnsureInTransaction(doc)
                            if connectedEle_MTOLength_param != None:
                                pipe_MTOLength_param.Set(connectedEle_MTOLength_param+pipeLength)
                            else:
                                pipe_MTOLength_param.Set(pipeLength)
                            TransactionManager.Instance.TransactionTaskDone()
                            connected_elements.append(connectedElement)
                            settedPipesMTOLength.append(pipe)
                        else: pass
            else:
                pipeLength1 = pipe.LookupParameter('Length').AsDouble()
                pipe_MTOLength_param1 = pipe.LookupParameter('FU_MTO Length')
                TransactionManager.Instance.EnsureInTransaction(doc)
                pipe_MTOLength_param1.Set(pipeLength1)
                TransactionManager.Instance.TransactionTaskDone()
                settedPipesMTOLength.append(pipe)
    return  settedPipesMTOLength

pipes = allPipesInActiveView()
settedPipesMTOLength = getConnectedElements(doc,pipes)

OUT =  settedPipesMTOLength

'''___'''