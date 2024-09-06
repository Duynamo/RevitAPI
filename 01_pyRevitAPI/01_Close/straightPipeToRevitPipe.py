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
"""_________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

class selectionFilter(ISelectionFilter):
	def __init__(self, category):
		self.category = category
	def AllowElement(self, element):
		if element.Category.Name == self.category:
			return True
		else:
			return False
def pickPipe():
	pipes = []
	pipeFilter = selectionFilter("Pipe Fittings")
	pipeRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter,"pick Pipe")
	pipe = doc.GetElement(pipeRef.ElementId)
	return pipe
def getPipeTypeByName(doc, pipeTypeName):
    pipeTypes = FilteredElementCollector(doc).OfClass(PipeType).ToElements()
    for pipeType in pipeTypes:
        pipeTypeNameParam = pipeType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if pipeTypeNameParam == pipeTypeName:
            return pipeType
    return None
def getPipingSystemByName(doc, systemName):
    pipingSystems = FilteredElementCollector(doc).OfClass(PipingSystemType).ToElements()
    for pipingSystem in pipingSystems:
        systemNameParam = pipingSystem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if systemNameParam == systemName:
            return pipingSystem
    return None
def getLevelByName(doc, levelName):
    levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    for level in levels:
        levelNameParam = level.get_Parameter(BuiltInParameter.DATUM_TEXT).AsString()
        if levelNameParam == levelName:
            return level
    return None

"""_________________________________________"""
straightPipe = pickPipe()
conns = [conn for conn in straightPipe.MEPModel.ConnectorManager.Connectors]
connsOrigin = [c.Origin for c in conns]
#region __lookup and set union pipe parameters 
TransactionManager.Instance.EnsureInTransaction(doc)
union_Diameter_param = straightPipe.LookupParameter('_Diameter').AsString()
pipe_Diameter_param = float(union_Diameter_param)
union_PipeType_param = straightPipe.LookupParameter('_Pipe Type').AsString()
pipe_PipeType_param = getPipeTypeByName(doc, union_PipeType_param)
union_PipingSystem_param = straightPipe.LookupParameter('_Piping System').AsString()
pipe_PipingSystem_param = getPipingSystemByName(doc, union_PipingSystem_param)
union_ReferenceLevel_param = straightPipe.LookupParameter('_Reference Level').AsString()
pipe_ReferenceLevel_param = getLevelByName(doc, union_ReferenceLevel_param)
TransactionManager.Instance.TransactionTaskDone()
#endregion
pipe = Pipe.Create(doc,pipe_PipingSystem_param.Id, pipe_PipeType_param.Id, pipe_ReferenceLevel_param.Id, connsOrigin[0], connsOrigin[1])
pipe_Diameter = pipe.LookupParameter('Diameter')
pipe_Diameter.Set(pipe_Diameter_param/304.8)
TransactionManager.Instance.EnsureInTransaction(doc)
if pipe:
    doc.Delete(straightPipe.Id)
TransactionManager.Instance.TransactionTaskDone()
OUT =  pipe