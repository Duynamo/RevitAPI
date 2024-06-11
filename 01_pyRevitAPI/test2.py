"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

sys.path.append('C:\Program Files\Autodesk\Revit 2022\AddIns\DynamoForRevit\IronPython.StdLib.2.7.9\duynamoLibrary')
from duynamoLibrary import *
clr.AddReference("ProtoGeometry")

from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*

clr.AddReference('DSCoreNodes')
from DSCore.List import Flatten

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI import*

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
#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

#region ___someFunctions
def uwList(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)
def getPipeParameter(pipes):
    paramDiameters = []
    paramPipingSystems = []
    paramLevels = []
    paramPipeTypes = []
    for p in pipes:
        paramDiameter = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8
        paramPipeTypeId = p.GetTypeId()
        paramPipeType = doc.GetElement(paramPipeTypeId)
        paramPipeTypes.append(paramPipeType)
        pipeTypeName = paramPipeType.LookupParameter("Type Name").AsString()
        paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
        paramPipingSystem = doc.GetElement(paramPipingSystemId)
        paramPipingSystems.append(paramPipingSystem)
        pipingSystemName = paramPipingSystem.LookupParameter("System Classification").AsValueString()
        paramLevelId = p.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
        paramLevel = doc.GetElement(paramLevelId)
        paramLevels.append(paramLevel)

    return [paramDiameter, paramPipingSystem, paramPipeType, paramLevel],[paramDiameter,pipingSystemName,pipeTypeName,paramLevel]
#endregion


eleList   = uwList(IN[0])

# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)


TransactionManager.Instance.TransactionTaskDone()

OUT = eleList