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
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList
def getPipeParameter(p):
    paramDiameters = []
    paramPipingSystems = []
    paramLevels = []
    paramPipeTypes = []

    paramDiameter = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8
    
    paramPipeTypeId = p.GetTypeId()
    paramPipeType = doc.GetElement(paramPipeTypeId)
	
    paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
    paramPipingSystem = doc.GetElement(paramPipingSystemId)

    paramLevelId = p.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
    paramLevel = doc.GetElement(paramLevelId)

    return [paramDiameter, paramPipingSystem, paramPipeType, paramLevel]

def offsetPointAlongVector(point, vector, offsetDistance):
    direction = vector.Normalize()
    scaledVector = direction.Multiply(offsetDistance)
    offsetPoint = point.ToRevitType().Add(scaledVector)
    return offsetPoint	
def pipeCreateFromPoints(desPointsList, sel_pipingSystem, sel_PipeType, sel_Level, diameter):
    lst_Points1 = [i for i in desPointsList]
    lst_Points2 = [i for i in desPointsList[1:]]
    linesList = []
    for pt1, pt2 in zip(lst_Points1,lst_Points2):
        line =  Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(pt1,pt2)
        linesList.append(line)
    firstPoint   = [x.StartPoint for x in linesList]
    secondPoint  = [x.EndPoint for x in linesList]
    pipesList = []
    pipesList1 = []
    TransactionManager.Instance.EnsureInTransaction(doc)
    for i,x in enumerate(firstPoint):
        try:
            levelId = sel_Level.Id
            sysTypeId = sel_pipingSystem.Id
            pipeTypeId = sel_PipeType.Id
            diam = diameter
            pipe = Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,sysTypeId,pipeTypeId,levelId,x.ToXyz(),secondPoint[i].ToXyz())
            param_Diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
            param_Diameter.SetValueString(diam.ToString())
            TransactionManager.Instance.EnsureInTransaction(doc)
            pipesList.append(pipe.ToDSType(False))
            pipesList1.append(pipe)
            # param_Length.Set()
            TransactionManager.Instance.TransactionTaskDone()
        except:
            pipesList.append(None)				
    TransactionManager.Instance.TransactionTaskDone()
    return pipesList
    
def createNewTee(pipe1, pipe2, pipe3):
    def NearestConnector(ConnectorSet, curCurve):
        MinLength = float("inf")
        result = None  # Initialize result to None
        for n in ConnectorSet:
            distance = curCurve.Location.Curve.Distance(n.Origin)
            if distance < MinLength:
                MinLength = distance
                result = n
        return result
    def closetConn(pipe1, pipe2, pipe3):
        connectors1 = list(pipe1.ConnectorManager.Connectors.GetEnumerator())
        connectors2 = list(pipe2.ConnectorManager.Connectors.GetEnumerator())
        connectors3 = list(pipe3.ConnectorManager.Connectors.GetEnumerator())
        Connector1 = NearestConnector(connectors1, pipe1)
        Connector2 = NearestConnector(connectors2, pipe2)
        Connector3 = NearestConnector(connectors3, pipe3) 
        return Connector1, Connector2, Connector3

    closetConn = closetConn(pipe1 , pipe2, pipe3)
    fittings = []
    TransactionManager.Instance.EnsureInTransaction(doc)
    fitting = doc.Create.NewTeeFitting(closetConn[0], closetConn[1], closetConn[2])
    fitting_dynamo = fitting.ToDSType(False)
    fittings.append(fitting_dynamo)
    TransactionManager.Instance.TransactionTaskDone()
    return fittings
#endregion


pulledPoint   = uwList(IN[0])
mainPipe	= UnwrapElement(IN[1])
XYZpulledPoint = pulledPoint.ToRevitType()
mainPipeParam = getPipeParameter(mainPipe)
mainPipeDiam = getPipeParameter(mainPipe)[0] / 304.8
tmpZ = (pulledPoint.Z+mainPipeDiam * 304.8)*5 / 304.8 
# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)
mainPipeConns = list(mainPipe.ConnectorManager.Connectors.GetEnumerator())
XYZconns = list(c.Origin for c in mainPipeConns)
toPointsConns = list(c.ToPoint() for c in XYZconns)
vector1 = pulledPoint.ToRevitType() - toPointsConns[0].ToRevitType()
vector2 = pulledPoint.ToRevitType() - toPointsConns[1].ToRevitType()
tmpPoint1 = offsetPointAlongVector(pulledPoint, vector1, mainPipeDiam).ToPoint()
tmpPoint2 = offsetPointAlongVector(pulledPoint, vector2, mainPipeDiam).ToPoint()
tmpPoint3 = XYZ(XYZpulledPoint.X, XYZpulledPoint.Y, tmpZ).ToPoint()

tmpListPoint1 = [pulledPoint, tmpPoint3]
tmpListPoint2 = [toPointsConns[0], tmpPoint2]
tmpListPoint3 = [toPointsConns[1], tmpPoint1]

tmpPipe1 = flatten(pipeCreateFromPoints(tmpListPoint1, mainPipeParam[1], mainPipeParam[2], mainPipeParam[3], mainPipeParam[0]))
tmpPipe2 = pipeCreateFromPoints(tmpListPoint2, mainPipeParam[1], mainPipeParam[2], mainPipeParam[3], mainPipeParam[0])
tmpPipe3 = pipeCreateFromPoints(tmpListPoint3, mainPipeParam[1], mainPipeParam[2], mainPipeParam[3], mainPipeParam[0])

pipes = flatten([tmpPipe1, tmpPipe2, tmpPipe3])

#newTee = createNewTee(pipes[0], pipes[1], pipes[2])
TransactionManager.Instance.TransactionTaskDone()

OUT = pipes
