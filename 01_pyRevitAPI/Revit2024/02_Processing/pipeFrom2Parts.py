"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

clr.AddReference("ProtoGeometry")
from math import *
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
class selectionFilter(ISelectionFilter):
    def __init__(self,categories):
        self.categories = categories
    def AllowElement(self, element):
        eleName = element.Category.Name in self.categories
        return eleName
def pickPipeOrPipePart():
    categories = ['Pipe Fittings', 'Pipe Accessories','Pipes']
    fittingFilter = selectionFilter(categories)
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory or Pipe')
    desEle = doc.GetElement(ref.ElementId)
    return desEle
def pickPipe():
    categories = 'Pipes'
    fittingFilter = selectionFilter(categories)
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, fittingFilter, 'Pick fitting or Accessory or Pipe')
    desEle = doc.GetElement(ref.ElementId)
    return desEle
def NearestConnector(connList, targetPart):
    TransactionManager.Instance.EnsureInTransaction(doc)
    if targetPart.Category.Name != 'Pipes':
        targetConns = list(targetPart.MEPModel.ConnectorManager.Connectors)
    if targetPart.Category.Name == 'Pipes':
        targetConns = list(targetPart.ConnectorManager.Connectors)
    nearestConn = None
    minDistance = float('inf')
    for conn in connList:
        for targetConn in targetConns:
            distance = conn.Origin.DistanceTo(targetConn.Origin)
            if distance < minDistance:
                minDistance = distance
                nearestConn = conn
    return nearestConn    
def closetConn(firstPart, secondPart): #return nearest connector of secondElbow to firstElbow
    connectors1 = list(secondPart.MEPModel.ConnectorManager.Connectors.GetEnumerator())
    Connector1 = NearestConnector(connectors1, firstPart)
    XYZconn	= Connector1.Origin
    return XYZconn    
def sortConnsOf1stElbow_To2ndElbow(firstPart, secondPart):
    sortConnList = []
    originConns  = []
    tmp = []
    tmp1 = []
    tmp2 = []
    if firstPart.Category.Name != 'Pipes':
        connList = list(firstPart.MEPModel.ConnectorManager.Connectors.GetEnumerator())
    if firstPart.Category.Name == 'Pipes':
        connList = list(firstPart.ConnectorManager.Connectors.GetEnumerator()) 
    nearConn = NearestConnector(connList, secondPart)
    tmp1.append(nearConn)
    farConn =  [c for c in connList  if c not in tmp1]
    tmp1.append([c for c in connList if c != nearConn])
    sortConnList = flatten(tmp1)
    tmp2.append([c.Origin for c in sortConnList])
    originConns = flatten(tmp2)
    return originConns
def getConnectedElements(elbow):
    connectedElements = []
    connList = list(elbow.MEPModel.ConnectorManager.Connectors)
    for conn in connList:
        for ref in conn.AllRefs:
            if ref.IsConnected:
                connectedElement = doc.GetElement(ref.Owner.Id)
                connectedElements.append(connectedElement)
    return connectedElements
def getPipeParameter(p):
    paramDiameters = []
    paramPipingSystems = []
    paramLevels = []
    paramPipeTypes = []

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
def findNearestConnector(connectorSet, targetPoint):
    nearest_connector = None
    minDistance = float('inf')
    for connector in connectorSet:
        distance = (connector.Origin - targetPoint).GetLength()
        if distance < minDistance:
            minDistance = distance
            nearestConnector = connector
    return nearestConnector
# Function to connect elbow to pipe
def connectFittingToPipe(doc, fitting, pipe):
    TransactionManager.Instance.EnsureInTransaction(doc)
    # Get the connectors of the elbow and pipe
    fittingConnectors = [conn for conn in fitting.MEPModel.ConnectorManager.Connectors]
    pipeConnectors = [conn for conn in pipe.ConnectorManager.Connectors]    
    # Find the nearest connectors between the fitting and pipe
    nearestElbowConn = findNearestConnector(fittingConnectors, pipeConnectors[0].Origin)
    nearestPipeConn = findNearestConnector(pipeConnectors, nearestElbowConn.Origin)    
    # Connect the nearest connectors
    nearestElbowConn.ConnectTo(nearestPipeConn)    
    TransactionManager.Instance.TransactionTaskDone()
    return nearestElbowConn, nearestPipeConn
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList
#endregion
firstPart   = pickPipeOrPipePart()
secondPart	= pickPipeOrPipePart()
basePipe = pickPipe()
tempPoint = []
# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)
nf_firstPartConns = sortConnsOf1stElbow_To2ndElbow(firstPart, secondPart)[0].ToPoint()
nf_secondPartConns = sortConnsOf1stElbow_To2ndElbow(secondPart, firstPart)[0].ToPoint()
TransactionManager.Instance.TransactionTaskDone()
tempPoint.append(nf_firstPartConns)
tempPoint.append(nf_secondPartConns)
#get base parameters of base part
baseParams = getPipeParameter(basePipe)
#create pipe from point list and base pipe parameters
TransactionManager.Instance.EnsureInTransaction(doc)
basePipeParam = getPipeParameter(basePipe)
diamParam = baseParams[0][0]
pipingSystemParam = baseParams[0][1]
pipeTypeParam = baseParams[0][2]
levelParam = baseParams[0][3]
tempPipe = pipeCreateFromPoints(tempPoint, pipingSystemParam, pipeTypeParam, levelParam, diamParam)
TransactionManager.Instance.TransactionTaskDone()
tempPipe = pipeCreateFromPoints()
#connect tempPipe to pipe part 1 and pipe part 2
TransactionManager.Instance.EnsureInTransaction(doc)
if firstPart and tempPipe:
	elbowConn, pipeConn = connectFittingToPipe(doc, firstPart, tempPipe)
if secondPart and tempPipe:
	elbowConn1, pipeConn1 = connectFittingToPipe(doc, secondPart, tempPipe)	
TransactionManager.Instance.TransactionTaskDone()

OUT = tempPipe