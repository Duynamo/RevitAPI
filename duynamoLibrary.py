#def 0001
#to Flatten nested lists to expect Level
def flattenTo1dList(arr):
    result = []
    def recursive_flatten(subarray):
        for item in subarray:
            if is_arrayitem(item):
                recursive_flatten(item)
            else:
                result.append(item) 
    recursive_flatten(arr)        
    return result

def isArray(array):
    return "List" in obj.__class__.__name__

def getArrayRank(array):
    if isArray(array)
        return 1 + max(getArrayRank(item) for item in array)
    else: 
        return 0

def flattenArrayByOptionals(objects, level = 1):
    result = []
    for item in objects:
        newRank = getArrayRank(item)
        if isArray(item) and newRank >= level :
            arr = flattenTo1dList(item)
            result.append(arr)
        else:
            result.append(item)
    return result

#region ___selectionFilter
from Autodesk.Revit.UI.Selection import ISelectionFilter
class selectionFilter(ISelectionFilter):
    def __init__(self, category1, category2,category3, category4):
        self.category1 = category1
        self.category2 = category2
        self.category3 = category3
        self.category4 = category4
    
    def AllowElement(self, element):
        if element.Category.Name == self.category1 or element.Category.Name == self.category2 or element.Category.Name == self.category3 or element.Category.Name == self.category4:
            return True
        else:
            return False
    def AllowReference(reference, point):
        return False
    
ele = selectionFilter('Structural Columns', 'Structural Framing', 'Pipes','Pipe Fittings')

#endrigion

#region___pickObjects and return Elements and its references
def pickObjects():
    from Autodesk.Revit.UI.Selection import ObjectType
    refs = uidoc.Selection.PickObjects(ObjectType.Element)
    return  [doc.GetElement(i.ElementId) for i in refs], [ref for ref in refs]
#endregion

#region ___to pick objects and return the picked elements list, refs and Ids
def pickObjects():
    from Autodesk.Revit.UI.Selection import ObjectType
    refs = uidoc.Selection.PickObjects(ObjectType.Element)
    elesId = []
    elesRef = []
    eles = []
    eles.append([doc.GetElement(i.ElementId) for i in refs])
    elesId.append(i.ElementId for i in refs)
    elesRef.append(ref for ref in refs)
    return  eles, elesId, elesRef
OUT = pickObjects()
#endregion

#region ___to get piping system from string Name
def getAllPipingSystemsInActiveView(doc):
	collector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_PipingSystem)
	pipingSystems = collector.ToElements()
	pipingSystemsName = []
	for system in pipingSystems:
		systemName = system.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
		pipingSystemsName.append(systemName)
	return pipingSystemsName, pipingSystems
allPipingSystemsInActiveView = getAllPipingSystemsInActiveView(doc)
allPipingSystemName = allPipingSystemsInActiveView[0]
allPipingSystem = allPipingSystemsInActiveView[1]
idOfInSystem = allPipingSystemName.index(inSystemName)
desPipingSystem = allPipingSystem[idOfInSystem]
#endregion

#region ___to run code in VsCode
re  = open(r"C:\Users\95053\Desktop\Python\RevitAPI-master\RevitAPI\WPF\\MainForm1.py", "r")
interpret = re.read()
#endregion

#region ___to get family by name
categories = [BuiltInCategory.OST_PipeAccessory]
desFamTypes = []
key = "FU_Support"
categoriesFilter = []
for category in categories:
    elementTypes = FilteredElementCollector(doc).OfCategory(category).WhereElementIsElementType().ToElements()
    for elementType in elementTypes:
        typeName = elementType.FamilyName
        if key in typeName:
            desFamTypes.append(elementType)
categoriesFilter.append(desFamTypes)
flat_categoriesFilter = [ item for sublist in categoriesFilter for item in sublist]
#endregion

#region ___to get pipe location curve
def getPipeLocationCurve(pipe):
    lCurve = pipe.Location.Curve
    return lCurve
#endregion

#region ___to pick pipe only
class selectionFilter(ISelectionFilter):
	def __init__(self, category):
		self.category = category
	def AllowElement(self, element):
		if element.Category.Name == self.category:
			return True
		else:
			return False
def pickPipes():
	pipes = []
	pipeFilter = selectionFilter("Pipes")
	pipesRef = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter,"pick Pipes")
	for ref in pipesRef:
		pipe = doc.GetElement(ref.ElementId)
		pipes.append(pipe)
	return pipes	
#endregion

#region ___to pop up alert messages
check = IN[0]
content = IN[1]
result = str(content)
button = TaskDialogCommontButtons.None
#button = TaskDialogCommontButtons.Ok
#button = TaskDialogCommontButtons.Cancel
#button = TaskDialogCommontButtons.Close
#button = TaskDialogCommontButtons.Retry
#button = TaskDialogCommontButtons.Yes

if check == True:
     TaskDialog.Show('Result', result, button)
else:
     result = 'Set True to Run'
OUT = result
#endregion

#region ___to get all categories in project
def allCateInPJ():
    categories = doc.Settings.Categories
    modelCate = []
    for c in categories:
        if c.CategoryType == CategoryType.Model and c.CanAddSubcategory:
            modelCate.append(Revit.Elements.Category.ById(c.Id.IntegerValue))
    return modelCate
#endregion

#region ___to Define list/unwrap list functions
def uwlist(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)
#endregion

#region ___to retrieve selected pipe type, piping system, diameter and reference level
def getPipeParameter(pipe):
    paramDiameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble()*304.8
    paramPipeTypeId = pipe.GetTypeId()
    paramPipeType = doc.GetElement(paramPipeTypeId)
    paramPipingSystemId = pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
    paramPipingSystem = doc.GetElement(paramPipingSystemId)
    paramLevelId = pipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
    paramLevel = doc.GetElement(paramLevelId)
    return  paramDiameter, paramPipingSystem, paramLevel, paramPipeType
#endregion

#region ___to delete elements
#you need to flat the Input list before delete it
TransactionManager.Instance.EnsureInTransaction(doc)
results,fails = [],[]
for e in elements:
	id = None
	try:
		id = e.Id
		del_id = doc.Delete(id)
		results.append(True)
	except:
		if id is not None:
			fails.append(e)		
		results.append(False)
TransactionManager.Instance.TransactionTaskDone()
#endregion

#region __to inputs all library
import clr
import sys 
import System   
import math

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *
from Autodesk.DesignScript.Geometry import Line, Point

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*

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
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

#region ___pull one point to curve
def pullPointToCurve(inPoint, inCurve):
    pulledPoint = inCurve.Project(inPoint).XYZPoint #pull the branch curve StartPoint to the main curve
    return pulledPoint
#endregion

#region ___pull end point of pipe to its perpendicular pipe
def pullPointToCurve(mPipe, bPipe):
    K = 304.8
    mPCurve = mPipe.Location.Curve
    mPCurveSP = mPCurve.GetEndPoint(0)
    diam = mPipe.Diameter
    bPCurve = bPipe.Location.Curve
    bPCurveSP = bPCurve.GetEndPoint(0)
    bPCurveEP = bPCurve.GetEndPoint(1)

    pMid = mPCurve.Project(bPCurveSP).XYZPoint
    len1 = pMid.DistanceTo(mPCurveSP)
    len2 = len1 - diam/2
    if len2 < 0:
        pMid = mPCurve.Project(bPCurveEP).XYZPoint
    dynMP = Autodesk.DesignScript.Geometry.Point.ByCoordinates(pMid.X*K, pMid.Y*K, pMid.Z*K) 
    return dynMP
#endregion

#region ___generate dynamo point and line from revit point and line
p3 = Autodesk.DesignScript.Geometry.Point.ByCoordinates(pMid.X*304.8, pMid.Y*304.8, pMid.Z*304.8)
tempLine1 = Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(p1,p2)
#endregion

#region ___get 3 closest connectors of 3 pipe
def NearestConnector(ConnectorSet, curCurve):
    MinLength = float("inf")
    result = None  # Initialize result to None
    for n in ConnectorSet:
        distance = curCurve.Location.Curve.Distance(n.Origin)
        if distance < MinLength:
            MinLength = distance
            result = n
    return result

def closetConn(Pipe1, Pipe2, Pipe3):
    connectors1 = list(Pipe1.ConnectorManager.Connectors.GetEnumerator())
    connectors2 = list(Pipe2.ConnectorManager.Connectors.GetEnumerator())
    connectors3 = list(Pipe3.ConnectorManager.Connectors.GetEnumerator())
    Connector1 = NearestConnector(connectors1, Pipe2)
    Connector2 = NearestConnector(connectors2, Pipe1)
    Connector3 = NearestConnector(connectors3, Pipe1)
    return Connector1, Connector2, Connector3
#endregion

#region ___create new Tee from 3 pipe
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
    def closetConn(Pipe1, Pipe2, Pipe3):
        connectors1 = list(Pipe1.ConnectorManager.Connectors.GetEnumerator())
        connectors2 = list(Pipe2.ConnectorManager.Connectors.GetEnumerator())
        connectors3 = list(Pipe3.ConnectorManager.Connectors.GetEnumerator())
        Connector1 = NearestConnector(connectors1, Pipe2)
        Connector2 = NearestConnector(connectors2, Pipe1)
        Connector3 = NearestConnector(connectors3, Pipe1) 
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

#region ___create new elbow
def CreateElbow(doc, pipes):
	connectors = []
	elbowList = []
	try:
		for pipe in pipes:
			connectors.append(pipe.ConnectorManager.Connectors)
	except:
		pass

	TransactionManager.Instance.EnsureInTransaction(doc)
	
	try:
		for set1, set2 in zip(connectors[:-1], connectors[1:]):
			conStart = list(set1)
			conEnd = list(set2)
			d1 = conStart[0].Origin.DistanceTo(conEnd[0].Origin)*conStart[0].Origin.DistanceTo(conEnd[1].Origin)
			d2 = conStart[1].Origin.DistanceTo(conEnd[0].Origin)*conStart[1].Origin.DistanceTo(conEnd[1].Origin)
			i = 0
			if d1 > d2:
				i = 1
	
			d3 = conEnd[0].Origin.DistanceTo(conStart[0].Origin)*conEnd[0].Origin.DistanceTo(conStart[1].Origin)
			d4 = conEnd[1].Origin.DistanceTo(conStart[0].Origin)*conEnd[1].Origin.DistanceTo(conStart[1].Origin)
			j = 0
			if d3 > d4:
				j = 1
	
			fitting = doc.Create.NewElbowFitting(conStart[i],conEnd[j])
			elbowList.append(fitting)
	except:
		pass

	return elbowList
#endregion

#region ___filter all pipe tags in Doc
def getAllPipeTags(doc):
    pipeTags = []
    tagCollector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeTags).WhereElementIsElementType().ToElements()
    pipeTags.append(i for i in tagCollector)
    return pipeTags
#endregion

#region ___get all pipes in Activeview
def allPipesInActiveView():
	pipesList = []
	pipesCollector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
	pipesList.append(i for i in pipesCollector)
	return pipesList
allPipes = allPipesInActiveView()

#endregion



