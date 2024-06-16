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
		systemName = system.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString()
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
def allCateInPJ(doc):
    categories = doc.Settings.Categories
    modelCate = []
    for c in categories:
        if c.CategoryType == CategoryType.Model and c.CanAddSubcategory:
            modelCate.append(Revit.Elements.Category.ById(c.Id.IntegerValue))
    return modelCate
#endregion

#region ___to Define list/unwrap list functions
def uwList(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)
#endregion

#region ___to retrieve selected pipe type, piping system, diameter and reference level
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

#region ___pull end point of BPipe to mPipe
def pullPointToCurve(mPipe, bPipe):
    #K = 304.8
    mPCurve = mPipe.Location.Curve
    nearestConOfBPipe =  closetConn(mPipe, bPipe)
    pMid = mPCurve.Project(nearestConOfBPipe).XYZPoint.ToPoint()
    #dynMP = Autodesk.DesignScript.Geometry.Point.ByCoordinates(pMid.X*K, pMid.Y*K, pMid.Z*K) 
    return pMid

def NearestConnector(ConnectorSet, curCurve):
    MinLength = float("inf")
    result = None  # Initialize result to None
    for n in ConnectorSet:
        distance = curCurve.Location.Curve.Distance(n.Origin)
        if distance < MinLength:
            MinLength = distance
            result = n
    return result

def closetConn(mPipe, bPipe):
    connectors1 = list(bPipe.ConnectorManager.Connectors.GetEnumerator())
    Connector1 = NearestConnector(connectors1, mPipe)
    XYZconn	= Connector1.Origin
    return XYZconn
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


#endregion

#region ___to pick multi points in floors plane or section planes
def pickPoints(doc):
	TransactionManager.Instance.EnsureInTransaction(doc)
	activeView = doc.ActiveView
	iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
	sketchPlane = SketchPlane.Create(doc, iRefPlane)
	doc.ActiveView.SketchPlane = sketchPlane
	condition = True
	pointsList = []
	dynPList = []
	n = 0
	msg = "Pick Points on Current Section plane, hit ESC when finished."
	TaskDialog.Show("^------^", msg)
	while condition:
		try:
			pt=uidoc.Selection.PickPoint()
			n += 1
			pointsList.append(pt)				
		except :
			condition = False
	doc.Delete(sketchPlane.Id)	
	for j in pointsList:
		dynP = j.ToPoint()
		dynPList.append(dynP)
	TransactionManager.Instance.TransactionTaskDone()			
	return dynPList, pointsList
#endregion

#region ___to pick objects
from Autodesk.Revit.UI.Selection import ObjectType
def pickObjects():
    refs = uidoc.Selection.PickObjects(ObjectType.Element)
    return  [doc.GetElement(i.ElementId) for i in refs], [ref for ref in refs]
#endregion

#region ___convert from Revit to Dynamo
//Elements
Element.ToDSType(bool);
//Geometry
XYZ.ToPoint() > Point
XYZ.ToVector() > Vector
Point.ToProtoType() > Point
List<XYZ>.ToPoints() > List<Point>
UV.ToProtoType() > UV
Curve.ToProtoType() > Curve
CurveArray.ToProtoType() > PolyCurve
PolyLine.ToProtoType() > PolyCurve
Plane.ToPlane() > Plane
Solid.ToProtoType() > Solid
Mesh.ToProtoType() > Mesh
IEnumerable<Mesh>.ToProtoType() > Mesh[]
Face.ToProtoType() > IEnumerable<Surface>
Transform.ToCoordinateSystem() > CoordinateSystem
BoundingBoxXYZ.ToProtoType() > BoundingBox
#endregion

#region ___convert from Dynamo to Revit
//Elements
Element.InternalElement
//Geometry
Point.ToRevitType() > XYZ
Vector.ToRevitType() > XYZ
Plane.ToPlane() > Plane
List<Point>.ToXyz() > List<XYZ>
Curve.ToRevitType() > Curve
PolyCurve.ToRevitType() > CurveLoop
Surface.ToRevitType() > IList<GeometryObject>
Solid.ToRevitType() > IList<GeometryObject>
Mesh.ToRevitType() > IList<GeometryObject>
CoordinateSystem.ToTransform() > Transform
CoordinateSystem.ToRevitBoundingBox() > BoundingBoxXYZ
BoundingBox.ToRevitType() > BoundingBoxXYZ
#endregion

#region ___change input to List
def toList(input):
      if isinstance(input, list):
            return UnwrapElement(input)
      else:
            return [UnwrapElement(input)]
#endregion

#region ___to flatten list
##need to import "collections" library
def flatten(l):
    for el in l:
        if isinstance(el, collections.Iterable) and not isinstance(el, basestring):
            for sub in flatten(el):
                yield sub
        else:
            yield el
#endregion

#region ___to check point in between line or not
def pointOnLine(p,a,b):
	if not (p-a).CrossProduct(b-a).IsAlmostEqualTo(XYZ.Zero):
		return False
	if a.X!=b.X:
		if a.X<p.X<b.X:
			return True
		if a.X>p.X>b.X:
			return True
	else:
		if a.Y<p.Y<b.Y:
			return True
		if a.Y>p.Y>b.Y:
			return True
	return False
#endregion

#region ___to retrieve vector of pipe
def getDirectionVector(pipe):
    curve = pipe.Location.Curve
    directionVector = curve.GetEndPoint(1) - curve.GetEndPoint(0)
    return directionVector.Normalize()
#endregion

#region ___to calculate the angle between two pipe
#need to import "from math import acos, degrees"
def getDirectionVector(pipe):
    curve = pipe.Location.Curve
    directionVector = curve.GetEndPoint(1) - curve.GetEndPoint(0)
    return directionVector.Normalize()

def calculateAngleBetweenPipes(mPipe, bPipe):
    curve1 = mPipe.Location.Curve
    v1 = curve1.GetEndPoint(1)-curve1.GetEndPoint(0)
    dir1 = v1.Normalize()
    curve2 = bPipe.Location.Curve
    v2 = curve2.GetEndPoint(1)-curve2.GetEndPoint(0)
    dir2 = v2.Normalize()
    
    dotProduct = dir1.DotProduct(dir2)
    mag1 = dir1.GetLength()
    mag2 = dir2.GetLength()
    cosAngle = dotProduct / (mag1 * mag2)
    angleRadians = acos(cosAngle)
    angleDegrees = degrees(angleRadians)
    angle1 = 180-angleDegrees
    minAngle = 0
    maxAngle = 0
    if angle1 < angleDegrees:
          minAngle = angle1
          maxAngle = angleDegrees
    else: 
          minAngle = angleDegrees
          maxAngle = angle1
    return minAngle,maxAngle
    
#endregion

#region ___to generate plane from point and line
def planeFromPointAndLine(point, line):
    direction = line.Direction
    up_vector = XYZ.BasisZ
    normal = direction.CrossProduct(up_vector).Normalize()
    plane = Plane.CreateByNormalAndOrigin(normal, point) 
    return plane
#endregion 

#region ___to generate plane from 3 points
def create_plane_from_three_points(point1, point2, point3):
    vector1 = point2 - point1
    vector2 = point3 - point1
    normal = vector1.CrossProduct(vector2).Normalize()
    plane = Plane.CreateByNormalAndOrigin(normal, point1)  
    return plane
#endregion

#region ___to offset the point along the vector
def offsetPointAlongVector(point, vector, offsetDistance):
    direction = vector.Normalize()
    scaledVector = direction.Multiply(offsetDistance)
    offsetPoint = point.Add(scaledVector)
    return offsetPoint
#endregion

#region ___to offset the point along the vector
def offsetPointAlongVector(point, vector, offsetDistance):
    point = point.ToPoint()
    direction = vector.ToVector().Normalize()
    scaledVector = direction.Multiply(offsetDistance)
    offsetPoint = point.ToRevitType().Add(scaledVector)
    return offsetPoint
#endregion

#region ___to short the connector of bPipe and retreive the vector
def NearestConnector(ConnectorSet, curCurve):
    MinLength = float("inf")
    result = None  # Initialize result to None
    for n in ConnectorSet:
        distance = curCurve.Location.Curve.Distance(n.Origin)
        if distance < MinLength:
            MinLength = distance
            result = n
    return result

def closetConn(mPipe, bPipe):
    connLst = []
    connectors = list(bPipe.ConnectorManager.Connectors.GetEnumerator())
    nearConn1 = [NearestConnector(connectors, mPipe)]
    farConn1  = [conn for conn in connectors if conn not in nearConn1]
    shortedConnLst = [ele for ele in nearConn1] + [ele for ele in farConn1]
    connLst = [conn.Origin.ToPoint() for conn in shortedConnLst]
    vector = connLst[1].ToRevitType()-connLst[0].ToRevitType()
    return Flatten(connLst), vector.ToVector()
#endregion

#region ______to create pipe from list points
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
#endregion

#region ______to connect pipes by Elbows
def createElbows(pipes):
    fittings = []
    connectors = {}
    connlist = []    
    for pipe in pipes:
        conns = pipe.ConnectorManager.Connectors
        for conn in conns:
            if conn.IsConnected:
                continue
            connectors[conn] = None
            connlist.append(conn)
    for k in connectors.keys():
        mindist = 1000000
        closest = None
        for conn in connlist:
            if conn.Owner.Id.Equals(k.Owner.Id):
                continue
            dist = k.Origin.DistanceTo(conn.Origin)
            if dist < mindist:
                mindist = dist
                closest = conn
        if mindist > margin:
            continue
        connectors[k] = closest
        connlist.remove(closest)
        try:
            del connectors[closest]
        except:
            pass
    for k,v in connectors.items():
        TransactionManager.Instance.EnsureInTransaction(doc)		
        try:
            fitting = doc.Create.NewElbowFitting(k,v)
            fittings.append(fitting.ToDSType(False))
        except:
            pass
        TransactionManager.Instance.TransactionTaskDone()
    return fittings
    #endregion

#region ____to get all pipe accessories from in active view
def getAllPipeAccessoriesInActiveView(doc):
    activeViewId = doc.ActiveView.Id
    collector = FilteredElementCollector(doc, activeViewId).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType()
    pipeAccessories = []
    for element in collector:
        pipeAccessories.append(element)
    return pipeAccessories
#endregion

def uniEle(inList):
    uniLst = []
    for ele in inList:
        if ele not in uniLst:
            uniLst.append(ele)
    return uniLst


def getFittingsName(fittings):
	elbows = []
	elbowsName = []
	tees = []
	teesName = []
	reducers = []
	reducersName = []

	for fitting in fittings:
		name = fitting.Symbol.LookupParameter('Family Name').AsString()
		if 'Elbow' in name: 
			elbows.append(fitting)
			elbowsName.append(name)
		if 'Tee' in name : 
			tees.append(fitting)
			teesName.append(name)
		if 'Reducer' in name : 
			reducers.append(fitting)
			reducersName.append(name)

	return [elbows,elbowsName], [tees,teesName], [reducers,reducersName]


def getConnectTo(connector):
    elemOwner = connector.Owner
    for refCon in connector.AllRefs:
        if  refCon.Owner.Id != elemOwner.Id and refCon.ConnectorType == ConnectorType.End:
            return refCon
    return None 


def duplicateColumns(baseColumn, bList, hList):
    newColumns = []
    sizeList = []
    ss = []
    for b, h in zip(bList, hList):
        size = "{} x {}mm".format(int(b), int(h))
        sizeList.append(size)
    TransactionManager.Instance.EnsureInTransaction(doc)
    try:
        for b, h, size in zip(bList, hList, sizeList):
            idList = List[ElementId]([c.Id for c in baseColumn])
            copyIds = ElementTransformUtils.CopyElements(doc, idList, doc, Transform.Identity, CopyPasteOptions())
            for id in copyIds:
                newCol = doc.GetElement(id)
                newCol.Name = size
                bParam = newCol.LookupParameter('b')
                hParam = newCol.LookupParameter('h')
                if bParam and hParam:
                    bParam.Set(b/304.8)
                    hParam.Set(h/304.8)
                newColumns.append(newCol)
    except Exception as e:
        pass
    TransactionManager.Instance.TransactionTaskDone()
    return newColumns

def chunkList(inputList, chunkSize):
    """Chop a list into chunks of specified size."""
    return [inputList[i:i+chunkSize] for i in range(0, len(inputList), chunkSize)]

def centerOf4Points(points):
    centerPoint = XYZ
    try:
        if len(points) == 4:
           for pt in points:
               sumX += pt.X 
               sumY += pt.Y 
               sumZ += pt.Z
        centerX = sumX /4
        centerY = sumY /4
        centerZ = sumZ /4
        centerPoint = XYZ(centerX, centerY, centerZ)
    except Exception as e:
        pass
    return centerPoint