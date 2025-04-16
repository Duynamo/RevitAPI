"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

clr.AddReference("ProtoGeometry")

from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB import Line, XYZ, IntersectionResultArray, SetComparisonResult

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI.Selection import ISelectionFilter
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
def projectPointToCurve(mPipe, bPipe):
    mPCurve = mPipe.Location.Curve
    nearestConOfBPipe =  closetConn(mPipe, bPipe)
    pMid = mPCurve.Project(nearestConOfBPipe).XYZPoint.ToPoint()
    return pMid
def NearestConnector(ConnectorSet, curCurve):
    MinLength = float("inf")
    result = None 
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
def pullPointToOxy(XYZpoint):
	newPoint = XYZ(XYZpoint.X, XYZpoint.Y, 0)
	return newPoint
def pullPipeToOxy(pipe):
	pCurve = pipe.Location.Curve
	startPoint = pCurve.GetEndPoint(0)
	endPoint = pCurve.GetEndPoint(1)
	oxySP = pullPointToOxy(startPoint)
	oxyEP = pullPointToOxy(endPoint)
	oxyCurve = Line.CreateBound(oxySP, oxyEP)
	return oxyCurve
def sortNFConn_bPipe(mPipe, bPipe):
    sortConnList = []
    originConns  = []
    tmp = []
    tmp1 = []
    tmp2 = []
    connList = list(bPipe.ConnectorManager.Connectors.GetEnumerator())
    nearConn = NearestConnector(connList, mPipe)
    tmp1.append(nearConn)
    farConn =  [c for c in connList  if c not in tmp1]
    tmp1.append([c for c in connList if c != nearConn])
    sortConnList = flatten(tmp1)
    tmp2.append([c.Origin for c in sortConnList])
    originConns = flatten(tmp2)
    #XYZconn	= nearConn.Origin
    return originConns	
def offsetPointAlongVector(point, vector, offsetDistance):
    direction = vector.Normalize()
    scaledVector = direction.Multiply(offsetDistance)
    offsetPoint = point.Add(scaledVector)
    return offsetPoint
def findIntersection(line1, line2):
    # Prepare the intersection result array
    intersectionResultArray = clr.Reference[Autodesk.Revit.DB.IntersectionResultArray]()
    # Find the intersection of the two lines
    result = line1.Intersect(line2, intersectionResultArray)
    # Check if the result is an overlap
    if result == SetComparisonResult.Overlap and intersectionResultArray.Size > 0:
        intersectionPoint = intersectionResultArray.get_Item(0).XYZPoint
        return intersectionPoint
    else:
        return None
def createPlaneContainingLine(line):
    startPoint = line.GetEndPoint(0)
    direction = line.Direction
    normal = direction.CrossProduct(XYZ.BasisZ)  
    plane = Plane.CreateByNormalAndOrigin(normal, startPoint)
    return plane
def find_intersection(line, plane):
    plane_normal = plane.Normal
    plane_origin = plane.Origin
    line_direction = line.Direction
    line_point = line.GetEndPoint(0)
    # Vector from the plane origin to the line point
    origin_to_point = line_point - plane_origin
    # Dot product of the plane normal and the line direction
    dot_product = plane_normal.DotProduct(line_direction)
    # Check if the line is parallel to the plane
    if abs(dot_product) < 1e-9:
        # The line is parallel to the plane, no intersection
        return None
    # Calculate the parameter t for the line equation
    t = plane_normal.DotProduct(plane_origin - line_point) / dot_product
    # Calculate the intersection point
    intersection_point = line_point + t * line_direction
    return intersection_point
def calculateDistance(pointA, pointB):
    dx = pointA.X - pointB.X
    dy = pointA.Y - pointB.Y
    dz = pointA.Z - pointB.Z
    distance = math.sqrt(dx**2 + dy**2 + dz**2)
    return distance
def pipeCreateFromPoints(desPointsList, sel_pipingSystem, sel_PipeType, sel_Level, diameter):
    lst_Points1 = [i for i in desPointsList]
    lst_Points2 = [i for i in desPointsList[1:]]
    linesList = []
    for pt1, pt2 in zip(lst_Points1,lst_Points2):
        line =  Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(pt1,pt2)
        linesList.append(line)
    firstPoint   = [x.StartPoint for x in linesList]
    secondPoint  = [x.EndPoint for x in linesList]
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

            # param_Length.Set()
            TransactionManager.Instance.TransactionTaskDone()
        except: pass			
    TransactionManager.Instance.TransactionTaskDone()
    return pipe
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
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList
def createElbow(doc, pipes):
    elbowList = []
    bestConns = []
    if not pipes or len(pipes) < 2:
        return elbowList
    TransactionManager.Instance.EnsureInTransaction(doc)
    try:
        # Duyệt qua các cặp pipe liên tiếp
        for pipe1, pipe2 in zip(pipes[:-1], pipes[1:]):
            try:
                # Lấy ConnectorManager của hai pipe
                conn_manager1 = pipe1.ConnectorManager
                conn_manager2 = pipe2.ConnectorManager                
                if not conn_manager1 or not conn_manager2:
                    continue                
                # Lấy tất cả connector
                connectors1 = list(conn_manager1.Connectors.GetEnumerator())
                connectors2 = list(conn_manager2.Connectors.GetEnumerator())               
                if len(connectors1) < 1 or len(connectors2) < 1:
                    continue               
                # Tìm cặp connector gần nhất
                min_distance = float('inf')
                best_conn1 = None
                best_conn2 = None               
                for conn1 in connectors1:
                    for conn2 in connectors2:
                        distance = conn1.Origin.DistanceTo(conn2.Origin)
                        if distance < min_distance:
                            min_distance = distance
                            best_conn1 = conn1
                            best_conn2 = conn2
                            bestConns.append(best_conn1)
                            bestConns.append(best_conn2)

                if best_conn1 is None or best_conn2 is None:
                    continue
                # Kiểm tra khoảng cách hợp lệ (ví dụ: dưới 1 feet)
                if min_distance > 1.0:  # Đơn vị feet
                    continue
                # bestConns.append(best_conn1, best_conn2)
                fitting = doc.Create.NewElbowFitting(best_conn1, best_conn2)
                elbowList.append(fitting)               
            except Exception as e:
                continue
    except Exception as e: pass
    return bestConns

class selectionFilter(ISelectionFilter):
    def __init__(self, category):
          self.category = category
    def AllowElement(self, element):
        if element.Category.Name == self.category:
            return True
        else:
            return False
def pickPipe():
    pipeFilter = selectionFilter('Pipes')
    pipeRef = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter)
    pipe = doc.GetElement(pipeRef.ElementId)
    return pipe
def find_closest_connector(mainPipe, branchPipe):
    """
    Tìm connector gần nhất của ống nhánh đến đường cong của ống chính.
    
    Args:
        mainPipe: Pipe chính (Pipe object)
        branchPipe: Pipe nhánh (Pipe object)
    
    Returns:
        Tuple (Connector, XYZ): Connector gần nhất và tọa độ của nó, hoặc (None, None) nếu lỗi.
    """   
    try:
        # Lấy đường cong của ống chính
        main_curve = mainPipe.Location.Curve
        if not main_curve:
            return None, None
        # Lấy connectors của ống nhánh
        branch_connectors = list(branchPipe.ConnectorManager.Connectors)
        if not branch_connectors:
            return None, None        
        # Tìm connector gần nhất
        min_distance = float("inf")
        closest_conn = None
        closest_xyz = None
        for conn in branch_connectors:
            if not hasattr(conn, "Origin"):
                continue
            # Chiếu điểm connector lên đường cong ống chính
            projection = main_curve.Project(conn.Origin)
            if projection is None:
                continue
            distance = conn.Origin.DistanceTo(projection.XYZPoint)
            if distance < min_distance:
                min_distance = distance
                closest_conn = conn
                closest_xyz = conn.Origin       
        if closest_conn is None:
            return None, None
        return closest_conn, closest_xyz
    except Exception as e:
        return None, None
#endregion
#region input Value
mPipe   = pickPipe()
bPipe	= pickPipe()
##
#region input Angle
class MyForm(Form):
    def __init__(self):
        #NOTE: Để tạo UI có kích thước tương đối so với màn hình làm việc
            #  Ta dùng Screen.PrimaryScreen.WorkingArea
        primary_screen = Screen.PrimaryScreen.WorkingArea
        screen_width = primary_screen.Width // 5
        screen_height = primary_screen.Height // 6
        self.Text = ''
        self.ClientSize = Size(screen_width, screen_height)
        self.Font = System.Drawing.Font("Meiryo UI", 7.5, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self.ForeColor = System.Drawing.Color.Red     
        # Create and add label
        self.label = Label()
        self.label.Text = "ANGLE"
        self.label.Size = Size(screen_width // 2, 50)
        self.label.Location = Point(20, 20)
        self.Controls.Add(self.label)
        # Create and add text box
        self.textBox = TextBox()
        self.textBox.Location = Point(30, 70)
        self.textBox.Size = Size(screen_width // 1.1, 200)
        self.textBox.KeyDown += self.textBox_KeyDown
        self.Controls.Add(self.textBox)
        
        # Create and add OK button
        self.okButton = Button()
        self.okButton.Text = 'OK'
        self.okButton.Size = Size(150,40)
        self.okButton.Location = Point(screen_width - 350, 
                                       screen_height - 70)
        self.okButton.Click += self.okButton_Click
        self.Controls.Add(self.okButton)
        # Create and add Cancel button
        self.cancelButton = Button()
        self.cancelButton.Text = 'CANCLE'
        self.cancelButton.Size = Size(150,40)
        self.cancelButton.Location = Point(screen_width - 180, 
                                           screen_height - 70)
        self.cancelButton.Click += self.cancelButton_Click
        self.Controls.Add(self.cancelButton)
        self.fvcLabel = Label()
        self.fvcLabel.Text = "@FVC"
        self.fvcLabel.Size = Size(150, 40)
        self.fvcLabel.Location = Point(20, screen_height - 50)  # Bottom left corner
        self.Controls.Add(self.fvcLabel)        
        # self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.result = None
        ###
    def textBox_KeyDown(self, sender, event):
            if event.KeyCode == Keys.Enter:
                self.okButton_Click(sender, None)  # Gọi cùng logic của OK
                event.Handled = True        
    def okButton_Click(self, sender, event):
        self.result = self.textBox.Text
        self.DialogResult = DialogResult.OK
        self.Close()
    def cancelButton_Click(self, sender, event):
        self.DialogResult = DialogResult.Cancel
        self.Close()
# Show the form and get the result
form = MyForm()
result = form.ShowDialog()
# Output the result if OK was clicked
if result == DialogResult.OK:
    text_input = form.result
else:
    text_input = None
###
returnAngle = float(text_input)
##
inAngle = returnAngle
#endregion
# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)
#pjPoint = projectPointToCurve(mPipe,bPipe)
nearConnOfBP = closetConn(mPipe,bPipe)
sortNearConnsBP = sortNFConn_bPipe(mPipe,bPipe)

mPipeCurve = mPipe.Location.Curve
bPipeCurve = bPipe.Location.Curve
bPipeCurve_SP = bPipeCurve.GetEndPoint(0)
bPipeCurve_EP = bPipeCurve.GetEndPoint(1)
offsetVector = sortNearConnsBP[0]-sortNearConnsBP[1]
newNearConn = offsetPointAlongVector(sortNearConnsBP[0] , offsetVector, 100)
#create a plane contain main pipe curve and scaled branch pipe curve
tmpLine = Line.CreateBound(newNearConn ,sortNearConnsBP[0])
tmpPlane = createPlaneContainingLine(mPipeCurve)
#get intersection point that On the tmpLine
intersectPoint = find_intersection(tmpLine, tmpPlane)
#create a plane contain tmpLine 
tmpPlane1 = createPlaneContainingLine(tmpLine)
#get intersection point that On the main pipe curve
intersectPoint1 = find_intersection(mPipeCurve, tmpPlane1)
#calculte distance of 2 intersection Point
dis = calculateDistance(intersectPoint, intersectPoint1)
#offset intersectPoint along branch pipe with the user input Angle
tmpPoint1 = intersectPoint
offsetVector1 = nearConnOfBP - tmpPoint1
#calculate offset distance
if inAngle <= 45:
	angle = 90 - inAngle
	angRad = math.radians(angle)
	tmpDis = dis*math.tan(angRad)
	tmpPoint2 = offsetPointAlongVector(intersectPoint, offsetVector1, tmpDis)
if inAngle >45 and inAngle <90:
	angle = inAngle
	angRad = math.radians(angle)
	tmpDis = dis*math.tan(angRad)
	tmpPoint2 = offsetPointAlongVector(intersectPoint, offsetVector1, tmpDis)
if inAngle == 90:
	tmpPoint2 = intersectPoint
if tmpPoint2:
	transVector = tmpPoint2 - sortNearConnsBP[0]
	transform = Autodesk.Revit.DB.Transform.CreateTranslation(transVector)
	ElementTransformUtils.MoveElement(doc, bPipe.Id, transVector)
TransactionManager.Instance.TransactionTaskDone()
pointList   = []
pointList.append(intersectPoint1.ToPoint())
pointList.append(tmpPoint2.ToPoint())
TransactionManager.Instance.EnsureInTransaction(doc)
basePipeParam = getPipeParameter(mPipe)
diamParam = basePipeParam[0][0]
pipingSystemParam = basePipeParam[0][1]
pipeTypeParam = basePipeParam[0][2]
levelParam = basePipeParam[0][3]
tempPipe = pipeCreateFromPoints(pointList, 
                                pipingSystemParam, 
                                pipeTypeParam, 
                                levelParam, 
                                diamParam)
TransactionManager.Instance.TransactionTaskDone()
pipeList = []
pipeList.append(bPipe)
pipeList.append(tempPipe)
TransactionManager.Instance.EnsureInTransaction(doc)
newElbow = createElbow(doc,pipeList )
TransactionManager.Instance.TransactionTaskDone()

TransactionManager.Instance.EnsureInTransaction(doc)
mPipeConns = list(mPipe.ConnectorManager.Connectors)
bestConn1 = find_closest_connector(mPipe, tempPipe)
projection = mPipeCurve.Project(bestConn1[1])
projectionPoint = projection.XYZPoint
listPoint1 = [mPipeConns[0].Origin.ToPoint(), projectionPoint.ToPoint()]
listPoint2 = [mPipeConns[1].Origin.ToPoint(), projectionPoint.ToPoint()]
tempPipe2 = pipeCreateFromPoints(listPoint1, 
                                pipingSystemParam, 
                                pipeTypeParam, 
                                levelParam, 
                                diamParam)
tempPipe3 = pipeCreateFromPoints(listPoint2, 
                                pipingSystemParam, 
                                pipeTypeParam, 
                                levelParam, 
                                diamParam)

tempPipe2Conns = list(tempPipe2.ConnectorManager.Connectors.GetEnumerator())
tempPipe3Conns = list(tempPipe3.ConnectorManager.Connectors.GetEnumerator())
conn1 = None
conn2 = None
for c1 in tempPipe2Conns:
    if c1.Origin.IsAlmostEqualTo(projectionPoint, 0.01):
        conn1 = c1
for c2 in tempPipe3Conns:
    if c2.Origin.IsAlmostEqualTo(projectionPoint, 0.01):
        conn2 = c2    
newTee = doc.Create.NewTeeFitting(conn1, conn2, bestConn1[0])
# doc.Delete(mPipe.Id)
TransactionManager.Instance.TransactionTaskDone()

# OUT = bPipe, tempPipe
OUT = 'Hello World'