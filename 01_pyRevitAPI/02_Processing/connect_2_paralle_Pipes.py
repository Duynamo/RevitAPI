"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

from math import sin,cos, tan
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
    paramDiameter = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8
    
    paramPipeTypeId = p.GetTypeId()
    paramPipeType = doc.GetElement(paramPipeTypeId)
    pipeTypeName = paramPipeType.LookupParameter("Type Name").AsString()
	
    paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
    paramPipingSystem = doc.GetElement(paramPipingSystemId)
    pipingSystemName = paramPipingSystem.LookupParameter("System Classification").AsValueString()

    paramLevelId = p.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
    paramLevel = doc.GetElement(paramLevelId)

    return paramPipingSystem, paramPipeType, paramLevel, paramDiameter
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
                connectors1 = list(conn_manager1.Connectors)
                connectors2 = list(conn_manager2.Connectors)               
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
        sorted_connectors = [closest_conn]
        for conn in branch_connectors:
            if conn != closest_conn:
                sorted_connectors.append(conn)
                break
        return closest_conn, closest_xyz, sorted_connectors
    except Exception as e:
        return None, None

from Autodesk.Revit.DB import XYZ
import math

def find_point_C(A, B, D, angle_degrees, distance_AC=None):
    """
    Tìm tọa độ điểm C trong tam giác ABC (3D), biết A, B, D và góc BAC.
    Đảm bảo C thuộc mặt phẳng ABD.
    
    Args:
        A (XYZ): Tọa độ điểm A
        B (XYZ): Tọa độ điểm B
        D (XYZ): Tọa độ điểm D (đồng phẳng với A, B)
        angle_degrees (float): Góc BAC (độ)
        distance_AC (float, optional): Khoảng cách AC, mặc định là |AB|
    
    Returns:
        list of XYZ: Danh sách [C1, C2] (hai điểm C thỏa mãn, hoặc rỗng nếu lỗi)
    """
    try:
        # Chuyển góc sang radian
        angle_radians = math.radians(angle_degrees)
        
        # Vectơ AB
        AB = XYZ(B.X - A.X, B.Y - A.Y, B.Z - A.Z)
        AB_length = AB.GetLength()
        if AB_length == 0:
            return []
        
        # Vectơ AD
        AD = XYZ(D.X - A.X, D.Y - A.Y, D.Z - A.Z)
        
        # Kiểm tra đồng phẳng (AB và AD không đồng tuyến)
        N = AB.CrossProduct(AD)  # Pháp tuyến mặt phẳng ABD
        if N.IsZeroLength():
            return []
        
        # Đặt khoảng cách AC
        d = distance_AC if distance_AC is not None else AB_length
        
        # Tính điểm D' trên AB
        cos_angle = math.cos(angle_radians)
        D_x = d * cos_angle * (AB.X / AB_length)
        D_y = d * cos_angle * (AB.Y / AB_length)
        D_z = d * cos_angle * (AB.Z / AB_length)
        D_prime = XYZ(D_x, D_y, D_z)
        
        # Bán kính đường tròn
        sin_angle = math.sin(angle_radians)
        radius = d * sin_angle
        
        # Tìm vectơ U1' trong mặt phẳng ABD
        # U1' = AD - thành phần của AD song song với AB
        AB_normalized = AB.Normalize()
        AD_parallel = AB_normalized.Multiply(AD.DotProduct(AB_normalized))
        U1_prime = (AD - AD_parallel).Normalize()
        if U1_prime.IsZeroLength():
            return []
        
        # Tính C1, C2 trong mặt phẳng ABD
        C1_local = D_prime + U1_prime.Multiply(radius)
        C2_local = D_prime + U1_prime.Multiply(-radius)
        
        # Dịch về hệ gốc
        C1 = XYZ(A.X + C1_local.X, A.Y + C1_local.Y, A.Z + C1_local.Z)
        C2 = XYZ(A.X + C2_local.X, A.Y + C2_local.Y, A.Z + C2_local.Z)
        
        # Kiểm tra góc (tùy chọn, để debug)
        AC1 = XYZ(C1.X - A.X, C1.Y - A.Y, C1.Z - A.Z)
        AC2 = XYZ(C2.X - A.X, C2.Y - A.Y, C2.Z - A.Z)
        angle_C1 = math.degrees(math.acos(AB.DotProduct(AC1) / (AB.GetLength() * AC1.GetLength())))
        angle_C2 = math.degrees(math.acos(AB.DotProduct(AC2) / (AB.GetLength() * AC2.GetLength())))
        if not (abs(angle_C1 - angle_degrees) < 1e-6 and abs(angle_C2 - angle_degrees) < 1e-6):
            pass
        return [C1, C2]
    
    except Exception as e:
        return []

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
#NOTE: Tại đây ta lấy về XYZ của 2 connector gần nhất của ống nhánh và ống chính với
mPipeCurve = mPipe.Location.Curve
bPipeCurve = bPipe.Location.Curve
bestConn_bPipe = find_closest_connector(mPipe, bPipe)[1]
bestConn_mPipe = find_closest_connector(bPipe, mPipe)[1]

sortedConns_mPipe = find_closest_connector(bPipe, mPipe)[2]
sortedCons_bPipe = find_closest_connector(mPipe, bPipe)[2]

#NOTE: Xác định mặ
# tempPoints = find_point_C(bestConn_mPipe, bestConn_bPipe,sortedCons_bPipe[1].Origin, inAngle)


# sortedXYZCons_bPipe = [c.Origin for c in sortedCons_bPipe]
# nearConn_bPipe = sortedXYZCons_bPipe[0]
# farConn_bPipe = sortedXYZCons_bPipe[1]
# minD = float('inf')
# desPoint = None
# for c in tempPoints:
#     tempDistance = calculateDistance(c, farConn_bPipe)
#     if tempDistance < minD:
#         minD = tempDistance
#         desPoint = c

#need to create temp line to project
tempPoint_bPipe_offsetVector = sortedCons_bPipe[1].Origin - sortedCons_bPipe[0].Origin
tempPoint_bPipe = offsetPointAlongVector(sortedCons_bPipe[1].Origin, 
                                         tempPoint_bPipe_offsetVector,
                                         bPipeCurve.Length)
tempLine = Line.CreateBound(tempPoint_bPipe, sortedCons_bPipe[1].Origin)

projectPoint = bPipeCurve.Project(bestConn_mPipe).XYZPoint
dis = calculateDistance(projectPoint, bestConn_mPipe) #Khoảng cách từ điểm gióng đến best connector main pipe
offsetVector = sortedCons_bPipe[1].Origin -sortedCons_bPipe[0].Origin
if inAngle <= 45:
	angle = 90 - inAngle
	angRad = math.radians(angle)
	offsetDis = dis*tan(angRad)
	tmpPoint = offsetPointAlongVector(projectPoint, offsetVector,offsetDis)
if inAngle > 45 and inAngle <90:
	angle = inAngle
	angRad = math.radians(angle)
	offsetDis = dis*tan(angRad)
	tmpPoint = offsetPointAlongVector(projectPoint, offsetVector,offsetDis)
if inAngle == 90:
	tmpPoint = projectPoint
#NOTE: Tại bước trên ta đã tìm được điểm gióng từ nearConn of main pipe lên
#      trên ống branch pipe ứng với góc input do người nhập
#      Ta đi tạo ống temp ở giữa
tempPipePoints = [bestConn_mPipe.ToPoint(), tmpPoint.ToPoint()]
mPipeParam = getPipeParameter(mPipe)
tempPipe = pipeCreateFromPoints(tempPipePoints, mPipeParam[0],
                                                mPipeParam[1],
                                                mPipeParam[2],
                                                mPipeParam[3])

OUT =  projectPoint, bestConn_mPipe, tmpPoint