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
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app   = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view  = doc.ActiveView
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
    nearestConOfBPipe = closetConn(mPipe, bPipe)
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
    Connector1  = NearestConnector(connectors1, mPipe)
    XYZconn     = Connector1.Origin
    return XYZconn
def pullPointToOxy(XYZpoint):
    newPoint = XYZ(XYZpoint.X, XYZpoint.Y, 0)
    return newPoint
def pullPipeToOxy(pipe):
    pCurve     = pipe.Location.Curve
    startPoint = pCurve.GetEndPoint(0)
    endPoint   = pCurve.GetEndPoint(1)
    oxySP      = pullPointToOxy(startPoint)
    oxyEP      = pullPointToOxy(endPoint)
    oxyCurve   = Line.CreateBound(oxySP, oxyEP)
    return oxyCurve
def sortNFConn_bPipe(mPipe, bPipe):
    tmp1     = []
    connList = list(bPipe.ConnectorManager.Connectors.GetEnumerator())
    nearConn = NearestConnector(connList, mPipe)
    tmp1.append(nearConn)
    tmp1.append([c for c in connList if c != nearConn])
    sortConnList = flatten(tmp1)
    originConns  = flatten([[c.Origin for c in sortConnList]])
    return originConns	
def offsetPointAlongVector(point, vector, offsetDistance):
    direction   = vector.Normalize()
    scaledVector = direction.Multiply(offsetDistance)
    offsetPoint  = point.Add(scaledVector)
    return offsetPoint
def findIntersection(line1, line2):
    intersectionResultArray = clr.Reference[Autodesk.Revit.DB.IntersectionResultArray]()
    result = line1.Intersect(line2, intersectionResultArray)
    if result == SetComparisonResult.Overlap and intersectionResultArray.Size > 0:
        return intersectionResultArray.get_Item(0).XYZPoint
    return None
def createPlaneContainingLine(line):
    startPoint = line.GetEndPoint(0)
    direction  = line.Direction
    normal     = direction.CrossProduct(XYZ.BasisZ)  
    plane      = Plane.CreateByNormalAndOrigin(normal, startPoint)
    return plane
def find_intersection(line, plane):
    plane_normal  = plane.Normal
    plane_origin  = plane.Origin
    line_direction = line.Direction
    line_point    = line.GetEndPoint(0)
    dot_product   = plane_normal.DotProduct(line_direction)
    if abs(dot_product) < 1e-9:
        return None
    t = plane_normal.DotProduct(plane_origin - line_point) / dot_product
    return line_point + t * line_direction
def calculateDistance(pointA, pointB):
    dx = pointA.X - pointB.X
    dy = pointA.Y - pointB.Y
    dz = pointA.Z - pointB.Z
    return math.sqrt(dx**2 + dy**2 + dz**2)
def pipeCreateFromPoints(desPointsList, sel_pipingSystem, sel_PipeType, sel_Level, diameter):
    lst_Points1 = [i for i in desPointsList]
    lst_Points2 = [i for i in desPointsList[1:]]
    linesList   = []
    for pt1, pt2 in zip(lst_Points1, lst_Points2):
        line = Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(pt1, pt2)
        linesList.append(line)
    firstPoint  = [x.StartPoint for x in linesList]
    secondPoint = [x.EndPoint   for x in linesList]
    TransactionManager.Instance.EnsureInTransaction(doc)
    pipe = None
    for i, x in enumerate(firstPoint):
        try:
            levelId    = sel_Level.Id
            sysTypeId  = sel_pipingSystem.Id
            pipeTypeId = sel_PipeType.Id
            diam       = diameter
            pipe = Autodesk.Revit.DB.Plumbing.Pipe.Create(
                doc, sysTypeId, pipeTypeId, levelId,
                x.ToXyz(), secondPoint[i].ToXyz())
            param_Diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
            param_Diameter.SetValueString(diam.ToString())
            TransactionManager.Instance.EnsureInTransaction(doc)
            TransactionManager.Instance.TransactionTaskDone()
        except: pass			
    TransactionManager.Instance.TransactionTaskDone()
    return pipe
def getPipeParameter(p):
    paramDiameter = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8
    paramPipeTypeId   = p.GetTypeId()
    paramPipeType     = doc.GetElement(paramPipeTypeId)
    paramPipingSystemId = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
    paramPipingSystem   = doc.GetElement(paramPipingSystemId)
    paramLevelId = p.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
    paramLevel   = doc.GetElement(paramLevelId)
    pipingSystemName = paramPipingSystem.LookupParameter("System Classification").AsValueString()
    pipeTypeName     = paramPipeType.LookupParameter("Type Name").AsString()
    return [paramDiameter, paramPipingSystem, paramPipeType, paramLevel], \
           [paramDiameter, pipingSystemName, pipeTypeName, paramLevel]

def createElbowBetweenPipes(doc, pipe1, pipe2):
    """
    Tạo elbow giữa 2 pipe, tìm cặp connector gần nhất.
    Trả về fitting hoặc None nếu thất bại.
    """
    try:
        conns1 = list(pipe1.ConnectorManager.Connectors.GetEnumerator())
        conns2 = list(pipe2.ConnectorManager.Connectors.GetEnumerator())
        min_dist   = float('inf')
        best_conn1 = None
        best_conn2 = None
        for c1 in conns1:
            for c2 in conns2:
                d = c1.Origin.DistanceTo(c2.Origin)
                if d < min_dist:
                    min_dist   = d
                    best_conn1 = c1
                    best_conn2 = c2
        if best_conn1 is None or best_conn2 is None:
            return None
        if min_dist > 1.0:
            return None
        fitting = doc.Create.NewElbowFitting(best_conn1, best_conn2)
        return fitting
    except Exception as ex:
        return None

def find_closest_connector(mainPipe, branchPipe):
    try:
        main_curve        = mainPipe.Location.Curve
        branch_connectors = list(branchPipe.ConnectorManager.Connectors)
        min_distance  = float("inf")
        closest_conn  = None
        closest_xyz   = None
        for conn in branch_connectors:
            if not hasattr(conn, "Origin"):
                continue
            projection = main_curve.Project(conn.Origin)
            if projection is None:
                continue
            distance = conn.Origin.DistanceTo(projection.XYZPoint)
            if distance < min_distance:
                min_distance = distance
                closest_conn = conn
                closest_xyz  = conn.Origin       
        return closest_conn, closest_xyz
    except:
        return None, None

def calcMidPoint(intersectPoint1, tmpPoint2, mPipeCurve, inAngle):
    """
    Tính midPoint trên đường vuông góc với ống chính từ intersectPoint1,
    sao cho pipe từ midPoint đến tmpPoint2 có góc = inAngle so với ống chính.

    Sơ đồ:
        intersectPoint1 ---[perp pipe 90 deg]---> midPoint
                                                      |
                                              [angled pipe @ inAngle]
                                                      |
                                                  tmpPoint2 (bPipe connector)
    """
    main_dir_n   = mPipeCurve.Direction.Normalize()
    total_vec    = tmpPoint2 - intersectPoint1

    # Phân tích thành thành phần song song và vuông góc với ống chính
    along_scalar  = total_vec.DotProduct(main_dir_n)
    along_vec     = main_dir_n.Multiply(along_scalar)
    perp_vec_raw  = total_vec - along_vec
    perp_len      = perp_vec_raw.GetLength()

    if perp_len < 1e-6:
        # tmpPoint2 nằm trên trục ống chính, không có thành phần vuông góc
        return intersectPoint1, False

    perp_n_local = perp_vec_raw.Normalize()

    if inAngle >= 90 or abs(along_scalar) < 1e-6:
        # Góc 90 deg: toàn bộ là đoạn vuông góc, không cần pipe góc
        return intersectPoint1.Add(perp_n_local.Multiply(perp_len)), False

    # Tính L1: độ dài đoạn vuông góc
    # Từ midPoint đến tmpPoint2 phải tạo góc inAngle so với ống chính:
    #   tan(inAngle) = (perp_len - L1) / |along_scalar|
    #   => L1 = perp_len - |along_scalar| * tan(inAngle)
    theta_rad = math.radians(inAngle)
    L1        = perp_len - abs(along_scalar) * math.tan(theta_rad)

    # Đảm bảo L1 > 0 và đủ dài để tạo pipe hợp lệ (min 0.05 ft)
    MIN_PIPE_LEN = 0.05
    if L1 < MIN_PIPE_LEN:
        L1 = MIN_PIPE_LEN

    midPoint   = intersectPoint1.Add(perp_n_local.Multiply(L1))
    need_angle = midPoint.DistanceTo(tmpPoint2) > MIN_PIPE_LEN
    return midPoint, need_angle
#endregion

#region ___input: Pick pipes
mPipe = pickPipe() if 'pickPipe' in dir() else None

class selectionFilter(ISelectionFilter):
    def __init__(self, category):
        self.category = category
    def AllowElement(self, element):
        return element.Category.Name == self.category

def pickPipe():
    pipeFilter = selectionFilter('Pipes')
    pipeRef    = uidoc.Selection.PickObject(
        Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter)
    return doc.GetElement(pipeRef.ElementId)

mPipe = pickPipe()
bPipe = pickPipe()
#endregion

#region ___input: Angle UI
class MyForm(Form):
    def __init__(self):
        primary_screen = Screen.PrimaryScreen.WorkingArea
        screen_width   = primary_screen.Width  // 5
        screen_height  = primary_screen.Height // 6
        self.Text        = ''
        self.ClientSize  = Size(screen_width, screen_height)
        self.Font        = System.Drawing.Font(
            "Meiryo UI", 7.5, System.Drawing.FontStyle.Bold,
            System.Drawing.GraphicsUnit.Point, 128)
        self.ForeColor   = System.Drawing.Color.Red

        self.label      = Label()
        self.label.Text = "ANGLE (degrees)"
        self.label.Size = Size(screen_width // 2, 50)
        self.label.Location = Point(20, 20)
        self.Controls.Add(self.label)

        self.textBox          = TextBox()
        self.textBox.Location = Point(30, 70)
        self.textBox.Size     = Size(int(screen_width // 1.1), 200)
        self.textBox.KeyDown += self.textBox_KeyDown
        self.Controls.Add(self.textBox)

        self.okButton          = Button()
        self.okButton.Text     = 'OK'
        self.okButton.Size     = Size(150, 40)
        self.okButton.Location = Point(screen_width - 350, screen_height - 70)
        self.okButton.Click   += self.okButton_Click
        self.Controls.Add(self.okButton)

        self.cancelButton          = Button()
        self.cancelButton.Text     = 'CANCEL'
        self.cancelButton.Size     = Size(150, 40)
        self.cancelButton.Location = Point(screen_width - 180, screen_height - 70)
        self.cancelButton.Click   += self.cancelButton_Click
        self.Controls.Add(self.cancelButton)

        self.fvcLabel          = Label()
        self.fvcLabel.Text     = "@FVC"
        self.fvcLabel.Size     = Size(150, 40)
        self.fvcLabel.Location = Point(20, screen_height - 50)
        self.Controls.Add(self.fvcLabel)
        self.result = None

    def textBox_KeyDown(self, sender, event):
        if event.KeyCode == Keys.Enter:
            self.okButton_Click(sender, None)
            event.Handled = True

    def okButton_Click(self, sender, event):
        self.result       = self.textBox.Text
        self.DialogResult = DialogResult.OK
        self.Close()

    def cancelButton_Click(self, sender, event):
        self.DialogResult = DialogResult.Cancel
        self.Close()

form   = MyForm()
result = form.ShowDialog()
if result == DialogResult.OK:
    text_input = form.result
else:
    text_input = None

inAngle = float(text_input)
#endregion

# ═══════════════════════════════════════════════════════════════
#  MAIN GEOMETRY CALCULATION
# ═══════════════════════════════════════════════════════════════
TransactionManager.Instance.EnsureInTransaction(doc)

nearConnOfBP    = closetConn(mPipe, bPipe)
sortNearConnsBP = sortNFConn_bPipe(mPipe, bPipe)

mPipeCurve    = mPipe.Location.Curve
bPipeCurve    = bPipe.Location.Curve
offsetVector  = sortNearConnsBP[0] - sortNearConnsBP[1]
newNearConn   = offsetPointAlongVector(sortNearConnsBP[0], offsetVector, 100)

# Mặt phẳng chứa ống chính
tmpLine       = Line.CreateBound(newNearConn, sortNearConnsBP[0])
tmpPlane      = createPlaneContainingLine(mPipeCurve)
intersectPoint = find_intersection(tmpLine, tmpPlane)

# Mặt phẳng vuông góc qua ống nhánh
tmpPlane1      = createPlaneContainingLine(tmpLine)
intersectPoint1 = find_intersection(mPipeCurve, tmpPlane1)

# Khoảng cách giữa 2 giao điểm
dis          = calculateDistance(intersectPoint, intersectPoint1)
offsetVector1 = nearConnOfBP - intersectPoint

# ── Tính tmpPoint2 theo góc người dùng nhập ───────────────────
if inAngle <= 45:
    angle   = 90 - inAngle
    angRad  = math.radians(angle)
    tmpDis  = dis * math.tan(angRad)
    tmpPoint2 = offsetPointAlongVector(intersectPoint, offsetVector1, tmpDis)
elif inAngle < 90:
    angle   = inAngle
    angRad  = math.radians(angle)
    tmpDis  = dis * math.tan(angRad)
    tmpPoint2 = offsetPointAlongVector(intersectPoint, offsetVector1, tmpDis)
else:
    tmpPoint2 = intersectPoint

# ── Di chuyển ống nhánh đến vị trí mới ───────────────────────
if tmpPoint2:
    transVector = tmpPoint2 - sortNearConnsBP[0]
    ElementTransformUtils.MoveElement(doc, bPipe.Id, transVector)

TransactionManager.Instance.TransactionTaskDone()

# ═══════════════════════════════════════════════════════════════
#  TÍNH MIDPOINT: điểm nối giữa ống vuông góc và ống góc user
#
#   intersectPoint1 --(perp 90 deg)--> midPoint --(inAngle)--> tmpPoint2
# ═══════════════════════════════════════════════════════════════
midPoint, need_angle_pipe = calcMidPoint(
    intersectPoint1, tmpPoint2, mPipeCurve, inAngle)

# ── Lấy thông số ống chính để tạo ống mới ────────────────────
basePipeParam     = getPipeParameter(mPipe)
diamParam         = basePipeParam[0][0]
pipingSystemParam = basePipeParam[0][1]
pipeTypeParam     = basePipeParam[0][2]
levelParam        = basePipeParam[0][3]

# ═══════════════════════════════════════════════════════════════
#  BƯỚC 1: Tạo ống vuông góc (90 deg từ ống chính)
#          intersectPoint1 --> midPoint
# ═══════════════════════════════════════════════════════════════
pointList_perp = [intersectPoint1.ToPoint(), midPoint.ToPoint()]
TransactionManager.Instance.EnsureInTransaction(doc)
tempPipe_perp = pipeCreateFromPoints(
    pointList_perp, pipingSystemParam, pipeTypeParam, levelParam, diamParam)
TransactionManager.Instance.TransactionTaskDone()

# ═══════════════════════════════════════════════════════════════
#  BƯỚC 2: Tạo ống góc user (inAngle từ ống chính)
#          midPoint --> tmpPoint2
#  Chỉ tạo nếu midPoint khác tmpPoint2
# ═══════════════════════════════════════════════════════════════
tempPipe_angle = None
if need_angle_pipe:
    pointList_angle = [midPoint.ToPoint(), tmpPoint2.ToPoint()]
    TransactionManager.Instance.EnsureInTransaction(doc)
    tempPipe_angle = pipeCreateFromPoints(
        pointList_angle, pipingSystemParam, pipeTypeParam, levelParam, diamParam)
    TransactionManager.Instance.TransactionTaskDone()

# ═══════════════════════════════════════════════════════════════
#  BƯỚC 3: Tạo elbow 90 deg
#          Nối tempPipe_perp -- tempPipe_angle (hoặc bPipe nếu ko có angle pipe)
# ═══════════════════════════════════════════════════════════════
TransactionManager.Instance.EnsureInTransaction(doc)
if need_angle_pipe and tempPipe_angle is not None:
    # Elbow 90 deg giữa ống vuông góc và ống góc user
    elbow_90 = createElbowBetweenPipes(doc, tempPipe_perp, tempPipe_angle)
else:
    # Không có ống góc: nối thẳng tempPipe_perp với bPipe
    elbow_90 = createElbowBetweenPipes(doc, tempPipe_perp, bPipe)
TransactionManager.Instance.TransactionTaskDone()

# ═══════════════════════════════════════════════════════════════
#  BƯỚC 4: Tạo elbow góc user
#          Nối tempPipe_angle -- bPipe
# ═══════════════════════════════════════════════════════════════
if need_angle_pipe and tempPipe_angle is not None:
    TransactionManager.Instance.EnsureInTransaction(doc)
    elbow_user = createElbowBetweenPipes(doc, tempPipe_angle, bPipe)
    TransactionManager.Instance.TransactionTaskDone()

# ═══════════════════════════════════════════════════════════════
#  BƯỚC 5: Tạo tee tại ống chính
#          Nối 2 nửa ống chính với tempPipe_perp
# ═══════════════════════════════════════════════════════════════
TransactionManager.Instance.EnsureInTransaction(doc)

mPipeConns  = list(mPipe.ConnectorManager.Connectors)

# Tìm connector của tempPipe_perp gần ống chính nhất
bestConn1   = find_closest_connector(mPipe, tempPipe_perp)
projection  = mPipeCurve.Project(bestConn1[1])
projectionPoint = projection.XYZPoint

# Tạo 2 nửa ống chính
listPoint1 = [mPipeConns[0].Origin.ToPoint(), projectionPoint.ToPoint()]
listPoint2 = [mPipeConns[1].Origin.ToPoint(), projectionPoint.ToPoint()]

tempPipe2 = pipeCreateFromPoints(
    listPoint1, pipingSystemParam, pipeTypeParam, levelParam, diamParam)
tempPipe3 = pipeCreateFromPoints(
    listPoint2, pipingSystemParam, pipeTypeParam, levelParam, diamParam)

# Tìm connector tại projectionPoint của mỗi nửa
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

# Tạo tee: 2 nửa ống chính + connector của tempPipe_perp
newTee = doc.Create.NewTeeFitting(conn1, conn2, bestConn1[0])

TransactionManager.Instance.TransactionTaskDone()

OUT = 'Done'
