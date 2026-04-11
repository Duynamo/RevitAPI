import clr
import sys
import System
import math

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB.Plumbing import *
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *
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
"""_____"""

#region ___def pickPipe
class selectionFilter(ISelectionFilter):
    def __init__(self, category):
        self.category = category
    
    def AllowElement(self, element):
        if element.Category.Name == self.category:
            return True
        else:
            return False

def pickPipe():
    pipeFilter = selectionFilter("Pipes")
    ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter, "pick Pipe")
    pipe = doc.GetElement(ref.ElementId)
    return pipe
#endregion

#region ___def getPipeLocationCurve
def getPipeLocationCurve(pipe):
    lCurve = pipe.Location.Curve
    return lCurve
#endregion

#region ___def flatten
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList
#endregion

#region ___def divideLineSegment
def divideLineSegment(line, length, startPoint, endPoint):
    points = []
    total_length = line.Length
    direction = (endPoint - startPoint).Normalize()
    current_point = startPoint
    points.append(current_point.ToPoint())
    while (current_point.DistanceTo(startPoint) + length) <= total_length:
        current_point = current_point + direction * length
        points.append(current_point.ToPoint())
    return points
#endregion

#region ____def splitPipeAtPoints
def splitPipeAtPoints(doc, pipe, points):
    newPipes = []
    currentPipe = pipe
    originalPipe = pipe
    TransactionManager.Instance.EnsureInTransaction(doc)
    for point in points:
        pipeLocation = currentPipe.Location
        if isinstance(pipeLocation, LocationCurve):
            pipeCurve = pipeLocation.Curve
            if pipeCurve is not None:
                if is_point_on_curve(pipeCurve, point):
                    newPipeIds = PlumbingUtils.BreakCurve(doc, currentPipe.Id, point)
                    if isinstance(newPipeIds, list):
                        newPipe = doc.GetElement(newPipeIds[0])  # Assuming the first element is the new pipe
                        newPipes.append(newPipe)
                        currentPipe = newPipe
                    else:
                        currentPipe = doc.GetElement(newPipeIds)
                else:
                    # Use the original pipe for splitting
                    newPipeIds = PlumbingUtils.BreakCurve(doc, originalPipe.Id, point)
                    if isinstance(newPipeIds, list):
                        newPipe = doc.GetElement(newPipeIds[0])  # Assuming the first element is the new pipe
                        newPipes.append(newPipe)
                        currentPipe = newPipe
                    else:
                        currentPipe = doc.GetElement(newPipeIds)
    TransactionManager.Instance.TransactionTaskDone()
    return newPipes

# Utility function to check if a point is on a curve
def is_point_on_curve(curve, point):
    projected_point = curve.Project(point)
    return projected_point.Distance < 1e-6
#endregion

class MainForm(Form):
    def __init__(self):
        self.InitializeComponent()
    
    def InitializeComponent(self):
        self._btt_pickPipe = System.Windows.Forms.Button()
        self._grb_sortConn = System.Windows.Forms.GroupBox()
        self._btt_CANCLE = System.Windows.Forms.Button()
        self._btt_SPLIT = System.Windows.Forms.Button()
        self._lb_FVC = System.Windows.Forms.Label()
        self._txb_Length = System.Windows.Forms.TextBox()
        self._lb_Length = System.Windows.Forms.Label()
        self._txb_K = System.Windows.Forms.TextBox()
        self._lb_splitNumber = System.Windows.Forms.Label()
        self._grb_inputData = System.Windows.Forms.GroupBox()
        self._rbt_pX = System.Windows.Forms.RadioButton()
        self._rbt_pY = System.Windows.Forms.RadioButton()
        self._rbt_pZ = System.Windows.Forms.RadioButton()
        self._rbt_sortByMax = System.Windows.Forms.RadioButton()
        self._rbt_sortByMin = System.Windows.Forms.RadioButton()
        self._grb_MinMax = System.Windows.Forms.GroupBox()
        self._grb_sortConn.SuspendLayout()
        self._grb_inputData.SuspendLayout()
        self._grb_MinMax.SuspendLayout()
        self.SuspendLayout()
        
        # Thiết lập kích thước form
        primary_screen = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea
        screen_width = primary_screen.Width // 4
        screen_height = primary_screen.Height // 2
        self.ClientSize = System.Drawing.Size(screen_width, screen_height)
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        
        # btt_pickPipe
        self._btt_pickPipe.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_pickPipe.ForeColor = System.Drawing.Color.Red
        self._btt_pickPipe.Location = System.Drawing.Point(int(screen_width * 0.3), int(screen_height * 0.05))
        self._btt_pickPipe.Name = "btt_pickPipe"
        self._btt_pickPipe.Size = System.Drawing.Size(int(screen_width * 0.35), int(screen_height * 0.1))
        self._btt_pickPipe.TabIndex = 0
        self._btt_pickPipe.Text = "PICK PIPE"
        self._btt_pickPipe.UseVisualStyleBackColor = True
        self._btt_pickPipe.Click += self.Btt_pickPipeClick
        
        # grb_inputData
        self._grb_inputData.Controls.Add(self._lb_Length)
        self._grb_inputData.Controls.Add(self._txb_Length)
        self._grb_inputData.Controls.Add(self._lb_splitNumber)
        self._grb_inputData.Controls.Add(self._txb_K)
        self._grb_inputData.Cursor = System.Windows.Forms.Cursors.Default
        self._grb_inputData.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._grb_inputData.Location = System.Drawing.Point(int(screen_width * 0.03), int(screen_height * 0.2))
        self._grb_inputData.Name = "grb_inputData"
        self._grb_inputData.RightToLeft = System.Windows.Forms.RightToLeft.No
        self._grb_inputData.Size = System.Drawing.Size(int(screen_width * 0.94), int(screen_height * 0.25))
        self._grb_inputData.TabIndex = 6
        self._grb_inputData.TabStop = False
        self._grb_inputData.Text = "Input"
        
        # lb_Length
        self._lb_Length.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._lb_Length.Location = System.Drawing.Point(int(screen_width * 0.05), int(screen_height * 0.05))
        self._lb_Length.Name = "lb_Length"
        self._lb_Length.Size = System.Drawing.Size(int(screen_width * 0.3), int(screen_height * 0.05))
        self._lb_Length.TabIndex = 2
        self._lb_Length.Text = "Length(mm):"
        self._lb_Length.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        
        # txb_Length
        self._txb_Length.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._txb_Length.Location = System.Drawing.Point(int(screen_width * 0.05), int(screen_height * 0.1))
        self._txb_Length.Name = "txb_Length"
        self._txb_Length.Size = System.Drawing.Size(int(screen_width * 0.35), int(screen_height * 0.07))
        self._txb_Length.TabIndex = 3
        self._txb_Length.Text = '3000'
        
        # lb_splitNumber
        self._lb_splitNumber.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._lb_splitNumber.Location = System.Drawing.Point(int(screen_width * 0.55), int(screen_height * 0.05))
        self._lb_splitNumber.Name = "lb_splitNumber"
        self._lb_splitNumber.Size = System.Drawing.Size(int(screen_width * 0.15), int(screen_height * 0.05))
        self._lb_splitNumber.TabIndex = 4
        self._lb_splitNumber.Text = "K:"
        self._lb_splitNumber.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        
        # txb_K
        self._txb_K.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._txb_K.Location = System.Drawing.Point(int(screen_width * 0.55), int(screen_height * 0.1))
        self._txb_K.Name = "txb_K"
        self._txb_K.Size = System.Drawing.Size(int(screen_width * 0.35), int(screen_height * 0.07))
        self._txb_K.TabIndex = 5
        self._txb_K.TextChanged += self.Txb_KTextChanged
        self._txb_K.Text = '1000'
        
        # grb_sortConn
        self._grb_sortConn.Controls.Add(self._rbt_pX)
        self._grb_sortConn.Controls.Add(self._rbt_pY)
        self._grb_sortConn.Controls.Add(self._rbt_pZ)
        self._grb_sortConn.Cursor = System.Windows.Forms.Cursors.Default
        self._grb_sortConn.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._grb_sortConn.Location = System.Drawing.Point(int(screen_width * 0.03), int(screen_height * 0.5))
        self._grb_sortConn.Name = "grb_sortConn"
        self._grb_sortConn.Size = System.Drawing.Size(int(screen_width * 0.45), int(screen_height * 0.3))
        self._grb_sortConn.TabIndex = 1
        self._grb_sortConn.TabStop = False
        self._grb_sortConn.Text = "Sort Connector by"
        self._grb_sortConn.UseCompatibleTextRendering = True
        
        # rbt_pX
        self._rbt_pX.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._rbt_pX.ForeColor = System.Drawing.Color.Red
        self._rbt_pX.Location = System.Drawing.Point(int(screen_width * 0.05), int(screen_height * 0.05))
        self._rbt_pX.Name = "rbt_pX"
        self._rbt_pX.Size = System.Drawing.Size(int(screen_width * 0.35), int(screen_height * 0.07))
        self._rbt_pX.TabIndex = 7
        self._rbt_pX.TabStop = True
        self._rbt_pX.Text = "pX"
        self._rbt_pX.UseVisualStyleBackColor = True
        self._rbt_pX.CheckedChanged += self.Rbt_pXCheckedChanged
        self._rbt_pX.Checked = True
        
        # rbt_pY
        self._rbt_pY.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._rbt_pY.ForeColor = System.Drawing.Color.Red
        self._rbt_pY.Location = System.Drawing.Point(int(screen_width * 0.05), int(screen_height * 0.12))
        self._rbt_pY.Name = "rbt_pY"
        self._rbt_pY.Size = System.Drawing.Size(int(screen_width * 0.35), int(screen_height * 0.07))
        self._rbt_pY.TabIndex = 7
        self._rbt_pY.TabStop = True
        self._rbt_pY.Text = "pY"
        self._rbt_pY.UseVisualStyleBackColor = True
        self._rbt_pY.CheckedChanged += self.Rbt_pYCheckedChanged
        
        # rbt_pZ
        self._rbt_pZ.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._rbt_pZ.ForeColor = System.Drawing.Color.Red
        self._rbt_pZ.Location = System.Drawing.Point(int(screen_width * 0.05), int(screen_height * 0.19))
        self._rbt_pZ.Name = "rbt_pZ"
        self._rbt_pZ.Size = System.Drawing.Size(int(screen_width * 0.35), int(screen_height * 0.07))
        self._rbt_pZ.TabIndex = 7
        self._rbt_pZ.TabStop = True
        self._rbt_pZ.Text = "pZ"
        self._rbt_pZ.UseVisualStyleBackColor = True
        self._rbt_pZ.CheckedChanged += self.Rbt_pZCheckedChanged
        
        # grb_MinMax
        self._grb_MinMax.Controls.Add(self._rbt_sortByMin)
        self._grb_MinMax.Controls.Add(self._rbt_sortByMax)
        self._grb_MinMax.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._grb_MinMax.Location = System.Drawing.Point(int(screen_width * 0.52), int(screen_height * 0.5))
        self._grb_MinMax.Name = "grb_MinMax"
        self._grb_MinMax.Size = System.Drawing.Size(int(screen_width * 0.45), int(screen_height * 0.3))
        self._grb_MinMax.TabIndex = 8
        self._grb_MinMax.TabStop = False
        self._grb_MinMax.Text = "MinMax"
        
        # rbt_sortByMin
        self._rbt_sortByMin.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._rbt_sortByMin.ForeColor = System.Drawing.Color.Fuchsia
        self._rbt_sortByMin.Location = System.Drawing.Point(int(screen_width * 0.05), int(screen_height * 0.05))
        self._rbt_sortByMin.Name = "rbt_sortByMin"
        self._rbt_sortByMin.Size = System.Drawing.Size(int(screen_width * 0.35), int(screen_height * 0.07))
        self._rbt_sortByMin.TabIndex = 7
        self._rbt_sortByMin.TabStop = True
        self._rbt_sortByMin.Text = "Sort By Min?"
        self._rbt_sortByMin.UseVisualStyleBackColor = True
        self._rbt_sortByMin.CheckedChanged += self.Rbt_sortByMaxCheckedChanged
        self._rbt_sortByMin.Checked = True
        
        # rbt_sortByMax
        self._rbt_sortByMax.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._rbt_sortByMax.ForeColor = System.Drawing.Color.Fuchsia
        self._rbt_sortByMax.Location = System.Drawing.Point(int(screen_width * 0.05), int(screen_height * 0.12))
        self._rbt_sortByMax.Name = "rbt_sortByMax"
        self._rbt_sortByMax.Size = System.Drawing.Size(int(screen_width * 0.35), int(screen_height * 0.07))
        self._rbt_sortByMax.TabIndex = 7
        self._rbt_sortByMax.TabStop = True
        self._rbt_sortByMax.Text = "Sort By Max?"
        self._rbt_sortByMax.UseVisualStyleBackColor = True
        self._rbt_sortByMax.CheckedChanged += self.Rbt_sortByMaxCheckedChanged
        
        # btt_SPLIT
        self._btt_SPLIT.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_SPLIT.ForeColor = System.Drawing.Color.Red
        self._btt_SPLIT.Location = System.Drawing.Point(int(screen_width * 0.45), int(screen_height * 0.85))
        self._btt_SPLIT.Name = "btt_SPLIT"
        self._btt_SPLIT.Size = System.Drawing.Size(int(screen_width * 0.25), int(screen_height * 0.1))
        self._btt_SPLIT.TabIndex = 0
        self._btt_SPLIT.Text = "SPLIT"
        self._btt_SPLIT.UseVisualStyleBackColor = True
        self._btt_SPLIT.Click += self.Btt_SPLITClick
        
        # btt_CANCLE
        self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
        self._btt_CANCLE.Location = System.Drawing.Point(int(screen_width * 0.72), int(screen_height * 0.85))
        self._btt_CANCLE.Name = "btt_CANCLE"
        self._btt_CANCLE.Size = System.Drawing.Size(int(screen_width * 0.25), int(screen_height * 0.1))
        self._btt_CANCLE.TabIndex = 0
        self._btt_CANCLE.Text = "CANCLE"
        self._btt_CANCLE.UseVisualStyleBackColor = True
        self._btt_CANCLE.Click += self.Btt_CANCLEClick
        
        # lb_FVC
        self._lb_FVC.Font = System.Drawing.Font("Meiryo UI", 4.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._lb_FVC.ForeColor = System.Drawing.Color.Blue
        self._lb_FVC.Location = System.Drawing.Point(int(screen_width * 0.05), int(screen_height * 0.9))
        self._lb_FVC.Name = "lb_FVC"
        self._lb_FVC.Size = System.Drawing.Size(int(screen_width * 0.15), int(screen_height * 0.05))
        self._lb_FVC.TabIndex = 2
        self._lb_FVC.Text = "@FVC"
        self._lb_FVC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        
        # MainForm
        self.Controls.Add(self._grb_MinMax)
        self.Controls.Add(self._grb_inputData)
        self.Controls.Add(self._lb_FVC)
        self.Controls.Add(self._grb_sortConn)
        self.Controls.Add(self._btt_SPLIT)
        self.Controls.Add(self._btt_CANCLE)
        self.Controls.Add(self._btt_pickPipe)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.Name = "MainForm"
        self.Text = "Split Pipe"
        self.TopMost = True
        self.Load += self.MainFormLoad
        self._grb_sortConn.ResumeLayout(False)
        self._grb_inputData.ResumeLayout(False)
        self._grb_inputData.PerformLayout()
        self._grb_MinMax.ResumeLayout(False)
        self.ResumeLayout(False)

    def Btt_pickPipeClick(self, sender, e):
        _pipe = pickPipe()
        self.selPipe = _pipe

    def Btt_SPLITClick(self, sender, e):
        '''__________'''
        pipe = self.selPipe
        new_dynPoints = []
        try:
            if pipe is not None:
                splitNumber1 = self._txb_K.Text
                if splitNumber1.strip():
                    try:
                        splitNumber = int(splitNumber1)
                    except ValueError:
                        splitNumber = None
                else:
                    splitNumber = None
                '''__________'''
                splitLength1 = self._txb_Length.Text
                if splitLength1.strip():
                    try:
                        splitLength = int(splitLength1)/304.8
                    except ValueError:
                        splitLength = None
                else:
                    splitLength = None
                '''___'''
                pipeCurve = pipe.Location.Curve
                conns = list(pipe.ConnectorManager.Connectors.GetEnumerator())
                originConns = list(c.Origin for c in conns)
                sortByCor_case1 = self._rbt_pX.Checked
                sortByCor_case2 = self._rbt_pY.Checked
                sortByCor_case3 = self._rbt_pZ.Checked
                '''___'''
                minCase = self._rbt_sortByMin.Checked
                maxCase = self._rbt_sortByMax.Checked
                '''___'''
                if sortByCor_case1 == True and minCase == True:
                    sortConns = sorted(originConns, key=lambda c: c.X)
                elif sortByCor_case1 == True and maxCase == True:
                    sortConns = sorted(originConns, key=lambda c: c.X)
                    sortConns.reverse()
                elif sortByCor_case2 == True and minCase == True:
                    sortConns = sorted(originConns, key=lambda c: c.Y)
                elif sortByCor_case2 == True and maxCase == True:
                    sortConns = sorted(originConns, key=lambda c: c.Y)
                    sortConns.reverse()
                elif sortByCor_case3 == True and minCase == True:
                    sortConns = sorted(originConns, key=lambda c: c.Z)
                elif sortByCor_case3 == True and maxCase == True:
                    sortConns = sorted(originConns, key=lambda c: c.Z)
                    sortConns.reverse()
                points = divideLineSegment(pipeCurve, splitLength, sortConns[0], sortConns[1])
                
                dynPoints = list(c.ToRevitType() for c in points)
                if splitNumber <= len(dynPoints):
                    new_dynPoints = dynPoints[1:splitNumber+1]
                else:
                    new_dynPoints = dynPoints[1:]
        except Exception as e:
            TransactionManager.Instance.ForceCloseTransaction()
            pass
        TransactionManager.Instance.EnsureInTransaction(doc)
        newPipes = splitPipeAtPoints(doc, pipe, new_dynPoints)
        TransactionManager.Instance.TransactionTaskDone()
        pass

    def Btt_CANCLEClick(self, sender, e):
        self.Close()
        pass

    def MainFormLoad(self, sender, e):
        pass

    def Txb_KTextChanged(self, sender, e):
        self.K = float(self._txb_K.Text)
        pass

    def Txb_LengthTextChanged(self, sender, e):
        self.length = float(self._txb_Length.Text)
        pass

    def Cbb_sortConnectorBySelectedIndexChanged(self, sender, e):
        pass

    def Rbt_pXCheckedChanged(self, sender, e):
        pass

    def Rbt_pYCheckedChanged(self, sender, e):
        pass

    def Rbt_pZCheckedChanged(self, sender, e):
        pass

    def Rbt_sortByMinCheckedChanged(self, sender, e):
        pass

    def Rbt_sortByMaxCheckedChanged(self, sender, e):
        pass

f = MainForm()
Application.Run(f)