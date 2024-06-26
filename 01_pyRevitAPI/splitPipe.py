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
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion
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
	ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter,"pick Pipe")
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
#region ___divideLineSegment
def divideLineSegment(doc, pipe, pointA, LengthA):
    TransactionManager.Instance.EnsureInTransaction(doc)
    lineSegment = pipe.Location.Curve
    start_point = lineSegment.GetEndPoint(0)
    end_point = lineSegment.GetEndPoint(1)
    vector = end_point - start_point
    total_length = vector.GetLength()*304.8
    direction = vector.Normalize()
    num_segments = int(total_length / LengthA)
    points = []
    desPoints =[]
    current_point = pointA
    for i in range(num_segments + 1):
        points.append(current_point)
        desPoints.append(current_point.ToPoint())
        # current_point = current_point + direction.Multiply(LengthA)
    # new_lines = []
    # for i in range(len(points) - 1):
    #     new_line = Line.CreateBound(points[i], points[i + 1])
    #     new_lines.append(new_line.ToProtoType())
    TransactionManager.Instance.TransactionTaskDone()
    return desPoints
#endregion
class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._btt_pickPipe = System.Windows.Forms.Button()
		self._grb_sortConn = System.Windows.Forms.GroupBox()
		self._cb_X = System.Windows.Forms.CheckBox()
		self._cb_Y = System.Windows.Forms.CheckBox()
		self._cb_Z = System.Windows.Forms.CheckBox()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._btt_SPLIT = System.Windows.Forms.Button()
		self._lb_FVC = System.Windows.Forms.Label()
		self._txb_Length = System.Windows.Forms.TextBox()
		self._lb_Length = System.Windows.Forms.Label()
		self._grb_sortConn.SuspendLayout()
		self.SuspendLayout()
		# 
		# btt_pickPipe
		# 
		self._btt_pickPipe.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_pickPipe.ForeColor = System.Drawing.Color.Red
		self._btt_pickPipe.Location = System.Drawing.Point(11, 32)
		self._btt_pickPipe.Name = "btt_pickPipe"
		self._btt_pickPipe.Size = System.Drawing.Size(133, 55)
		self._btt_pickPipe.TabIndex = 0
		self._btt_pickPipe.Text = "PICK PIPE"
		self._btt_pickPipe.UseVisualStyleBackColor = True
		self._btt_pickPipe.Click += self.Btt_pickPipeClick
		# 
		# grb_sortConn
		# 
		self._grb_sortConn.Controls.Add(self._cb_Z)
		self._grb_sortConn.Controls.Add(self._cb_Y)
		self._grb_sortConn.Controls.Add(self._cb_X)
		self._grb_sortConn.Cursor = System.Windows.Forms.Cursors.Default
		self._grb_sortConn.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._grb_sortConn.Location = System.Drawing.Point(152, 27)
		self._grb_sortConn.Name = "grb_sortConn"
		self._grb_sortConn.Size = System.Drawing.Size(172, 124)
		self._grb_sortConn.TabIndex = 1
		self._grb_sortConn.TabStop = False
		self._grb_sortConn.Text = "Sort Connector by"
		self._grb_sortConn.UseCompatibleTextRendering = True
		# 
		# cb_X
		# 
		self._cb_X.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._cb_X.Location = System.Drawing.Point(60, 22)
		self._cb_X.Name = "cb_X"
		self._cb_X.Size = System.Drawing.Size(52, 24)
		self._cb_X.TabIndex = 0
		self._cb_X.Text = "X"
		self._cb_X.UseVisualStyleBackColor = True
		self._cb_X.CheckedChanged += self.Cb_XCheckedChanged
		# 
		# cb_Y
		# 
		self._cb_Y.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._cb_Y.Location = System.Drawing.Point(60, 59)
		self._cb_Y.Name = "cb_Y"
		self._cb_Y.Size = System.Drawing.Size(52, 24)
		self._cb_Y.TabIndex = 0
		self._cb_Y.Text = "Y"
		self._cb_Y.UseVisualStyleBackColor = True
		self._cb_Y.CheckedChanged += self.Cb_YCheckedChanged
		# 
		# cb_Z
		# 
		self._cb_Z.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._cb_Z.Location = System.Drawing.Point(60, 94)
		self._cb_Z.Name = "cb_Z"
		self._cb_Z.Size = System.Drawing.Size(52, 24)
		self._cb_Z.TabIndex = 0
		self._cb_Z.Text = "Z"
		self._cb_Z.UseVisualStyleBackColor = True
		self._cb_Z.CheckedChanged += self.Cb_ZCheckedChanged
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
		self._btt_CANCLE.Location = System.Drawing.Point(212, 185)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(101, 37)
		self._btt_CANCLE.TabIndex = 0
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# btt_SPLIT
		# 
		self._btt_SPLIT.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_SPLIT.ForeColor = System.Drawing.Color.Red
		self._btt_SPLIT.Location = System.Drawing.Point(105, 185)
		self._btt_SPLIT.Name = "btt_SPLIT"
		self._btt_SPLIT.Size = System.Drawing.Size(101, 37)
		self._btt_SPLIT.TabIndex = 0
		self._btt_SPLIT.Text = "SPLIT"
		self._btt_SPLIT.UseVisualStyleBackColor = True
		self._btt_SPLIT.Click += self.Btt_SPLITClick
		# 
		# lb_FVC
		# 
		self._lb_FVC.Font = System.Drawing.Font("Meiryo UI", 4.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_FVC.Location = System.Drawing.Point(11, 209)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(56, 20)
		self._lb_FVC.TabIndex = 2
		self._lb_FVC.Text = "@FVC"
		self._lb_FVC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# txb_Length
		# 
		self._txb_Length.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_Length.Location = System.Drawing.Point(11, 121)
		self._txb_Length.Name = "txb_Length"
		self._txb_Length.Size = System.Drawing.Size(133, 23)
		self._txb_Length.TabIndex = 3
		self._txb_Length.TextChanged += self.Txb_LengthTextChanged
		# 
		# lb_Length
		# 
		self._lb_Length.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_Length.Location = System.Drawing.Point(11, 98)
		self._lb_Length.Name = "lb_Length"
		self._lb_Length.Size = System.Drawing.Size(72, 20)
		self._lb_Length.TabIndex = 2
		self._lb_Length.Text = "Length:"
		self._lb_Length.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(325, 234)
		self.Controls.Add(self._txb_Length)
		self.Controls.Add(self._lb_Length)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._grb_sortConn)
		self.Controls.Add(self._btt_SPLIT)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_pickPipe)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.Text = "Split Pipe"
		self.TopMost = True
		# self.Load += self.MainFormLoad
		self._grb_sortConn.ResumeLayout(False)
		self.ResumeLayout(False)
		self.PerformLayout()

	def Btt_pickPipeClick(self, sender, e):
		pipe = pickPipe()
		conns = list(pipe.ConnectorManager.Connectors.GetEnumerator())
		lst = list(c.Origin for c in conns)
		self.connsOrigin = list(c.ToPoint() for c in lst)
		self.selPipe = pipe
		pass

	def Cb_XCheckedChanged(self, sender, e):
		pass

	def Cb_YCheckedChanged(self, sender, e):
		pass

	def Cb_ZCheckedChanged(self, sender, e):
		pass

	def Txb_LengthTextChanged(self, sender, e):
		self.length = self._txb_Length
		pass
	def Btt_SPLITClick(self, sender, e):
		pipe = self.selPipe
		if self._cb_X.Checked == True:
			sortedConns = sorted(self.connsOrigin, key = lambda c: c.X)
		elif self._cb_Y.Checked == True:
			sortedConns = sorted(self.connsOrigin, key = lambda c: c.Y)
		elif self._cb_Z.Checked == True:
			sortedConns = sorted(self.connsOrigin, key = lambda c: c.Z)
		sortedConn = sortedConns[0]
		splitLength = float(self.length)
		splitPoints = divideLineSegment(doc, self.selPipe, sortedConn, splitLength)

		pass

	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass
	
f = MainForm()
Application.Run(f)