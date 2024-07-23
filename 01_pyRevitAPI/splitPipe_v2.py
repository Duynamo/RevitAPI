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
from Autodesk.Revit.DB.Plumbing import*
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
#region ___def devideLineSegment
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
	TransactionManager.Instance.EnsureInTransaction(doc)
	for point in points:
		#currentPipe = pipe
		pipeLocation = currentPipe.Location
		if isinstance(pipeLocation, LocationCurve):
			pipeCurve = pipeLocation.Curve
			if pipeCurve is not None:
				newPipeIds = PlumbingUtils.BreakCurve(doc, currentPipe.Id, point)             
				newPipe = doc.GetElement(newPipeIds)
				newPipes.append(newPipe)
				currentPipe = newPipe
	TransactionManager.Instance.TransactionTaskDone
	return newPipes
		
#endregion

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
		self._groupBox1 = System.Windows.Forms.GroupBox()
		self._grb_inputData = System.Windows.Forms.GroupBox()
		self._rbtt_endPoint0 = System.Windows.Forms.RadioButton()
		self._rbtt_endPoint1 = System.Windows.Forms.RadioButton()
		self._grb_sortConn.SuspendLayout()
		self._grb_inputData.SuspendLayout()
		self.SuspendLayout()
		# 
		# btt_pickPipe
		# 
		self._btt_pickPipe.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_pickPipe.ForeColor = System.Drawing.Color.Red
		self._btt_pickPipe.Location = System.Drawing.Point(18, 37)
		self._btt_pickPipe.Name = "btt_pickPipe"
		self._btt_pickPipe.Size = System.Drawing.Size(133, 41)
		self._btt_pickPipe.TabIndex = 0
		self._btt_pickPipe.Text = "PICK PIPE"
		self._btt_pickPipe.UseVisualStyleBackColor = True
		self._btt_pickPipe.Click += self.Btt_pickPipeClick
		# 
		# grb_sortConn
		# 
		self._grb_sortConn.Controls.Add(self._rbtt_endPoint1)
		self._grb_sortConn.Controls.Add(self._rbtt_endPoint0)
		self._grb_sortConn.Controls.Add(self._groupBox1)
		self._grb_sortConn.Cursor = System.Windows.Forms.Cursors.Default
		self._grb_sortConn.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._grb_sortConn.Location = System.Drawing.Point(11, 106)
		self._grb_sortConn.Name = "grb_sortConn"
		self._grb_sortConn.Size = System.Drawing.Size(158, 91)
		self._grb_sortConn.TabIndex = 1
		self._grb_sortConn.TabStop = False
		self._grb_sortConn.Text = "Sort Connector by"
		self._grb_sortConn.UseCompatibleTextRendering = True
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
		self._btt_CANCLE.Location = System.Drawing.Point(259, 214)
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
		self._btt_SPLIT.Location = System.Drawing.Point(152, 214)
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
		self._lb_FVC.ForeColor = System.Drawing.Color.Blue
		self._lb_FVC.Location = System.Drawing.Point(12, 234)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(56, 20)
		self._lb_FVC.TabIndex = 2
		self._lb_FVC.Text = "@FVC"
		self._lb_FVC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# txb_Length
		# 
		self._txb_Length.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_Length.Location = System.Drawing.Point(30, 60)
		self._txb_Length.Name = "txb_Length"
		self._txb_Length.Size = System.Drawing.Size(133, 23)
		self._txb_Length.TabIndex = 3
		self._txb_Length.TextChanged += self.Txb_LengthTextChanged
		self._txb_Length.Text = "1000"
		# 
		# lb_Length
		# 
		self._lb_Length.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_Length.Location = System.Drawing.Point(25, 33)
		self._lb_Length.Name = "lb_Length"
		self._lb_Length.Size = System.Drawing.Size(117, 20)
		self._lb_Length.TabIndex = 2
		self._lb_Length.Text = "Length(mm):"
		self._lb_Length.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# txb_K
		# 
		self._txb_K.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_K.Location = System.Drawing.Point(30, 130)
		self._txb_K.Name = "txb_K"
		self._txb_K.Size = System.Drawing.Size(133, 23)
		self._txb_K.TabIndex = 5
		self._txb_K.TextChanged += self.Txb_KTextChanged
		self._txb_K.Text = '1000'
		# 
		# lb_splitNumber
		# 
		self._lb_splitNumber.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_splitNumber.Location = System.Drawing.Point(22, 104)
		self._lb_splitNumber.Name = "lb_splitNumber"
		self._lb_splitNumber.Size = System.Drawing.Size(36, 20)
		self._lb_splitNumber.TabIndex = 4
		self._lb_splitNumber.Text = "K:"
		self._lb_splitNumber.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# groupBox1
		# 
		self._groupBox1.Location = System.Drawing.Point(155, 61)
		self._groupBox1.Name = "groupBox1"
		self._groupBox1.Size = System.Drawing.Size(200, 100)
		self._groupBox1.TabIndex = 6
		self._groupBox1.TabStop = False
		self._groupBox1.Text = "groupBox1"
		# 
		# grb_inputData
		# 
		self._grb_inputData.Controls.Add(self._lb_Length)
		self._grb_inputData.Controls.Add(self._txb_K)
		self._grb_inputData.Controls.Add(self._txb_Length)
		self._grb_inputData.Controls.Add(self._lb_splitNumber)
		self._grb_inputData.Cursor = System.Windows.Forms.Cursors.Default
		self._grb_inputData.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._grb_inputData.Location = System.Drawing.Point(175, 32)
		self._grb_inputData.Name = "grb_inputData"
		self._grb_inputData.RightToLeft = System.Windows.Forms.RightToLeft.No
		self._grb_inputData.Size = System.Drawing.Size(185, 165)
		self._grb_inputData.TabIndex = 6
		self._grb_inputData.TabStop = False
		self._grb_inputData.Text = "input"
		# 
		# rbtt_endPoint0
		# 
		self._rbtt_endPoint0.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._rbtt_endPoint0.ForeColor = System.Drawing.Color.Red
		self._rbtt_endPoint0.Location = System.Drawing.Point(7, 26)
		self._rbtt_endPoint0.Name = "rbtt_endPoint0"
		self._rbtt_endPoint0.Size = System.Drawing.Size(119, 24)
		self._rbtt_endPoint0.TabIndex = 7
		self._rbtt_endPoint0.TabStop = True
		self._rbtt_endPoint0.Text = "endPoint0"
		self._rbtt_endPoint0.UseVisualStyleBackColor = True
		self._rbtt_endPoint0.CheckedChanged += self.Rbtt_endPoint0CheckedChanged
		# 
		# rbtt_endPoint1
		# 
		self._rbtt_endPoint1.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._rbtt_endPoint1.ForeColor = System.Drawing.Color.Red
		self._rbtt_endPoint1.Location = System.Drawing.Point(7, 56)
		self._rbtt_endPoint1.Name = "rbtt_endPoint1"
		self._rbtt_endPoint1.Size = System.Drawing.Size(119, 24)
		self._rbtt_endPoint1.TabIndex = 8
		self._rbtt_endPoint1.TabStop = True
		self._rbtt_endPoint1.Text = "endPoint1"
		self._rbtt_endPoint1.UseVisualStyleBackColor = True
		self._rbtt_endPoint1.CheckedChanged += self.Rbtt_endPoint1CheckedChanged
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(369, 263)
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
		self.ResumeLayout(False)
	def Txb_LengthTextChanged(self, sender, e):
		#self.length = float(self._txb_Length.Text)
		pass
	def Btt_pickPipeClick(self, sender, e):
		_pipe = pickPipe()
		self.selPipe = _pipe
		pass
	def Btt_SPLITClick(self, sender, e):
		'''__________'''
		splitNumber1 = self._txb_K.Text
		if splitNumber1.strip():
			try:
				splitNumber = int(splitNumber1)
			except vars.ValueError:
				splitNumber = None
		else:
			splitNumber = None
		'''__________'''
		splitLength1 = self._txb_Length.Text
		if splitLength1.strip():
			try:
				splitLength = int(splitLength1)/304.8
			except vars.ValueError:
				splitLength = None
		else:
			splitLength = None
		'''__________'''
		pipe = self.selPipe
		# inKey = self._cbb_sortConnectorBy.SelectedItem
		check1 = self._rbtt_endPoint0.Checked
		check2 = self._rbtt_endPoint1.Checked
		sortConns = []
		TransactionManager.Instance.EnsureInTransaction(doc)
		try:
			if splitNumber > 0:
				if pipe is not None:
					pipeCurve  = pipe.Location.Curve
					conns = list(pipe.ConnectorManager.Connectors.GetEnumerator())
					originConns = [c.Origin for c in conns]
					if check1 == True:
						sortConns.append(originConns[0])
						sortConns.append(originConns[1])
					if check2 == True:
						sortConns.append(originConns[1])
						sortConns.append(originConns[0])
					else :
						sortConns = originConns
					points = divideLineSegment(pipeCurve, splitLength, sortConns[0], sortConns[1])
					dynPoints = [c.ToRevitType() for c in points]
					new_dynPoints = []
					if splitNumber <= len(dynPoints):
						new_dynPoints = dynPoints[1:splitNumber+1]
					else:
						new_dynPoints = dynPoints[1:]

		except Exception as e:
			TransactionManager.Instance.ForceCloseTransaction()
			pass
		

        
		TransactionManager.Instance.EnsureInTransaction(doc)
		newPipes = splitPipeAtPoints(doc, pipe, new_dynPoints)
		TransactionManager.Instance.TransactionTaskDone
		pass
		TransactionManager.Instance.TransactionTaskDone
	def Txb_KTextChanged(self, sender, e):
		self.K = float(self._txb_K.Text)
		pass
	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass
	def Cbb_sortConnectorBySelectedIndexChanged(self, sender, e):
		pass
	def MainFormLoad(self, sender, e):
		pass	
	def Rbtt_endPoint0CheckedChanged(self, sender, e):
		pass

	def Rbtt_endPoint1CheckedChanged(self, sender, e):
		pass	
f = MainForm()
Application.Run(f)