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
from Autodesk.Revit.UI.Selection import ISelectionFilter
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

"""_____"""
class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._btt_PickPipe = System.Windows.Forms.Button()
		self._groupBox1 = System.Windows.Forms.GroupBox()
		self._cb_X = System.Windows.Forms.CheckBox()
		self._cb_Y = System.Windows.Forms.CheckBox()
		self._cb_Z = System.Windows.Forms.CheckBox()
		self._gb_Length = System.Windows.Forms.GroupBox()
		self._txb_splitLength = System.Windows.Forms.TextBox()
		self._btt_Split = System.Windows.Forms.Button()
		self._btt_Cancle = System.Windows.Forms.Button()
		self._groupBox1.SuspendLayout()
		self._gb_Length.SuspendLayout()
		self.SuspendLayout()
		# 
		# btt_PickPipe
		# 
		self._btt_PickPipe.AllowDrop = True
		self._btt_PickPipe.BackColor = System.Drawing.Color.FromArgb(192, 255, 192)
		self._btt_PickPipe.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
		self._btt_PickPipe.Cursor = System.Windows.Forms.Cursors.Arrow
		self._btt_PickPipe.Font = System.Drawing.Font("Meiryo UI", 10.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_PickPipe.Location = System.Drawing.Point(104, 21)
		self._btt_PickPipe.Name = "btt_PickPipe"
		self._btt_PickPipe.Size = System.Drawing.Size(164, 32)
		self._btt_PickPipe.TabIndex = 0
		self._btt_PickPipe.Text = "Select Pipe"
		self._btt_PickPipe.UseVisualStyleBackColor = False
		self._btt_PickPipe.Click += self.Btt_PickPipeClick
		# 
		# groupBox1
		# 
		self._groupBox1.Controls.Add(self._cb_Z)
		self._groupBox1.Controls.Add(self._cb_Y)
		self._groupBox1.Controls.Add(self._cb_X)
		self._groupBox1.Location = System.Drawing.Point(13, 70)
		self._groupBox1.Name = "groupBox1"
		self._groupBox1.Size = System.Drawing.Size(334, 81)
		self._groupBox1.TabIndex = 1
		self._groupBox1.TabStop = False
		self._groupBox1.Text = "Sort Pipe Connector"
		self._groupBox1.Enter += self.GroupBox1Enter
		# 
		# cb_X
		# 
		self._cb_X.Location = System.Drawing.Point(19, 38)
		self._cb_X.Name = "cb_X"
		self._cb_X.Size = System.Drawing.Size(69, 24)
		self._cb_X.TabIndex = 0
		self._cb_X.Text = "X"
		self._cb_X.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		self._cb_X.UseVisualStyleBackColor = True
		self._cb_X.UseWaitCursor = True
		self._cb_X.CheckedChanged += self.Cb_XCheckedChanged
		# 
		# cb_Y
		# 
		self._cb_Y.Location = System.Drawing.Point(140, 38)
		self._cb_Y.Name = "cb_Y"
		self._cb_Y.Size = System.Drawing.Size(69, 24)
		self._cb_Y.TabIndex = 1
		self._cb_Y.Text = "Y"
		self._cb_Y.UseVisualStyleBackColor = True
		self._cb_Y.CheckedChanged += self.Cb_YCheckedChanged
		# 
		# cb_Z
		# 
		self._cb_Z.Location = System.Drawing.Point(262, 38)
		self._cb_Z.Name = "cb_Z"
		self._cb_Z.Size = System.Drawing.Size(69, 24)
		self._cb_Z.TabIndex = 1
		self._cb_Z.Text = "Z"
		self._cb_Z.UseVisualStyleBackColor = True
		self._cb_Z.CheckedChanged += self.Cb_ZCheckedChanged
		# 
		# gb_Length
		# 
		self._gb_Length.Controls.Add(self._txb_splitLength)
		self._gb_Length.Location = System.Drawing.Point(13, 179)
		self._gb_Length.Name = "gb_Length"
		self._gb_Length.Size = System.Drawing.Size(334, 73)
		self._gb_Length.TabIndex = 2
		self._gb_Length.TabStop = False
		self._gb_Length.Text = "Split length"
		# 
		# txb_splitLength
		# 
		self._txb_splitLength.Location = System.Drawing.Point(37, 28)
		self._txb_splitLength.Name = "txb_splitLength"
		self._txb_splitLength.Size = System.Drawing.Size(275, 29)
		self._txb_splitLength.TabIndex = 0
		self._txb_splitLength.TextChanged += self.Txb_splitLengthTextChanged
		# 
		# btt_Split
		# 
		self._btt_Split.BackColor = System.Drawing.Color.FromArgb(255, 128, 128)
		self._btt_Split.Location = System.Drawing.Point(50, 273)
		self._btt_Split.Name = "btt_Split"
		self._btt_Split.Size = System.Drawing.Size(110, 38)
		self._btt_Split.TabIndex = 3
		self._btt_Split.Text = "SPLIT"
		self._btt_Split.UseVisualStyleBackColor = False
		self._btt_Split.Click += self.Btt_SplitClick
		# 
		# btt_Cancle
		# 
		self._btt_Cancle.BackColor = System.Drawing.Color.FromArgb(192, 255, 255)
		self._btt_Cancle.Location = System.Drawing.Point(204, 273)
		self._btt_Cancle.Name = "btt_Cancle"
		self._btt_Cancle.Size = System.Drawing.Size(110, 38)
		self._btt_Cancle.TabIndex = 3
		self._btt_Cancle.Text = "CANCLE"
		self._btt_Cancle.UseVisualStyleBackColor = False
		self._btt_Cancle.Click += self.Btt_CancleClick
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(362, 336)
		self.Controls.Add(self._btt_Cancle)
		self.Controls.Add(self._btt_Split)
		self.Controls.Add(self._gb_Length)
		self.Controls.Add(self._groupBox1)
		self.Controls.Add(self._btt_PickPipe)
		self.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.Text = "Split Pipe"
		self.TopMost = True
		self._groupBox1.ResumeLayout(False)
		self._gb_Length.ResumeLayout(False)
		self._gb_Length.PerformLayout()
		self.ResumeLayout(False)


	def Btt_PickPipeClick(self, sender, e):
		pass
	

	def GroupBox1Enter(self, sender, e):
		pass

	def Cb_XCheckedChanged(self, sender, e):
		pass

	def Cb_YCheckedChanged(self, sender, e):
		pass

	def Cb_ZCheckedChanged(self, sender, e):
		pass

	def Txb_splitLengthTextChanged(self, sender, e):
		pass

	def Btt_SplitClick(self, sender, e):
		pass

	def Btt_CancleClick(self, sender, e):
		pass
	
f = MainForm()
Application.Run(f)