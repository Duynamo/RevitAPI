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

#Filter all pipe tags in Doc
pipe_tags = []
tagCollector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeTags).WhereElementIsElementType().ToElements()
pipe_tags.append(i for i in tagCollector)

#Filter all DimStyle in Doc
dim_Styles = []
dimStyleCollector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Dimensions).ToElements()
dim_Styles.append(i for i in dimStyleCollector)

#Filter all pipe in ActiveView
pipesList = []
pipesCollector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
pipesList.append(i for i in pipesCollector)

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._cbb_PipeTag = System.Windows.Forms.ComboBox()
		self._lb_PipeTag = System.Windows.Forms.Label()
		self._lb_CenterDim = System.Windows.Forms.Label()
		self._btt_PipeTag = System.Windows.Forms.Button()
		self._btt_CenterDim = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._lb_FVC = System.Windows.Forms.Label()
		self._cbb_DimStyte = System.Windows.Forms.ComboBox()
		self.SuspendLayout()
		# 
		# cbb_PipeTag
		# 
		self._cbb_PipeTag.Cursor = System.Windows.Forms.Cursors.Default
		self._cbb_PipeTag.DisplayMember = "Name"
		self._cbb_PipeTag.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cbb_PipeTag.ForeColor = System.Drawing.Color.Red
		self._cbb_PipeTag.FormattingEnabled = True
		self._cbb_PipeTag.Location = System.Drawing.Point(122, 37)
		self._cbb_PipeTag.Name = "cbb_PipeTag"
		self._cbb_PipeTag.Size = System.Drawing.Size(229, 27)
		self._cbb_PipeTag.TabIndex = 0
		self._cbb_PipeTag.SelectedIndexChanged += self.Cbb_PipeTagSelectedIndexChanged
		self._cbb_PipeTag.Items.AddRange(System.Array[System.Object](tagCollector)) 
		self._cbb_PipeTag.SelectedIndex = 0				
		# 
		# lb_PipeTag
		# 
		self._lb_PipeTag.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
		self._lb_PipeTag.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_PipeTag.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._lb_PipeTag.ForeColor = System.Drawing.Color.Red
		self._lb_PipeTag.Location = System.Drawing.Point(13, 37)
		self._lb_PipeTag.Name = "lb_PipeTag"
		self._lb_PipeTag.Size = System.Drawing.Size(102, 27)
		self._lb_PipeTag.TabIndex = 1
		self._lb_PipeTag.Text = "Pipe Tag"	
		# 
		# lb_CenterDim
		# 
		self._lb_CenterDim.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
		self._lb_CenterDim.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_CenterDim.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._lb_CenterDim.ForeColor = System.Drawing.Color.Red
		self._lb_CenterDim.Location = System.Drawing.Point(12, 125)
		self._lb_CenterDim.Name = "lb_CenterDim"
		self._lb_CenterDim.Size = System.Drawing.Size(103, 27)
		self._lb_CenterDim.TabIndex = 2
		self._lb_CenterDim.Text = "Center Dim"
		# 
		# btt_PipeTag
		# 
		self._btt_PipeTag.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_PipeTag.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_PipeTag.FlatAppearance.BorderSize = 2
		self._btt_PipeTag.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue
		self._btt_PipeTag.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(255, 255, 128)
		self._btt_PipeTag.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_PipeTag.ForeColor = System.Drawing.Color.Red
		self._btt_PipeTag.Location = System.Drawing.Point(56, 210)
		self._btt_PipeTag.Name = "btt_PipeTag"
		self._btt_PipeTag.Size = System.Drawing.Size(85, 28)
		self._btt_PipeTag.TabIndex = 4
		self._btt_PipeTag.Text = "Pipe Tag"
		self._btt_PipeTag.UseVisualStyleBackColor = True
		self._btt_PipeTag.Click += self.Btt_PipeTagClick
		# 
		# btt_CenterDim
		# 
		self._btt_CenterDim.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_CenterDim.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_CenterDim.FlatAppearance.BorderSize = 2
		self._btt_CenterDim.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue
		self._btt_CenterDim.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(255, 255, 128)
		self._btt_CenterDim.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CenterDim.ForeColor = System.Drawing.Color.Red
		self._btt_CenterDim.Location = System.Drawing.Point(148, 210)
		self._btt_CenterDim.Name = "btt_CenterDim"
		self._btt_CenterDim.Size = System.Drawing.Size(108, 28)
		self._btt_CenterDim.TabIndex = 5
		self._btt_CenterDim.Text = "CenterDim"
		self._btt_CenterDim.UseVisualStyleBackColor = True
		self._btt_CenterDim.Click += self.Btt_TrueLengthClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_CANCLE.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_CANCLE.FlatAppearance.BorderSize = 2
		self._btt_CANCLE.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue
		self._btt_CANCLE.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(255, 255, 128)
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
		self._btt_CANCLE.Location = System.Drawing.Point(263, 210)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(88, 28)
		self._btt_CANCLE.TabIndex = 6
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# lb_FVC
		# 
		self._lb_FVC.Font = System.Drawing.Font("Meiryo UI", 5, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_FVC.Location = System.Drawing.Point(10, 235)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(38, 17)
		self._lb_FVC.TabIndex = 7
		self._lb_FVC.Text = "@FVC"
		# 
		# cbb_DimStyte
		# 
		self._cbb_DimStyte.Cursor = System.Windows.Forms.Cursors.Default
		self._cbb_DimStyte.DisplayMember = "Name"
		self._cbb_DimStyte.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cbb_DimStyte.ForeColor = System.Drawing.Color.Red
		self._cbb_DimStyte.FormattingEnabled = True
		self._cbb_DimStyte.Location = System.Drawing.Point(124, 123)
		self._cbb_DimStyte.Name = "cbb_DimStyte"
		self._cbb_DimStyte.Size = System.Drawing.Size(229, 27)
		self._cbb_DimStyte.TabIndex = 8
		self._cbb_DimStyte.SelectedIndexChanged += self.Cbb_DimStyteSelectedIndexChanged
		self._cbb_DimStyte.Items.AddRange(System.Array[System.Object](dimStyleCollector)) 
		# self._cbb_DimStyte.SelectedIndex = 0				
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(365, 258)
		self.Controls.Add(self._cbb_DimStyte)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_CenterDim)
		self.Controls.Add(self._btt_PipeTag)
		self.Controls.Add(self._lb_CenterDim)
		self.Controls.Add(self._lb_PipeTag)
		self.Controls.Add(self._cbb_PipeTag)
		self.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.Text = "Set Parameter"
		self.TopMost = True
		self.ResumeLayout(False)

	def Cbb_PipeTagSelectedIndexChanged(self, sender, e):
		pass

	def Btt_PipeTagClick(self, sender, e):
		desTag = self._cbb_PipeTag.SelectedItem
		pipeTags = []
		TransactionManager.Instance.EnsureInTransaction(doc)
		for pipe in pipesCollector:
			pipeRef = Reference(pipe)
			tagMode = TagMode.TM_ADDBY_CATEGORY
			tag = IndependentTag.Create(doc, view.Id, pipeRef, True , tagMode, TagOrientation.AnyModelDirection, XYZ.BasisZ)
			pipeLocation = pipe.Location.Curve.Evaluate(0.5, True)
			tagLocation = XYZ(pipeLocation.X, pipeLocation.Y, pipeLocation.Z)
			tag.TagHeadPosition = tagLocation 
			pipeTags.append(tag)		
		TransactionManager.Instance.TransactionTaskDone()
		msg = "Pipes were tagged."
		TaskDialog.Show("^------^", msg)		
		self.Close()	
		pass

	def Btt_TrueLengthClick(self, sender, e):
		TransactionManager.Instance.EnsureInTransaction(doc)
		centerPointsList = []
		desDimStyle = self._cbb_DimStyte.SelectedItem
		for pipe in pipesCollector:
			pipeLocationCurve = pipe.Location.Curve
			centerPoint = pipeLocationCurve.Evaluate(0.5, True)
			centerPointsList.append(centerPoint)

		for i in range(len(centerPointsList) - 1):
			for j in range(i + 1, len(centerPointsList)):
				reference1 = Reference(centerPointsList[i])
				reference2 = Reference(centerPointsList[j])
				line = Line.CreateBound(centerPointsList[i], centerPointsList[j])
				dimension = doc.Create.NewDimension(doc.ActiveView,line, reference1, reference2, desDimStyle)

				# Set the location of the dimension text
				# dimension.Leader.End = centerPointsList[i]
				# dimension.Leader.End = centerPointsList[j]

		TransactionManager.Instance.TransactionTaskDone()		
		pass

	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass

	def Cbb_DimStyteSelectedIndexChanged(self, sender, e):
		pass
f = MainForm()
Application.Run(f)
