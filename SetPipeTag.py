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
		self._btt_SetPipeTag = System.Windows.Forms.Button()
		self._lb_FVC = System.Windows.Forms.Label()
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
		# btt_SetPipeTag
		# 
		self._btt_SetPipeTag.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_SetPipeTag.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_SetPipeTag.FlatAppearance.BorderSize = 2
		self._btt_SetPipeTag.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue
		self._btt_SetPipeTag.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(255, 255, 128)
		self._btt_SetPipeTag.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_SetPipeTag.ForeColor = System.Drawing.Color.Red
		self._btt_SetPipeTag.Location = System.Drawing.Point(224, 124)
		self._btt_SetPipeTag.Name = "btt_SetPipeTag"
		self._btt_SetPipeTag.Size = System.Drawing.Size(127, 28)
		self._btt_SetPipeTag.TabIndex = 4
		self._btt_SetPipeTag.Text = "Set Pipe Tag"
		self._btt_SetPipeTag.UseVisualStyleBackColor = True
		self._btt_SetPipeTag.Click += self.Btt_PipeTagClick
		# 
		# lb_FVC
		# 
		self._lb_FVC.Font = System.Drawing.Font("Meiryo UI", 5, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_FVC.Location = System.Drawing.Point(13, 176)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(38, 17)
		self._lb_FVC.TabIndex = 7
		self._lb_FVC.Text = "@FVC"
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(365, 204)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._btt_SetPipeTag)
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
		# msg = "Pipes were tagged."
		# TaskDialog.Show("^------^", msg)		
		self.Close()	
		pass


f = MainForm()
Application.Run(f)
