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

"""_____________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = uiapp.ActiveUIDocument
view = doc.ActiveView
"""_____________________________"""
def pickObjects():
	elements = []
	planarFace = []
	refs = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "pick face")
	for i in refs:
		ele = doc.GetElement(i.ElementId)
		face = ele.GetGeometryObjectFromReference(i)
		DBFace = face.GetSurface()
		elements.append(ele)
		planarFace.append(DBFace)
	return  refs, elements, planarFace
""""""
"""_____________________________"""
class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._grb_Face = System.Windows.Forms.GroupBox()
		self._btt_pickFace = System.Windows.Forms.Button()
		self._txb_disFromPickedFace = System.Windows.Forms.TextBox()
		self._lb_disFromFaceToCBL = System.Windows.Forms.Label()
		self._grb_Pipes = System.Windows.Forms.GroupBox()
		self._btt_pickPipes = System.Windows.Forms.Button()
		self._lb_ = System.Windows.Forms.Label()
		self._clb_selectedPipes = System.Windows.Forms.CheckedListBox()
		self._lb_pickedPipes = System.Windows.Forms.Label()
		self._lb_FVC = System.Windows.Forms.Label()
		self._btt_set = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._btt_removePipe = System.Windows.Forms.Button()
		self._ckb_AllNone = System.Windows.Forms.CheckBox()
		self._grb_Face.SuspendLayout()
		self._grb_Pipes.SuspendLayout()
		self.SuspendLayout()
		# 
		# grb_Face
		# 
		self._grb_Face.Controls.Add(self._lb_disFromFaceToCBL)
		self._grb_Face.Controls.Add(self._txb_disFromPickedFace)
		self._grb_Face.Controls.Add(self._btt_pickFace)
		self._grb_Face.Location = System.Drawing.Point(13, 33)
		self._grb_Face.Name = "grb_Face"
		self._grb_Face.Size = System.Drawing.Size(373, 113)
		self._grb_Face.TabIndex = 0
		self._grb_Face.TabStop = False
		self._grb_Face.Text = "getFace"
		# 
		# btt_pickFace
		# 
		self._btt_pickFace.BackColor = System.Drawing.Color.FromArgb(255, 255, 128)
		self._btt_pickFace.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_pickFace.ForeColor = System.Drawing.Color.Red
		self._btt_pickFace.Location = System.Drawing.Point(6, 42)
		self._btt_pickFace.Name = "btt_pickFace"
		self._btt_pickFace.Size = System.Drawing.Size(110, 50)
		self._btt_pickFace.TabIndex = 0
		self._btt_pickFace.Text = "Pick Face"
		self._btt_pickFace.UseVisualStyleBackColor = False
		self._btt_pickFace.Click += self.Btt_pickFaceClick
		# 
		# txb_disFromPickedFace
		# 
		self._txb_disFromPickedFace.Font = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_disFromPickedFace.ForeColor = System.Drawing.Color.Blue
		self._txb_disFromPickedFace.Location = System.Drawing.Point(149, 70)
		self._txb_disFromPickedFace.Multiline = True
		self._txb_disFromPickedFace.Name = "txb_disFromPickedFace"
		self._txb_disFromPickedFace.Size = System.Drawing.Size(208, 22)
		self._txb_disFromPickedFace.TabIndex = 1
		self._txb_disFromPickedFace.TextChanged += self.TextBox1TextChanged
		# 
		# lb_disFromFaceToCBL
		# 
		self._lb_disFromFaceToCBL.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_disFromFaceToCBL.Location = System.Drawing.Point(149, 18)
		self._lb_disFromFaceToCBL.Name = "lb_disFromFaceToCBL"
		self._lb_disFromFaceToCBL.Size = System.Drawing.Size(187, 49)
		self._lb_disFromFaceToCBL.TabIndex = 2
		self._lb_disFromFaceToCBL.Text = "Dis From Picked Face to Closest Level below"
		# 
		# grb_Pipes
		# 
		self._grb_Pipes.Controls.Add(self._ckb_AllNone)
		self._grb_Pipes.Controls.Add(self._btt_removePipe)
		self._grb_Pipes.Controls.Add(self._lb_pickedPipes)
		self._grb_Pipes.Controls.Add(self._clb_selectedPipes)
		self._grb_Pipes.Controls.Add(self._btt_pickPipes)
		self._grb_Pipes.Location = System.Drawing.Point(13, 160)
		self._grb_Pipes.Name = "grb_Pipes"
		self._grb_Pipes.Size = System.Drawing.Size(373, 196)
		self._grb_Pipes.TabIndex = 1
		self._grb_Pipes.TabStop = False
		self._grb_Pipes.Text = "getPipe"
		# 
		# btt_pickPipes
		# 
		self._btt_pickPipes.BackColor = System.Drawing.Color.FromArgb(255, 255, 128)
		self._btt_pickPipes.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_pickPipes.ForeColor = System.Drawing.Color.Red
		self._btt_pickPipes.Location = System.Drawing.Point(6, 69)
		self._btt_pickPipes.Name = "btt_pickPipes"
		self._btt_pickPipes.Size = System.Drawing.Size(110, 50)
		self._btt_pickPipes.TabIndex = 3
		self._btt_pickPipes.Text = "Pick Pipes"
		self._btt_pickPipes.UseVisualStyleBackColor = False
		self._btt_pickPipes.Click += self.Btt_pickPipesClick
		# 
		# clb_selectedPipes
		# 
		self._clb_selectedPipes.AllowDrop = True
		self._clb_selectedPipes.CheckOnClick = True
		self._clb_selectedPipes.FormattingEnabled = True
		self._clb_selectedPipes.Location = System.Drawing.Point(142, 69)
		self._clb_selectedPipes.Name = "clb_selectedPipes"
		self._clb_selectedPipes.Size = System.Drawing.Size(225, 106)
		self._clb_selectedPipes.TabIndex = 4
		self._clb_selectedPipes.SelectedIndexChanged += self.Clb_selectedPipesSelectedIndexChanged
		# 
		# lb_pickedPipes
		# 
		self._lb_pickedPipes.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_pickedPipes.Location = System.Drawing.Point(142, 18)
		self._lb_pickedPipes.Name = "lb_pickedPipes"
		self._lb_pickedPipes.Size = System.Drawing.Size(131, 23)
		self._lb_pickedPipes.TabIndex = 3
		self._lb_pickedPipes.Text = "Selected Pipes"
		# 
		# lb_FVC
		# 
		self._lb_FVC.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_FVC.Location = System.Drawing.Point(12, 433)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(50, 23)
		self._lb_FVC.TabIndex = 3
		self._lb_FVC.Text = "@FVC"
		# 
		# btt_set
		# 
		self._btt_set.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_set.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_set.ForeColor = System.Drawing.Color.Red
		self._btt_set.Location = System.Drawing.Point(162, 393)
		self._btt_set.Name = "btt_set"
		self._btt_set.Size = System.Drawing.Size(102, 43)
		self._btt_set.TabIndex = 4
		self._btt_set.Text = "SET"
		self._btt_set.UseVisualStyleBackColor = True
		self._btt_set.Click += self.Btt_setClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
		self._btt_CANCLE.Location = System.Drawing.Point(284, 393)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(102, 43)
		self._btt_CANCLE.TabIndex = 5
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# btt_removePipe
		# 
		self._btt_removePipe.BackColor = System.Drawing.Color.FromArgb(255, 255, 128)
		self._btt_removePipe.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_removePipe.ForeColor = System.Drawing.Color.Red
		self._btt_removePipe.Location = System.Drawing.Point(6, 125)
		self._btt_removePipe.Name = "btt_removePipe"
		self._btt_removePipe.Size = System.Drawing.Size(110, 50)
		self._btt_removePipe.TabIndex = 5
		self._btt_removePipe.Text = "Remove"
		self._btt_removePipe.UseVisualStyleBackColor = False
		self._btt_removePipe.Click += self.Btt_removePipeClick
		# 
		# ckb_AllNone
		# 
		self._ckb_AllNone.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._ckb_AllNone.Location = System.Drawing.Point(279, 39)
		self._ckb_AllNone.Name = "ckb_AllNone"
		self._ckb_AllNone.Size = System.Drawing.Size(88, 24)
		self._ckb_AllNone.TabIndex = 6
		self._ckb_AllNone.Text = "All/None"
		self._ckb_AllNone.UseVisualStyleBackColor = True
		self._ckb_AllNone.CheckedChanged += self.Ckb_AllNoneCheckedChanged
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(398, 459)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_set)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._lb_)
		self.Controls.Add(self._grb_Pipes)
		self.Controls.Add(self._grb_Face)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.Text = "setPipeDistance"
		self.TopMost = True
		self._grb_Face.ResumeLayout(False)
		self._grb_Face.PerformLayout()
		self._grb_Pipes.ResumeLayout(False)
		self.ResumeLayout(False)



	def Btt_pickFaceClick(self, sender, e):
		desFace = pickFace()
		pass

	def TextBox1TextChanged(self, sender, e):
		pass

	def Btt_pickPipesClick(self, sender, e):
		pass

	def Btt_removePipeClick(self, sender, e):
		pass

	def Clb_selectedPipesSelectedIndexChanged(self, sender, e):
		pass

	def Ckb_AllNoneCheckedChanged(self, sender, e):
		pass

	def Btt_setClick(self, sender, e):
		pass

	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass
	
f = MainForm()
Application.Run(f)
