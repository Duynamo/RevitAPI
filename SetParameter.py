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

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._cbb_PipeTag = System.Windows.Forms.ComboBox()
		self._lb_PipeTag = System.Windows.Forms.Label()
		self._lb_CenterDim = System.Windows.Forms.Label()
		self._txb_TrueLengthName = System.Windows.Forms.TextBox()
		self._btt_PipeTag = System.Windows.Forms.Button()
		self._btt_CenterDim = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._lb_FVC = System.Windows.Forms.Label()
		self.SuspendLayout()
		# 
		# cbb_PipeTag
		# 
		self._cbb_PipeTag.Cursor = System.Windows.Forms.Cursors.Default
		self._cbb_PipeTag.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cbb_PipeTag.ForeColor = System.Drawing.Color.Red
		self._cbb_PipeTag.FormattingEnabled = True
		self._cbb_PipeTag.Location = System.Drawing.Point(122, 37)
		self._cbb_PipeTag.Name = "cbb_PipeTag"
		self._cbb_PipeTag.Size = System.Drawing.Size(229, 27)
		self._cbb_PipeTag.TabIndex = 0
		self._cbb_PipeTag.SelectedIndexChanged += self.Cbb_PipeTagSelectedIndexChanged
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
		# txb_TrueLengthName
		# 
		self._txb_TrueLengthName.ForeColor = System.Drawing.Color.Red
		self._txb_TrueLengthName.Location = System.Drawing.Point(122, 125)
		self._txb_TrueLengthName.Name = "txb_TrueLengthName"
		self._txb_TrueLengthName.Size = System.Drawing.Size(229, 27)
		self._txb_TrueLengthName.TabIndex = 3
		self._txb_TrueLengthName.TextChanged += self.Txb_TrueLengthNameTextChanged
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
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(365, 258)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_CenterDim)
		self.Controls.Add(self._btt_PipeTag)
		self.Controls.Add(self._txb_TrueLengthName)
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
		self.PerformLayout()


	def Cbb_PipeTagSelectedIndexChanged(self, sender, e):
		pass

	def Txb_TrueLengthNameTextChanged(self, sender, e):
		pass

	def Btt_PipeTagClick(self, sender, e):
		pass

	def Btt_TrueLengthClick(self, sender, e):
		pass

	def Btt_CANCLEClick(self, sender, e):
		pass

f = MainForm()
Application.Run(f)
