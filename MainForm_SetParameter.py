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
		self._lb_TrueLength = System.Windows.Forms.Label()
		self._txb_TrueLengthName = System.Windows.Forms.TextBox()
		self._btt_PipeTag = System.Windows.Forms.Button()
		self._btt_TrueLength = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self.SuspendLayout()
		# 
		# cbb_PipeTag
		# 
		self._cbb_PipeTag.FormattingEnabled = True
		self._cbb_PipeTag.Location = System.Drawing.Point(122, 37)
		self._cbb_PipeTag.Name = "cbb_PipeTag"
		self._cbb_PipeTag.Size = System.Drawing.Size(229, 23)
		self._cbb_PipeTag.TabIndex = 0
		self._cbb_PipeTag.SelectedIndexChanged += self.Cbb_PipeTagSelectedIndexChanged
		# 
		# lb_PipeTag
		# 
		self._lb_PipeTag.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
		self._lb_PipeTag.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_PipeTag.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._lb_PipeTag.Location = System.Drawing.Point(13, 37)
		self._lb_PipeTag.Name = "lb_PipeTag"
		self._lb_PipeTag.Size = System.Drawing.Size(83, 23)
		self._lb_PipeTag.TabIndex = 1
		self._lb_PipeTag.Text = "Pipe Tag"
		# 
		# lb_TrueLength
		# 
		self._lb_TrueLength.BackColor = System.Drawing.Color.FromArgb(255, 255, 192)
		self._lb_TrueLength.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_TrueLength.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._lb_TrueLength.Location = System.Drawing.Point(13, 96)
		self._lb_TrueLength.Name = "lb_TrueLength"
		self._lb_TrueLength.Size = System.Drawing.Size(94, 49)
		self._lb_TrueLength.TabIndex = 2
		self._lb_TrueLength.Text = "芯々寸法 Parameter"
		# 
		# txb_TrueLengthName
		# 
		self._txb_TrueLengthName.Location = System.Drawing.Point(122, 109)
		self._txb_TrueLengthName.Name = "txb_TrueLengthName"
		self._txb_TrueLengthName.Size = System.Drawing.Size(229, 22)
		self._txb_TrueLengthName.TabIndex = 3
		self._txb_TrueLengthName.TextChanged += self.Txb_TrueLengthNameTextChanged
		# 
		# btt_PipeTag
		# 
		self._btt_PipeTag.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_PipeTag.Location = System.Drawing.Point(22, 199)
		self._btt_PipeTag.Name = "btt_PipeTag"
		self._btt_PipeTag.Size = System.Drawing.Size(94, 53)
		self._btt_PipeTag.TabIndex = 4
		self._btt_PipeTag.Text = "Set Pipe Tag"
		self._btt_PipeTag.UseVisualStyleBackColor = True
		self._btt_PipeTag.Click += self.Btt_PipeTagClick
		# 
		# btt_TrueLength
		# 
		self._btt_TrueLength.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_TrueLength.Location = System.Drawing.Point(139, 199)
		self._btt_TrueLength.Name = "btt_TrueLength"
		self._btt_TrueLength.Size = System.Drawing.Size(94, 53)
		self._btt_TrueLength.TabIndex = 5
		self._btt_TrueLength.Text = "芯々寸法"
		self._btt_TrueLength.UseVisualStyleBackColor = True
		self._btt_TrueLength.Click += self.Btt_TrueLengthClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.Location = System.Drawing.Point(254, 199)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(94, 53)
		self._btt_CANCLE.TabIndex = 6
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(365, 276)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_TrueLength)
		self.Controls.Add(self._btt_PipeTag)
		self.Controls.Add(self._txb_TrueLengthName)
		self.Controls.Add(self._lb_TrueLength)
		self.Controls.Add(self._lb_PipeTag)
		self.Controls.Add(self._cbb_PipeTag)
		self.Name = "MainForm"
		self.Text = "Set Parameter"
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