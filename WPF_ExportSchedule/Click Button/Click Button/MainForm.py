import clr
import sys 
import System   
clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
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

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._bttGetXY = System.Windows.Forms.Button()
		self._bttGetZ = System.Windows.Forms.Button()
		self._bttOK = System.Windows.Forms.Button()
		self._bttCANCLE = System.Windows.Forms.Button()
		self._tbXY = System.Windows.Forms.TextBox()
		self.SuspendLayout()
		# 
		# bttGetXY
		# 
		self._bttGetXY.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._bttGetXY.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128, True)
		self._bttGetXY.Location = System.Drawing.Point(32, 27)
		self._bttGetXY.Name = "bttGetXY"
		self._bttGetXY.Size = System.Drawing.Size(151, 101)
		self._bttGetXY.TabIndex = 0
		self._bttGetXY.Text = "get X,Y value"
		self._bttGetXY.UseVisualStyleBackColor = True
		self._bttGetXY.Click += self.BttGetXYClick
		# 
		# bttGetZ
		# 
		self._bttGetZ.Location = System.Drawing.Point(269, 27)
		self._bttGetZ.Name = "bttGetZ"
		self._bttGetZ.Size = System.Drawing.Size(151, 101)
		self._bttGetZ.TabIndex = 1
		self._bttGetZ.Text = "get Z value"
		self._bttGetZ.UseVisualStyleBackColor = True
		self._bttGetZ.Click += self.BttGetZClick
		# 
		# bttOK
		# 
		self._bttOK.BackColor = System.Drawing.SystemColors.ActiveCaption
		self._bttOK.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._bttOK.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128, True)
		self._bttOK.ForeColor = System.Drawing.Color.Red
		self._bttOK.Location = System.Drawing.Point(174, 335)
		self._bttOK.Name = "bttOK"
		self._bttOK.Size = System.Drawing.Size(100, 29)
		self._bttOK.TabIndex = 2
		self._bttOK.Text = "OK"
		self._bttOK.UseVisualStyleBackColor = False
		self._bttOK.Click += self.BttOKClick
		# 
		# bttCANCLE
		# 
		self._bttCANCLE.BackColor = System.Drawing.SystemColors.ActiveCaption
		self._bttCANCLE.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._bttCANCLE.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128, True)
		self._bttCANCLE.ForeColor = System.Drawing.Color.Red
		self._bttCANCLE.Location = System.Drawing.Point(280, 335)
		self._bttCANCLE.Name = "bttCANCLE"
		self._bttCANCLE.Size = System.Drawing.Size(100, 29)
		self._bttCANCLE.TabIndex = 3
		self._bttCANCLE.Text = "CANCLE"
		self._bttCANCLE.UseVisualStyleBackColor = False
		self._bttCANCLE.Click += self.BttCANCLEClick
		# 
		# tbXY
		# 
		self._tbXY.Location = System.Drawing.Point(32, 153)
		self._tbXY.Name = "tbXY"
		self._tbXY.Size = System.Drawing.Size(203, 22)
		self._tbXY.TabIndex = 4
		self._tbXY.TextChanged += self.TbXYTextChanged
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(471, 451)
		self.Controls.Add(self._tbXY)
		self.Controls.Add(self._bttCANCLE)
		self.Controls.Add(self._bttOK)
		self.Controls.Add(self._bttGetZ)
		self.Controls.Add(self._bttGetXY)
		self.Name = "MainForm"
		self.Text = "ClickButton"
		self.ResumeLayout(False)
		self.PerformLayout()


	def BttGetXYClick(self, sender, e):
		pass
	
	def BttGetZClick(self, sender, e):
		pass

	def BttOKClick(self, sender, e):
		pass

	def BttCANCLEClick(self, sender, e):
		pass

	def TbXYTextChanged(self, sender, e):
		pass
	
f = MainForm()
Application.Run(f)