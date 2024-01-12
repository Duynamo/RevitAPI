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


inValue = UnwrapElement(IN[0])


class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._button1 = System.Windows.Forms.Button()
		self._button2 = System.Windows.Forms.Button()
		self._buttonOK = System.Windows.Forms.Button()
		self._buttonCANCLE = System.Windows.Forms.Button()
		self.SuspendLayout()
		# 
		# button1
		# 
		self._button1.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._button1.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128, True)
		self._button1.Location = System.Drawing.Point(32, 27)
		self._button1.Name = "button1"
		self._button1.Size = System.Drawing.Size(151, 101)
		self._button1.TabIndex = 0
		self._button1.Text = "get X,Y value"
		self._button1.UseVisualStyleBackColor = True
		self._button1.Click += self.Button1Click
		# 
		# button2
		# 
		self._button2.Location = System.Drawing.Point(269, 27)
		self._button2.Name = "button2"
		self._button2.Size = System.Drawing.Size(151, 101)
		self._button2.TabIndex = 1
		self._button2.Text = "get Z value"
		self._button2.UseVisualStyleBackColor = True
		self._button2.Click += self.Button2Click
		# 
		# buttonOK
		# 
		self._buttonOK.BackColor = System.Drawing.SystemColors.ActiveCaption
		self._buttonOK.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._buttonOK.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128, True)
		self._buttonOK.ForeColor = System.Drawing.Color.Red
		self._buttonOK.Location = System.Drawing.Point(174, 335)
		self._buttonOK.Name = "buttonOK"
		self._buttonOK.Size = System.Drawing.Size(100, 29)
		self._buttonOK.TabIndex = 2
		self._buttonOK.Text = "OK"
		self._buttonOK.UseVisualStyleBackColor = False
		self._buttonOK.Click += self.buttonOK
		# 
		# buttonCANCLE
		# 
		self._buttonCANCLE.BackColor = System.Drawing.SystemColors.ActiveCaption
		self._buttonCANCLE.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._buttonCANCLE.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128, True)
		self._buttonCANCLE.ForeColor = System.Drawing.Color.Red
		self._buttonCANCLE.Location = System.Drawing.Point(280, 335)
		self._buttonCANCLE.Name = "buttonCANCLE"
		self._buttonCANCLE.Size = System.Drawing.Size(100, 29)
		self._buttonCANCLE.TabIndex = 3
		self._buttonCANCLE.Text = "CANCLE"
		self._buttonCANCLE.UseVisualStyleBackColor = False
		self._buttonCANCLE.Click += self.buttonCANCLE
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(471, 451)
		self.Controls.Add(self._buttonCANCLE)
		self.Controls.Add(self._buttonOK)
		self.Controls.Add(self._button2)
		self.Controls.Add(self._button1)
		self.Name = "MainForm"
		self.Text = "ClickButton"
		self.ResumeLayout(False)


	def Button1Click(self, sender, e):
		TransactionManager.Instance.EnsureInTransaction(doc)
	
		x=True
		dypoint = []
		rpointM = []
		rpointI = []
		counter=0
		msg = 'Pick Points on current Workplane in order, hit ESC when finished.'

		TaskDialog.Show("Duynamo", msg)

		while x == True and inValue == True:
			# msg = 'Pick Points on current Workplane in order, hit ESC when finished.'
			# TaskDialog.Show("Duynamo", msg)
			try:
				pt=uidoc.Selection.PickPoint()
				rpM=Point.ByCoordinates(pt.X*304.8,pt.Y*304.8,pt.Z*304.8)
				rpI=Point.ByCoordinates(pt.X,pt.Y,pt.Z)
				counter += 1
				dypoint.append(pt)
				rpointM.append(rpM)
				rpointI.append(rpI)
			except:
				x=False
		
		TransactionManager.Instance.TransactionTaskDone()
	

		pass

	def Button2Click(self, sender, e):
		pass

	def buttonOK(self, sender, e):
		pass

	def buttonCANCLE(self, sender, e):
		self.Close()
		pass
	

f = MainForm()
Application.Run(f)