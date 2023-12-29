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



def getAllPipingSystems(doc):
	collector = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PipingSystem)
	pipingSystems = collector.ToElements()
	pipingSystemsName = []
	for system in pipingSystems:
		systemName = system.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
		pipingSystemsName.append(systemName)
	return pipingSystemsName

pipingSystemsCollector = getAllPipingSystems(doc)

def getAllPipeTypes(doc):
	collector1 = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_PipeCurves)
	pipeTypes = collector1.ToElements()
	pipeTypesName = []
	for pipeType in pipeTypes:
		pipeTypeName = pipeType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
		pipeTypesName.append(pipeTypeName)
	return pipeTypesName

pipeTypesCollector = getAllPipeTypes(doc)

levelsCollector = FilteredElementCollector(doc).OfClass(Level).ToElements()
levelsNameCollector = []
for level in levelsCollector:
	levelName = level.Name
	levelsNameCollector.append(levelName)

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._groupBox1 = System.Windows.Forms.GroupBox()
		self._button1 = System.Windows.Forms.Button()
		self._button2 = System.Windows.Forms.Button()
		self._comboBox1 = System.Windows.Forms.ComboBox()
		self._label1 = System.Windows.Forms.Label()
		self._groupBox2 = System.Windows.Forms.GroupBox()
		self._label2 = System.Windows.Forms.Label()
		self._comboBox2 = System.Windows.Forms.ComboBox()
		self._groupBox3 = System.Windows.Forms.GroupBox()
		self._label3 = System.Windows.Forms.Label()
		self._comboBox3 = System.Windows.Forms.ComboBox()
		self._groupBox4 = System.Windows.Forms.GroupBox()
		self._label4 = System.Windows.Forms.Label()
		self._textBox1 = System.Windows.Forms.TextBox()
		self._button3 = System.Windows.Forms.Button()
		self._button4 = System.Windows.Forms.Button()
		self._label5 = System.Windows.Forms.Label()
		self._groupBox1.SuspendLayout()
		self._groupBox2.SuspendLayout()
		self._groupBox3.SuspendLayout()
		self._groupBox4.SuspendLayout()
		self.SuspendLayout()
		# 
		# groupBox1
		# 
		self._groupBox1.Controls.Add(self._label1)
		self._groupBox1.Controls.Add(self._comboBox1)
		self._groupBox1.Location = System.Drawing.Point(12, 159)
		self._groupBox1.Name = "groupBox1"
		self._groupBox1.Size = System.Drawing.Size(449, 47)
		self._groupBox1.TabIndex = 0
		self._groupBox1.TabStop = False
		# 
		# button1
		# 
		self._button1.BackColor = System.Drawing.SystemColors.Info
		self._button1.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, 128, True)
		self._button1.ForeColor = System.Drawing.Color.Red
		self._button1.Location = System.Drawing.Point(52, 28)
		self._button1.Name = "button1"
		self._button1.Size = System.Drawing.Size(140, 58)
		self._button1.TabIndex = 1
		self._button1.Text = "get X,Y value"
		self._button1.UseVisualStyleBackColor = False
		self._button1.Click += self.Button1Click
		# 
		# button2
		# 
		self._button2.BackColor = System.Drawing.SystemColors.Info
		self._button2.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._button2.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, 128, True)
		self._button2.ForeColor = System.Drawing.Color.Red
		self._button2.Location = System.Drawing.Point(279, 28)
		self._button2.Name = "button2"
		self._button2.Size = System.Drawing.Size(140, 58)
		self._button2.TabIndex = 2
		self._button2.Text = "get Z value"
		self._button2.UseVisualStyleBackColor = False
		# 
		# comboBox1
		# 
		self._comboBox1.AllowDrop = True
		self._comboBox1.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._comboBox1.FormattingEnabled = True
		self._comboBox1.Location = System.Drawing.Point(180, 12)
		self._comboBox1.Name = "comboBox1"
		self._comboBox1.Size = System.Drawing.Size(263, 23)
		self._comboBox1.TabIndex = 3
		# 
		# label1
		# 
		self._label1.BackColor = System.Drawing.SystemColors.Info
		self._label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._label1.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._label1.ForeColor = System.Drawing.Color.Blue
		self._label1.Location = System.Drawing.Point(19, 12)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(142, 23)
		self._label1.TabIndex = 4
		self._label1.Text = "PipingSystemType"
		# 
		# groupBox2
		# 
		self._groupBox2.Controls.Add(self._label2)
		self._groupBox2.Controls.Add(self._comboBox2)
		self._groupBox2.Location = System.Drawing.Point(12, 212)
		self._groupBox2.Name = "groupBox2"
		self._groupBox2.Size = System.Drawing.Size(449, 47)
		self._groupBox2.TabIndex = 5
		self._groupBox2.TabStop = False
		# 
		# label2
		# 
		self._label2.BackColor = System.Drawing.SystemColors.Info
		self._label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._label2.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._label2.ForeColor = System.Drawing.Color.Blue
		self._label2.Location = System.Drawing.Point(19, 12)
		self._label2.Name = "label2"
		self._label2.Size = System.Drawing.Size(142, 23)
		self._label2.TabIndex = 4
		self._label2.Text = "PipeType"
		self._label2.Click += self.Label2Click
		# 
		# comboBox2
		# 
		self._comboBox2.AllowDrop = True
		self._comboBox2.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._comboBox2.FormattingEnabled = True
		self._comboBox2.Location = System.Drawing.Point(180, 12)
		self._comboBox2.Name = "comboBox2"
		self._comboBox2.Size = System.Drawing.Size(263, 23)
		self._comboBox2.TabIndex = 3
		# 
		# groupBox3
		# 
		self._groupBox3.Controls.Add(self._label3)
		self._groupBox3.Controls.Add(self._comboBox3)
		self._groupBox3.Location = System.Drawing.Point(12, 265)
		self._groupBox3.Name = "groupBox3"
		self._groupBox3.Size = System.Drawing.Size(449, 47)
		self._groupBox3.TabIndex = 6
		self._groupBox3.TabStop = False
		# 
		# label3
		# 
		self._label3.BackColor = System.Drawing.SystemColors.Info
		self._label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._label3.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._label3.ForeColor = System.Drawing.Color.Blue
		self._label3.Location = System.Drawing.Point(19, 12)
		self._label3.Name = "label3"
		self._label3.Size = System.Drawing.Size(142, 23)
		self._label3.TabIndex = 4
		self._label3.Text = "RefLevel"
		# 
		# comboBox3
		# 
		self._comboBox3.AllowDrop = True
		self._comboBox3.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._comboBox3.FormattingEnabled = True
		self._comboBox3.Location = System.Drawing.Point(180, 12)
		self._comboBox3.Name = "comboBox3"
		self._comboBox3.Size = System.Drawing.Size(263, 23)
		self._comboBox3.TabIndex = 3
		# 
		# groupBox4
		# 
		self._groupBox4.Controls.Add(self._textBox1)
		self._groupBox4.Controls.Add(self._label4)
		self._groupBox4.Location = System.Drawing.Point(12, 318)
		self._groupBox4.Name = "groupBox4"
		self._groupBox4.Size = System.Drawing.Size(449, 47)
		self._groupBox4.TabIndex = 7
		self._groupBox4.TabStop = False
		# 
		# label4
		# 
		self._label4.BackColor = System.Drawing.SystemColors.Info
		self._label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._label4.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._label4.ForeColor = System.Drawing.Color.Blue
		self._label4.Location = System.Drawing.Point(19, 12)
		self._label4.Name = "label4"
		self._label4.Size = System.Drawing.Size(142, 23)
		self._label4.TabIndex = 4
		self._label4.Text = "Diameter"
		# 
		# textBox1
		# 
		self._textBox1.Location = System.Drawing.Point(180, 12)
		self._textBox1.Name = "textBox1"
		self._textBox1.Size = System.Drawing.Size(263, 22)
		self._textBox1.TabIndex = 5
		# 
		# button3
		# 
		self._button3.BackColor = System.Drawing.SystemColors.ScrollBar
		self._button3.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._button3.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._button3.ForeColor = System.Drawing.Color.Red
		self._button3.Location = System.Drawing.Point(317, 395)
		self._button3.Name = "button3"
		self._button3.Size = System.Drawing.Size(49, 30)
		self._button3.TabIndex = 8
		self._button3.Text = "OK"
		self._button3.UseVisualStyleBackColor = False
		# 
		# button4
		# 
		self._button4.BackColor = System.Drawing.SystemColors.ScrollBar
		self._button4.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._button4.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._button4.ForeColor = System.Drawing.Color.Red
		self._button4.Location = System.Drawing.Point(372, 395)
		self._button4.Name = "button4"
		self._button4.Size = System.Drawing.Size(89, 30)
		self._button4.TabIndex = 8
		self._button4.Text = "CANCLE"
		self._button4.UseVisualStyleBackColor = False
		# 
		# label5
		# 
		self._label5.Location = System.Drawing.Point(1, 434)
		self._label5.Name = "label5"
		self._label5.Size = System.Drawing.Size(53, 19)
		self._label5.TabIndex = 9
		self._label5.Text = "@FVC"
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(482, 453)
		self.Controls.Add(self._label5)
		self.Controls.Add(self._button4)
		self.Controls.Add(self._button3)
		self.Controls.Add(self._groupBox4)
		self.Controls.Add(self._groupBox3)
		self.Controls.Add(self._groupBox2)
		self.Controls.Add(self._button2)
		self.Controls.Add(self._button1)
		self.Controls.Add(self._groupBox1)
		self.Name = "MainForm"
		self.Text = "PipeCreate"
		self._groupBox1.ResumeLayout(False)
		self._groupBox2.ResumeLayout(False)
		self._groupBox3.ResumeLayout(False)
		self._groupBox4.ResumeLayout(False)
		self._groupBox4.PerformLayout()
		self.ResumeLayout(False)


	def Label2Click(self, sender, e):
		pass


	def ComboBox1_SelectedIndexChanged(self, sender, e):
		TransactionManager.Instance.EnsureInTransaction(doc)



		TransactionManager.Instance.TransactionTaskDone()
		pass
	
	def ComboBox2_SelectedIndexChanged(self, sender, e):
		TransactionManager.Instance.EnsureInTransaction(doc)



		TransactionManager.Instance.TransactionTaskDone()
		pass	

	def ComboBox3_SelectedIndexChanged(self, sender, e):
		TransactionManager.Instance.EnsureInTransaction(doc)



		TransactionManager.Instance.TransactionTaskDone()
		pass



f = MainForm()
Application.Run(f)

	def Button1Click(self, sender, e):
		pass