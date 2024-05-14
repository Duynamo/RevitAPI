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
"""________________________________________________________________________________________"""
def getAllPipingSystemsName(doc):
	collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem)
	pipingSystems = collector.ToElements()
	pipingSystemsName = []
	for system in pipingSystems:
		systemName = system.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
		pipingSystemsName.append(systemName)
	return pipingSystemsName
lst = getAllPipingSystemsName(doc)
pipingSystemsCollector = [item for item in lst if item is not None]

def getAllPipingSystems(doc):
	collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem)
	pipingSystems = collector.ToElements()
	return pipingSystems
"""_______________________________________________________________________________________"""
def getAllPipeTypesName(doc):
	collector1 = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves)
	pipeTypes = collector1.ToElements()
	pipeTypesName = []
	for pipeType in pipeTypes:
		pipeTypeName = pipeType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
		pipeTypesName.append(pipeTypeName)
	return pipeTypesName
lst1 = getAllPipeTypesName(doc)
pipeTypesCollector = [item for item in lst1 if item is not None]
"""_______________________________________________________________________________________"""
def createModelText(doc, points, text_size, start_value=1):
	modelTextsList = []
	if not isinstance(points, list) or len(points) == 0:
		raise ValueError("Invalid input for points.")
	for i, point in enumerate(points):
		number = start_value + i
		textValue = "Point " + str(number)
		textSize = 15
		modelTextsList.append(i for i in createModelText(doc, point, textValue, textSize))
	return modelTextsList	
"""_______________________________________________________________________________________"""
def getAllPipeTypes(doc):
	collector1 = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves)
	pipeTypes = collector1.ToElements()
	return pipeTypes
"""______________________________________________________________________________________"""

levelsCollector = FilteredElementCollector(doc).OfClass(Level).ToElements()
"""________________________________________________________________________________________"""
def logger(title, content):
	import datetime
	date = datetime.datetime.now()
	f = open(r"C:\Users\95053\Desktop\Python\RevitAPI-master\RevitAPI\python.log", 'a')
	# C:\Users\95053\Desktop\Python\RevitAPI-master\RevitAPI\PipeCreate.py
	f.write(str(date) + '\n' + title + '\n' + str(content) + '\n')
	f.close()
"""________________________________________________________________________________________"""
class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	#________________________________________________________________________________________#	
	def InitializeComponent(self):
		self._groupBox1 = System.Windows.Forms.GroupBox()
		self._btt_getXY = System.Windows.Forms.Button()
		self._btt_getZ = System.Windows.Forms.Button()
		self._cbb_PipingSystemType = System.Windows.Forms.ComboBox()
		self._label1 = System.Windows.Forms.Label()
		self._groupBox2 = System.Windows.Forms.GroupBox()
		self._label2 = System.Windows.Forms.Label()
		self._cbb_PipeType = System.Windows.Forms.ComboBox()
		self._groupBox3 = System.Windows.Forms.GroupBox()
		self._label3 = System.Windows.Forms.Label()
		self._cbb_RefLevel = System.Windows.Forms.ComboBox()
		self._groupBox4 = System.Windows.Forms.GroupBox()
		self._label4 = System.Windows.Forms.Label()
		self._txb_Diameter = System.Windows.Forms.TextBox()
		self._btt_CREATE = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._label5 = System.Windows.Forms.Label()
		self._groupBox5 = System.Windows.Forms.GroupBox()
		self._clb_XYValue = System.Windows.Forms.CheckedListBox()
		self._clb_ZValue = System.Windows.Forms.CheckedListBox()
		self._cbb_AllXY = System.Windows.Forms.CheckBox()
		self._cbb_AllZ = System.Windows.Forms.CheckBox()
		self._total_XYValue = System.Windows.Forms.TextBox()
		self._total_ZValue = System.Windows.Forms.TextBox()
		self._label6 = System.Windows.Forms.Label()
		self._label7 = System.Windows.Forms.Label()
		self._btt_LoopZ = System.Windows.Forms.Button()		
		self._btt_ResetXYValue = System.Windows.Forms.Button()
		self._btt_ResetZValue = System.Windows.Forms.Button()		
		self._groupBox1.SuspendLayout()
		self._groupBox2.SuspendLayout()
		self._groupBox3.SuspendLayout()
		self._groupBox4.SuspendLayout()
		self._groupBox5.SuspendLayout()
		self.SuspendLayout()
		# 
		# groupBox1
		# 
		self._groupBox1.Controls.Add(self._label1)
		self._groupBox1.Controls.Add(self._cbb_PipingSystemType)
		self._groupBox1.Location = System.Drawing.Point(12, 234)
		self._groupBox1.Name = "groupBox1"
		self._groupBox1.Size = System.Drawing.Size(449, 47)
		self._groupBox1.TabIndex = 0
		self._groupBox1.TabStop = False
		# 
		# btt_getXY
		# 
		# self._btt_getXY.BackColor = System.Drawing.SystemColors.ScrollBar
		self._btt_getXY.FlatAppearance.BorderColor = System.Drawing.Color.Red		
		self._btt_getXY.Cursor = System.Windows.Forms.Cursors.AppStarting	
		self._btt_getXY.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Yellow
		self._btt_getXY.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Cyan			
		self._btt_getXY.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_getXY.ForeColor = System.Drawing.Color.Red
		self._btt_getXY.Location = System.Drawing.Point(55, 8)
		self._btt_getXY.Name = "btt_getXY"
		self._btt_getXY.Size = System.Drawing.Size(140, 30)
		self._btt_getXY.TabIndex = 1
		self._btt_getXY.Text = "get X,Y value"
		self._btt_getXY.UseVisualStyleBackColor = False
		self._btt_getXY.Click += self.Btt_getXYClick
		# 
		# btt_getZ
		# 
		# self._btt_getZ.BackColor = System.Drawing.SystemColors.ScrollBar
		self._btt_getZ.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_getZ.FlatAppearance.BorderColor = System.Drawing.Color.Red		
		self._btt_getZ.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Yellow
		self._btt_getZ.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Cyan
		self._btt_getZ.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_getZ.ForeColor = System.Drawing.Color.Red
		self._btt_getZ.Location = System.Drawing.Point(250, 8)
		self._btt_getZ.Name = "btt_getZ"
		self._btt_getZ.Size = System.Drawing.Size(110, 30)
		self._btt_getZ.TabIndex = 2
		self._btt_getZ.Text = "get Z value"
		self._btt_getZ.UseVisualStyleBackColor = False
		self._btt_getZ.Click += self.Btt_getZClick
		# 
		# cbb_PipingSystemType
		# 
		self._cbb_PipingSystemType.DisplayMember = "Name"		
		self._cbb_PipingSystemType.AllowDrop = True
		self._cbb_PipingSystemType.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._cbb_PipingSystemType.FormattingEnabled = True
		self._cbb_PipingSystemType.Location = System.Drawing.Point(180, 12)
		self._cbb_PipingSystemType.Name = "cbb_PipingSystemType"
		self._cbb_PipingSystemType.Size = System.Drawing.Size(263, 23)
		self._cbb_PipingSystemType.TabIndex = 3
		self._cbb_PipingSystemType.SelectedIndexChanged += self.Cbb_PipingSystemTypeSelectedIndexChanged
		self._cbb_PipingSystemType.Items.AddRange(System.Array[System.Object](pipingSystemsCollector)) 
		self._cbb_PipingSystemType.SelectedIndex = 0		
		# 
		# label1
		# 
		self._label1.BackColor = System.Drawing.SystemColors.Info
		self._label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._label1.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._label1.ForeColor = System.Drawing.Color.Red
		self._label1.Location = System.Drawing.Point(6, 14)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(142, 23)
		self._label1.TabIndex = 4
		self._label1.Text = "PipingSystemType"
		# 
		# groupBox2
		# 
		self._groupBox2.Controls.Add(self._label2)
		self._groupBox2.Controls.Add(self._cbb_PipeType)
		self._groupBox2.Location = System.Drawing.Point(12, 287)
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
		self._label2.ForeColor = System.Drawing.Color.Red
		self._label2.Location = System.Drawing.Point(6, 12)
		self._label2.Name = "label2"
		self._label2.Size = System.Drawing.Size(142, 23)
		self._label2.TabIndex = 4
		self._label2.Text = "PipeType"
		# 
		# cbb_PipeType
		# 
		self._cbb_PipeType.DisplayMember = "Name"		
		self._cbb_PipeType.AllowDrop = True
		self._cbb_PipeType.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._cbb_PipeType.FormattingEnabled = True
		self._cbb_PipeType.Location = System.Drawing.Point(180, 12)
		self._cbb_PipeType.Name = "cbb_PipeType"
		self._cbb_PipeType.Size = System.Drawing.Size(263, 23)
		self._cbb_PipeType.TabIndex = 3
		self._cbb_PipeType.SelectedIndexChanged += self.Cbb_PipeTypeSelectedIndexChanged
		self._cbb_PipeType.Items.AddRange(System.Array[System.Object](pipeTypesCollector))
		self._cbb_PipeType.SelectedIndex = 0		
		# 
		# groupBox3
		# 
		self._groupBox3.Controls.Add(self._label3)
		self._groupBox3.Controls.Add(self._cbb_RefLevel)
		self._groupBox3.Location = System.Drawing.Point(12, 340)
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
		self._label3.ForeColor = System.Drawing.Color.Red
		self._label3.Location = System.Drawing.Point(6, 12)
		self._label3.Name = "label3"
		self._label3.Size = System.Drawing.Size(142, 23)
		self._label3.TabIndex = 4
		self._label3.Text = "RefLevel"
		# 
		# cbb_RefLevel
		# 
		self._cbb_RefLevel.DisplayMember = "Name"		
		self._cbb_RefLevel.AllowDrop = True
		self._cbb_RefLevel.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._cbb_RefLevel.FormattingEnabled = True
		self._cbb_RefLevel.Location = System.Drawing.Point(180, 12)
		self._cbb_RefLevel.Name = "cbb_RefLevel"
		self._cbb_RefLevel.Size = System.Drawing.Size(263, 23)
		self._cbb_RefLevel.TabIndex = 3
		self._cbb_RefLevel.SelectedIndexChanged += self.Cbb_RefLevelSelectedIndexChanged
		self._cbb_RefLevel.Items.AddRange(System.Array[System.Object](levelsCollector))
		self._cbb_RefLevel.SelectedIndex = 0				
		# 
		# groupBox4
		# 
		self._groupBox4.Controls.Add(self._txb_Diameter)
		self._groupBox4.Controls.Add(self._label4)
		self._groupBox4.Location = System.Drawing.Point(12, 393)
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
		self._label4.ForeColor = System.Drawing.Color.Red
		self._label4.Location = System.Drawing.Point(6, 12)
		self._label4.Name = "label4"
		self._label4.Size = System.Drawing.Size(142, 23)
		self._label4.TabIndex = 4
		self._label4.Text = "Diameter"
		# 
		# txb_Diameter
		# 
		self._txb_Diameter.Location = System.Drawing.Point(180, 12)
		self._txb_Diameter.Name = "txb_Diameter"
		self._txb_Diameter.Size = System.Drawing.Size(263, 22)
		self._txb_Diameter.TabIndex = 5
		self._txb_Diameter.TextChanged += self.Txb_DiameterTextChanged
		# 
		# CREATE
		# 
		# self._btt_CREATE.BackColor = System.Drawing.SystemColors.ScrollBar
		self._btt_CREATE.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._btt_CREATE.FlatAppearance.BorderColor = System.Drawing.Color.Red		
		self._btt_CREATE.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Yellow
		self._btt_CREATE.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Cyan		
		self._btt_CREATE.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_CREATE.ForeColor = System.Drawing.Color.Red
		self._btt_CREATE.Location = System.Drawing.Point(262, 464)
		self._btt_CREATE.Name = "btt_CREATE"
		self._btt_CREATE.Size = System.Drawing.Size(104, 30)
		self._btt_CREATE.TabIndex = 8
		self._btt_CREATE.Text = "CREATE"
		self._btt_CREATE.UseVisualStyleBackColor = False
		self._btt_CREATE.Click += self.Btt_CREATEClick
		# 
		# btt_CANCLE
		# 
		# self._btt_CANCLE.BackColor = System.Drawing.SystemColors.ScrollBar
		self._btt_CANCLE.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._btt_CANCLE.FlatAppearance.BorderColor = System.Drawing.Color.Red		
		self._btt_CANCLE.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Yellow
		self._btt_CANCLE.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Cyan	
		self._btt_CANCLE.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
		self._btt_CANCLE.Location = System.Drawing.Point(372, 464)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(89, 30)
		self._btt_CANCLE.TabIndex = 8
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = False
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# label5
		# 
		self._label5.Location = System.Drawing.Point(15, 480)
		self._label5.Name = "label5"
		self._label5.Size = System.Drawing.Size(53, 19)
		self._label5.TabIndex = 9
		self._label5.Text = "@FVC"
		# 
		# groupBox5
		# 
		self._groupBox5.Controls.Add(self._btt_ResetZValue)
		self._groupBox5.Controls.Add(self._btt_ResetXYValue)		
		self._groupBox5.Controls.Add(self._label7)
		self._groupBox5.Controls.Add(self._label6)
		self._groupBox5.Controls.Add(self._total_ZValue)
		self._groupBox5.Controls.Add(self._total_XYValue)
		self._groupBox5.Controls.Add(self._cbb_AllZ)
		self._groupBox5.Controls.Add(self._cbb_AllXY)
		self._groupBox5.Controls.Add(self._clb_ZValue)
		self._groupBox5.Controls.Add(self._clb_XYValue)
		self._groupBox5.Location = System.Drawing.Point(13, 44)
		self._groupBox5.Name = "groupBox5"
		self._groupBox5.Size = System.Drawing.Size(457, 184)
		self._groupBox5.TabIndex = 10
		self._groupBox5.TabStop = False
		# self._groupBox5.Enter += self.GroupBox5Enter
		# 
		# clb_XYValue
		# 
		self._clb_XYValue.AllowDrop = True
		self._clb_XYValue.CheckOnClick = True
		self._clb_XYValue.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_XYValue.FormattingEnabled = True
		self._clb_XYValue.HorizontalScrollbar = True
		self._clb_XYValue.Location = System.Drawing.Point(0, 40)
		self._clb_XYValue.Name = "clb_XYValue"
		self._clb_XYValue.Size = System.Drawing.Size(213, 115)
		self._clb_XYValue.TabIndex = 0
		self._clb_XYValue.SelectedIndexChanged += self.Clb_XYValueSelectedIndexChanged

		# 
		# clb_ZValue
		# 
		self._clb_ZValue.AllowDrop = True
		self._clb_ZValue.CheckOnClick = True
		self._clb_ZValue.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_ZValue.FormattingEnabled = True
		self._clb_ZValue.HorizontalScrollbar = True
		self._clb_ZValue.Location = System.Drawing.Point(232, 40)
		self._clb_ZValue.Name = "clb_ZValue"
		self._clb_ZValue.Size = System.Drawing.Size(216, 115)
		self._clb_ZValue.TabIndex = 1
		self._clb_ZValue.SelectedIndexChanged += self.Clb_ZValueSelectedIndexChanged
		
		# 
		# cbb_AllXY
		# 
		self._cbb_AllXY.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._cbb_AllXY.Location = System.Drawing.Point(2, 12)
		self._cbb_AllXY.Name = "cbb_AllXY"
		self._cbb_AllXY.Size = System.Drawing.Size(114, 24)
		self._cbb_AllXY.TabIndex = 2
		self._cbb_AllXY.Text = "All/None"
		self._cbb_AllXY.UseVisualStyleBackColor = True
		# self._cbb_AllXY.Checked = True
		self._cbb_AllXY.CheckedChanged += self.Cb_AllXYCheckedChanged
		# 
		# cbb_AllZ
		# 
		self._cbb_AllZ.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._cbb_AllZ.Location = System.Drawing.Point(235, 13)
		self._cbb_AllZ.Name = "cbb_AllZ"
		self._cbb_AllZ.Size = System.Drawing.Size(103, 24)
		self._cbb_AllZ.TabIndex = 2
		self._cbb_AllZ.Text = "All/None"
		self._cbb_AllZ.UseVisualStyleBackColor = True
		# self._cbb_AllZ.Checked = True
		self._cbb_AllZ.CheckedChanged += self.Cbb_AllZCheckedChanged
		# 
		# total_XYValue
		# 
		self._total_XYValue.Location = System.Drawing.Point(179, 13)
		self._total_XYValue.Name = "total_XYValue"
		self._total_XYValue.Size = System.Drawing.Size(34, 22)
		self._total_XYValue.TabIndex = 3
		self._total_XYValue.TextChanged += self.Total_XYValueTextChanged
		# 
		# total_ZValue
		# 
		self._total_ZValue.Location = System.Drawing.Point(414, 13)
		self._total_ZValue.Name = "total_ZValue"
		self._total_ZValue.Size = System.Drawing.Size(34, 22)
		self._total_ZValue.TabIndex = 3
		self._total_ZValue.TextChanged += self.Total_ZValueTextChanged
		# 
		# label6
		# 
		self._label6.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._label6.Location = System.Drawing.Point(370, 16)
		self._label6.Name = "label6"
		self._label6.Size = System.Drawing.Size(38, 19)
		self._label6.TabIndex = 4
		self._label6.Text = "Total"
		# 
		# label7
		# 
		self._label7.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._label7.Location = System.Drawing.Point(135, 16)
		self._label7.Name = "label7"
		self._label7.Size = System.Drawing.Size(38, 19)
		self._label7.TabIndex = 4
		self._label7.Text = "Total"
		# 
		# btt_LoopZ
		# 
		# self._btt_LoopZ.BackColor = System.Drawing.SystemColors.ScrollBar
		self._btt_LoopZ.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_LoopZ.FlatAppearance.BorderColor = System.Drawing.Color.Red		
		self._btt_LoopZ.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Yellow
		self._btt_LoopZ.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Cyan		
		self._btt_LoopZ.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_LoopZ.ForeColor = System.Drawing.Color.Red
		self._btt_LoopZ.Location = System.Drawing.Point(383, 8)
		self._btt_LoopZ.Name = "btt_LoopZ"
		self._btt_LoopZ.Size = System.Drawing.Size(78, 30)
		self._btt_LoopZ.TabIndex = 11
		self._btt_LoopZ.Text = "Loop Z"
		self._btt_LoopZ.UseVisualStyleBackColor = False
		self._btt_LoopZ.Click += self.Btt_LoopZClick		
		# 
		# btt_ResetXYValue
		# 
		self._btt_ResetXYValue.AllowDrop = True
		self._btt_ResetXYValue.BackColor = System.Drawing.SystemColors.ScrollBar
		self._btt_ResetXYValue.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_ResetXYValue.Font = System.Drawing.Font("MS UI Gothic", 7.5, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_ResetXYValue.ForeColor = System.Drawing.Color.Red
		self._btt_ResetXYValue.Location = System.Drawing.Point(58, 155)
		self._btt_ResetXYValue.Name = "btt_ResetXYValue"
		self._btt_ResetXYValue.Size = System.Drawing.Size(89, 24)
		self._btt_ResetXYValue.TabIndex = 12
		self._btt_ResetXYValue.Text = "Reset XY"
		self._btt_ResetXYValue.UseVisualStyleBackColor = False
		self._btt_ResetXYValue.Click += self.Btt_ResetXYValueClick
		# 
		# btt_ResetZValue
		# 
		self._btt_ResetZValue.AllowDrop = True
		self._btt_ResetZValue.BackColor = System.Drawing.SystemColors.ScrollBar
		self._btt_ResetZValue.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_ResetZValue.Font = System.Drawing.Font("MS UI Gothic", 7.5, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_ResetZValue.ForeColor = System.Drawing.Color.Red
		self._btt_ResetZValue.Location = System.Drawing.Point(293, 155)
		self._btt_ResetZValue.Name = "btt_ResetZValue"
		self._btt_ResetZValue.Size = System.Drawing.Size(89, 24)
		self._btt_ResetZValue.TabIndex = 13
		self._btt_ResetZValue.Text = "Reset Z"
		self._btt_ResetZValue.UseVisualStyleBackColor = False
		self._btt_ResetZValue.Click += self.Btt_ResetZValueClick		
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(482, 501)
		self.Controls.Add(self._btt_LoopZ)
		self.Controls.Add(self._groupBox5)
		self.Controls.Add(self._label5)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_CREATE)
		self.Controls.Add(self._groupBox4)
		self.Controls.Add(self._groupBox3)
		self.Controls.Add(self._groupBox2)
		self.Controls.Add(self._btt_getZ)
		self.Controls.Add(self._btt_getXY)
		self.Controls.Add(self._groupBox1)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		self.MaximizeBox = False
		self.MinimizeBox = False
		self.Name = "MainForm"
		self.Text = "PipeCreate"
		self.TopMost = True
		self.TransparencyKey = System.Drawing.Color.FromArgb(255, 192, 192)		
		self._groupBox1.ResumeLayout(False)
		self._groupBox2.ResumeLayout(False)
		self._groupBox3.ResumeLayout(False)
		self._groupBox4.ResumeLayout(False)
		self._groupBox4.PerformLayout()
		self._groupBox5.ResumeLayout(False)
		self._groupBox5.PerformLayout()
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.Cursor = System.Windows.Forms.Cursors.Hand
		self.ResumeLayout(False)
	#____________________________________________________________________________________________#
	def Btt_getXYClick(self, sender, e):
		TransactionManager.Instance.EnsureInTransaction(doc)
		activeView = doc.ActiveView
		iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
		sketchPlane = SketchPlane.Create(doc, iRefPlane)
		doc.ActiveView.SketchPlane = sketchPlane
		condition = True
		pointsXY = []
		n = 0
		msg = "Pick Points on Current Section plane, hit ESC when finished."
		TaskDialog.Show("^------^", msg)
		while condition:
			try:
				pt=uidoc.Selection.PickPoint()
				n += 1
				pointsXY.append(pt)				
			except :
				condition = False
		doc.Delete(sketchPlane.Id)	
		for j in pointsXY:
			rpM = Autodesk.DesignScript.Geometry.Point.ByCoordinates(j.X*304.8, j.Y*304.8, j.Z*304.8)
			self._clb_XYValue.Items.Add(rpM)
		self._cbb_AllXY.Checked = True	
		TransactionManager.Instance.TransactionTaskDone()			
		pass
	"""_____________________________________________________________________________________"""	
	def Clb_XYValueSelectedIndexChanged(self, sender, e):		
		var = self._clb_XYValue.CheckedItems.Count
		n = 0
		if var != 0:
			for i in range(var):
				n += 1
				self._total_XYValue.Text = str(n)				
		else:
			self._total_XYValue.Text = str(0)	
		pass
	"""_____________________________________________________________________________________"""
	def Btt_getZClick(self, sender, e):
		TransactionManager.Instance.EnsureInTransaction(doc)
		activeView = doc.ActiveView
		iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
		sketchPlane = SketchPlane.Create(doc, iRefPlane)
		doc.ActiveView.SketchPlane = sketchPlane
		condition = True
		pointsZ = []
		n = 0
		msg = "Pick Points on Current Section plane, hit ESC when finished."
		TaskDialog.Show("^------^", msg)
		while condition:
			try:
				# logger('Line383:', n)
				pt=uidoc.Selection.PickPoint()
				n += 1
				pointsZ.append(pt)
			except :
				condition = False
		doc.Delete(sketchPlane.Id)	
		for j in pointsZ:
			rpM = Autodesk.DesignScript.Geometry.Point.ByCoordinates(j.X*304.8, j.Y*304.8, j.Z*304.8)
			self._clb_ZValue.Items.Add(rpM)
		self._cbb_AllZ.Checked = True	
		TransactionManager.Instance.TransactionTaskDone()			
		pass
	"""________________________________________________________"""
	def Clb_ZValueSelectedIndexChanged(self, sender, e):
		varZ = self._clb_ZValue.CheckedItems.Count
		n = 0
		if varZ != 0:
			for i in range(varZ):
				n += 1
				self._total_ZValue.Text = str(n)		
		else:
			self._total_ZValue.Text = str(0)				
		pass
	"""_________________________________________________________"""
	def Cbb_PipingSystemTypeSelectedIndexChanged(self, sender, e):
		pass
	"""_________________________________________________________"""
	def Cbb_PipeTypeSelectedIndexChanged(self, sender, e):
		pass
	"""_________________________________________________________"""
	def Cbb_RefLevelSelectedIndexChanged(self, sender, e):
		pass
	"""_________________________________________________________"""
	def Txb_DiameterTextChanged(self, sender, e):
		pass
	"""_________________________________________________________"""
	def Cb_AllXYCheckedChanged(self, sender, e):
		var = self._clb_XYValue.Items.Count
		rangers = range(var)
		for i in rangers:
			if self._cbb_AllXY.Checked == True:
				self._clb_XYValue.SetItemChecked(i, True)
				self._total_XYValue.Text = str(var)
			else:
				self._clb_XYValue.SetItemChecked(i, False)
				self._total_XYValue.Text = str(0)
		pass
	"""_________________________________________________________"""
	def Total_XYValueTextChanged(self, sender, e):
		pass
	"""_________________________________________________________"""
	def Cbb_AllZCheckedChanged(self, sender, e):
		var = self._clb_ZValue.Items.Count
		rangers = range(var)
		for i in rangers:
			if self._cbb_AllZ.Checked == True:
				self._clb_ZValue.SetItemChecked(i, True)
				self._total_ZValue.Text = str(var)
			else:
				self._clb_ZValue.SetItemChecked(i, False)
				self._total_ZValue.Text = str(0)		
		pass
	"""_________________________________________________________"""
	def Total_ZValueTextChanged(self, sender, e):
		pass	
	"""_________________________________________________________"""
	def Btt_CREATEClick(self, sender, e):
		pipingSystem = self._cbb_PipingSystemType.SelectedItem
		all_PipingSystem = getAllPipingSystems(doc)		
		all_PipingSystemsName = getAllPipingSystemsName(doc)
		sel_PipingSystemIdx = all_PipingSystemsName.index(pipingSystem)
		sel_pipingSystem = all_PipingSystem[sel_PipingSystemIdx]		
		"""_____________________________________________________"""
		pipeType = self._cbb_PipeType.SelectedItem
		all_PipeTypes = getAllPipeTypes(doc)
		all_PipeTypesName = getAllPipeTypesName(doc)
		sel_PipeTypeIdx = all_PipeTypesName.index(pipeType)
		sel_PipeType = all_PipeTypes[sel_PipeTypeIdx]
		"""_____________________________________________________"""
		refLevel = self._cbb_RefLevel.SelectedItem
		all_Levels = list(levelsCollector)
		sel_LevelIdx = all_Levels.index(refLevel)
		sel_Level = all_Levels[sel_LevelIdx]
		"""_____________________________________________________"""
		diameter1 = self._txb_Diameter.Text
		if diameter1.strip():
			try:
				diameter = int(diameter1)
			except vars.ValueError:
				diameter = None
		else:
			diameter = None
		#_________________________________________________________#
			point_XYValues = []
			point_ZValues = []
			desPointsList = []
			for pXY in self._clb_XYValue.CheckedItems:
				point_XYValues.append(pXY)
			for pZ in self._clb_ZValue.CheckedItems:
				point_ZValues.append(pZ)
			for pXY ,pZ  in zip(point_XYValues, point_ZValues):
				desPoint = Autodesk.DesignScript.Geometry.Point.ByCoordinates(pXY.X, pXY.Y, pZ.Z)
				desPointsList.append(desPoint)
			lst_Points1 = [i for i in desPointsList]
			lst_Points2 = [i for i in desPointsList[1:]]
			linesList = []
			for pt1, pt2 in zip(lst_Points1,lst_Points2):
				line =  Autodesk.DesignScript.Geometry.Line.ByStartPointEndPoint(pt1,pt2)
				linesList.append(line)
			firstPoint   = [x.StartPoint for x in linesList]
			secondPoint  = [x.EndPoint for x in linesList]
			pipesList = []
			pipesList1 = []
			TransactionManager.Instance.EnsureInTransaction(doc)
			for i,x in enumerate(firstPoint):
				try:
					levelId = sel_Level.Id
					sysTypeId = sel_pipingSystem.Id
					pipeTypeId = sel_PipeType.Id
					diam = diameter
					pipe = Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,sysTypeId,pipeTypeId,levelId,x.ToXyz(),secondPoint[i].ToXyz())
					param_Diameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
					param_Diameter.SetValueString(diam.ToString())
					param_Length = pipe.LookupParameter("CenterLength")
					pipeLength = pipe.LookupParameter("Length").AsDouble()
					TransactionManager.Instance.EnsureInTransaction(doc)
					pipesList.append(pipe.ToDSType(False))
					pipesList1.append(pipe)
					param_Length.Set()
					TransactionManager.Instance.TransactionTaskDone()
				except:
					pipesList.append(None)				
			TransactionManager.Instance.TransactionTaskDone()

		fittings = []
		connectors = {}
		connlist = []
		margin = 0				
		for ele in pipesList1:
			conns = ele.ConnectorManager.Connectors
			for conn in conns:
				if conn.IsConnected:
					continue
				connectors[conn] = None
				connlist.append(conn)				
		"""________________________________________________________________"""
		for k in connectors.keys():
			mindist = 1000000
			closest = None
			for conn in connlist:
				if conn.Owner.Id.Equals(k.Owner.Id):
					continue
				dist = k.Origin.DistanceTo(conn.Origin)
				if dist < mindist:
					mindist = dist
					closest = conn
			if mindist > margin:
				continue
			connectors[k] = closest
			connlist.remove(closest)
			try:
				del connectors[closest]
			except:
				pass			
		"""________________________________________________________________"""
		for k,v in connectors.items():
			TransactionManager.Instance.EnsureInTransaction(doc)		
			try:
				fitting = doc.Create.NewElbowFitting(k,v)
				fittings.append(fitting.ToDSType(False))
			except:
				pass
			TransactionManager.Instance.TransactionTaskDone()		
		"""________________________________________________________________"""
		if len(pipesList) != 0:
			msg = "Mission passed"
			TaskDialog.Show("^---Congrat---^", msg)	
		else: 	
			msg = "Mission Failed"
			TaskDialog.Show("^---Try again---^", msg)	
		TransactionManager.Instance.EnsureInTransaction(doc)		
		self._clb_XYValue.Items.Clear()		
		self._clb_ZValue.Items.Clear()		
		self._total_XYValue.Clear()	
		self._total_ZValue.Clear()
		TransactionManager.Instance.EnsureInTransaction(doc)	
		self._cbb_AllXY.Checked = False
		self._cbb_AllZ.Checked = False
		TransactionManager.Instance.TransactionTaskDone()	
		TransactionManager.Instance.TransactionTaskDone()	
		# self.Close()
		pass
	"""_____________________________________________________________________"""
	def Btt_LoopZClick(self, sender, e):
		TransactionManager.Instance.EnsureInTransaction(doc)
		activeView = doc.ActiveView
		iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
		sketchPlane = SketchPlane.Create(doc, iRefPlane)
		doc.ActiveView.SketchPlane = sketchPlane
		condition = True
		pointsZ = []
		n = 0
		msg = "Pick only 1 Points on Current Section plane, hit ESC when finished."
		TaskDialog.Show("^------^", msg)
		while condition:
			try:
				# logger('Line383:', n)
				pt=uidoc.Selection.PickPoint()
				pointsZ.append(pt)
			except :
				condition = False
		doc.Delete(sketchPlane.Id)	
		loop_n = self._clb_XYValue.Items.Count
		loopZ = []
		for j in pointsZ:
			rpM = Autodesk.DesignScript.Geometry.Point.ByCoordinates(j.X*304.8, j.Y*304.8, j.Z*304.8)
			rpM1 = [rpM]*loop_n
			for m in rpM1:
				self._clb_ZValue.Items.Add(m)
		self._cbb_AllZ.Checked = True		
		TransactionManager.Instance.TransactionTaskDone()			

		pass
	"""_____________________________________________________________________"""	
	def Btt_ResetXYValueClick(self, sender, e):
		self._clb_XYValue.Items.Clear()
		pass
	"""_____________________________________________________________________"""
	def Btt_ResetZValueClick(self, sender, e):
		self._clb_ZValue.Items.Clear()	
		pass	
	"""_____________________________________________________________________"""			
	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass
	"""_____________________________________________________________________"""	
f = MainForm()
Application.Run(f)