import clr
import sys 
import System   


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

"""____________________________________________________________________________________________"""
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

def getAllPipeTypes(doc):
	collector1 = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves)
	pipeTypes = collector1.ToElements()
	return pipeTypes
"""_____________________________________________________________________________________________"""

levelsCollector = FilteredElementCollector(doc).OfClass(Level).ToElements()
# for level in levelsCollector:
# 	levelName = level.Name
# 	levelsNameCollector.append(levelName)

def logger(title, content):
	import datetime
	date = datetime.datetime.now()
	f = open(r"C:\Users\95053\Desktop\Python\RevitAPI-master\RevitAPI\python.log", 'a')
	# C:\Users\95053\Desktop\Python\RevitAPI-master\RevitAPI\PipeCreate.py
	f.write(str(date) + '\n' + title + '\n' + str(content) + '\n')
	f.close()

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
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
		self._btt_OK = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._label5 = System.Windows.Forms.Label()
		self._groupBox5 = System.Windows.Forms.GroupBox()
		self._clb_XYValue = System.Windows.Forms.CheckedListBox()
		self._clb_ZValue = System.Windows.Forms.CheckedListBox()
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
		self._groupBox1.Location = System.Drawing.Point(12, 174)
		self._groupBox1.Name = "groupBox1"
		self._groupBox1.Size = System.Drawing.Size(449, 47)
		self._groupBox1.TabIndex = 0
		self._groupBox1.TabStop = False
		# 
		# btt_getXY
		# 
		self._btt_getXY.BackColor = System.Drawing.SystemColors.Info
		self._btt_getXY.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_getXY.ForeColor = System.Drawing.Color.Red
		self._btt_getXY.Location = System.Drawing.Point(13, 8)
		self._btt_getXY.Name = "btt_getXY"
		self._btt_getXY.Size = System.Drawing.Size(140, 30)
		self._btt_getXY.TabIndex = 1
		self._btt_getXY.Text = "get X,Y value"
		self._btt_getXY.UseVisualStyleBackColor = False
		self._btt_getXY.Click += self.Btt_getXYClick
		# 
		# btt_getZ
		# 
		self._btt_getZ.BackColor = System.Drawing.SystemColors.Info
		self._btt_getZ.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_getZ.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_getZ.ForeColor = System.Drawing.Color.Red
		self._btt_getZ.Location = System.Drawing.Point(295, 8)
		self._btt_getZ.Name = "btt_getZ"
		self._btt_getZ.Size = System.Drawing.Size(140, 30)
		self._btt_getZ.TabIndex = 2
		self._btt_getZ.Text = "get Z value"
		self._btt_getZ.UseVisualStyleBackColor = False
		self._btt_getZ.Click += self.Btt_getZClick
		# 
		# cbb_PipingSystemType
		# 
		# self._cbb_PipingSystemType.DisplayMember = "Name"
		self._cbb_PipingSystemType.AllowDrop = True
		self._cbb_PipingSystemType.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._cbb_PipingSystemType.FormattingEnabled = True
		self._cbb_PipingSystemType.Location = System.Drawing.Point(180, 12)
		self._cbb_PipingSystemType.Name = "cbb_PipingSystemType"
		self._cbb_PipingSystemType.Size = System.Drawing.Size(263, 23)
		self._cbb_PipingSystemType.TabIndex = 3
		self._cbb_PipingSystemType.SelectedIndexChanged += self.Cbb_PipingSystemTypeSelectedIndexChanged
		self._cbb_PipingSystemType.Items.AddRange(System.Array[System.Object](pipingSystemsCollector)) 
		# 
		# label1
		# 
		self._label1.BackColor = System.Drawing.SystemColors.Info
		self._label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._label1.Font = System.Drawing.Font("MS UI Gothic", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._label1.ForeColor = System.Drawing.Color.Blue
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
		self._groupBox2.Location = System.Drawing.Point(12, 227)
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
		self._label2.Location = System.Drawing.Point(6, 12)
		self._label2.Name = "label2"
		self._label2.Size = System.Drawing.Size(142, 23)
		self._label2.TabIndex = 4
		self._label2.Text = "PipeType"
		# 
		# cbb_PipeType
		# 
		# self._cbb_PipeType.DisplayMember = "Name"
		self._cbb_PipeType.AllowDrop = True
		self._cbb_PipeType.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._cbb_PipeType.FormattingEnabled = True
		self._cbb_PipeType.Location = System.Drawing.Point(180, 12)
		self._cbb_PipeType.Name = "cbb_PipeType"
		self._cbb_PipeType.Size = System.Drawing.Size(263, 23)
		self._cbb_PipeType.TabIndex = 3
		self._cbb_PipeType.SelectedIndexChanged += self.Cbb_PipeTypeSelectedIndexChanged
		self._cbb_PipeType.Items.AddRange(System.Array[System.Object](pipeTypesCollector))
		# 
		# groupBox3
		# 
		self._groupBox3.Controls.Add(self._label3)
		self._groupBox3.Controls.Add(self._cbb_RefLevel)
		self._groupBox3.Location = System.Drawing.Point(12, 280)
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
		# 
		# groupBox4
		# 
		self._groupBox4.Controls.Add(self._txb_Diameter)
		self._groupBox4.Controls.Add(self._label4)
		self._groupBox4.Location = System.Drawing.Point(12, 333)
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
		# btt_OK
		# 
		self._btt_OK.BackColor = System.Drawing.SystemColors.ScrollBar
		self._btt_OK.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._btt_OK.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_OK.ForeColor = System.Drawing.Color.Red
		self._btt_OK.Location = System.Drawing.Point(317, 410)
		self._btt_OK.Name = "btt_OK"
		self._btt_OK.Size = System.Drawing.Size(49, 30)
		self._btt_OK.TabIndex = 8
		self._btt_OK.Text = "OK"
		self._btt_OK.UseVisualStyleBackColor = False
		self._btt_OK.Click += self.Btt_OKClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.BackColor = System.Drawing.SystemColors.ScrollBar
		self._btt_CANCLE.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._btt_CANCLE.Font = System.Drawing.Font("MS UI Gothic", 10, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128, True)
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
		self._btt_CANCLE.Location = System.Drawing.Point(372, 410)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(89, 30)
		self._btt_CANCLE.TabIndex = 8
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = False
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# label5
		# 
		self._label5.Location = System.Drawing.Point(1, 434)
		self._label5.Name = "label5"
		self._label5.Size = System.Drawing.Size(53, 19)
		self._label5.TabIndex = 9
		self._label5.Text = "@FVC"
		# 
		# groupBox5
		# 
		self._groupBox5.Controls.Add(self._clb_ZValue)
		self._groupBox5.Controls.Add(self._clb_XYValue)
		self._groupBox5.Location = System.Drawing.Point(13, 44)
		self._groupBox5.Name = "groupBox5"
		self._groupBox5.Size = System.Drawing.Size(448, 124)
		self._groupBox5.TabIndex = 10
		self._groupBox5.TabStop = False
		# 
		# clb_XYValue
		# 
		self._clb_XYValue.AllowDrop = True
		self._clb_XYValue.CheckOnClick = True
		self._clb_XYValue.FormattingEnabled = True
		self._clb_XYValue.HorizontalScrollbar = True	
		self._clb_XYValue.Location = System.Drawing.Point(0, 21)
		self._clb_XYValue.Name = "clb_XYValue"
		self._clb_XYValue.Size = System.Drawing.Size(264, 89)
		self._clb_XYValue.TabIndex = 0
		self._clb_XYValue.SelectedIndexChanged += self.Clb_XYValueSelectedIndexChanged	
		# 
		# clb_ZValue
		# 
		self._clb_ZValue.AllowDrop = True
		self._clb_ZValue.CheckOnClick = True
		self._clb_ZValue.FormattingEnabled = True
		self._clb_ZValue.HorizontalScrollbar = True			
		self._clb_ZValue.Location = System.Drawing.Point(282, 21)
		self._clb_ZValue.Name = "clb_ZValue"
		self._clb_ZValue.Size = System.Drawing.Size(166, 89)
		self._clb_ZValue.TabIndex = 1
		self._clb_ZValue.SelectedIndexChanged += self.Clb_ZValueSelectedIndexChanged
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(482, 453)
		self.Controls.Add(self._groupBox5)
		self.Controls.Add(self._label5)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_OK)
		self.Controls.Add(self._groupBox4)
		self.Controls.Add(self._groupBox3)
		self.Controls.Add(self._groupBox2)
		self.Controls.Add(self._btt_getZ)
		self.Controls.Add(self._btt_getXY)
		self.Controls.Add(self._groupBox1)
		self.Name = "MainForm"
		self.Text = "PipeCreate"
		self._groupBox1.ResumeLayout(False)
		self._groupBox2.ResumeLayout(False)
		self._groupBox3.ResumeLayout(False)
		self._groupBox4.ResumeLayout(False)
		self._groupBox4.PerformLayout()
		self._groupBox5.ResumeLayout(False)
		self.ResumeLayout(False)

	def Btt_getXYClick(self, sender, e):
		# TransactionManager.Instance.EnsureInTransaction(doc)
		# condition = True
		# pointsXY = []
		# n = 0

		# msg = "Pick Points on Current Floor plane, hit ESC when finished."
		# TaskDialog.Show("^---Ai An Banh Mi Khong??---^", msg)

		# while condition:
		# 	try:
		# 		logger('Line383:', n)
		# 		pt=uidoc.Selection.PickPoint()
		# 		n += 1
		# 		pointsXY.append(pt)
		# 	except :
		# 		condition = False
		# for i in pointsXY:
		# 	rpM = Autodesk.DesignScript.Geometry.Point.ByCoordinates(i.X*304.8,i.Y*304.8,i.Z*304.8)
		# 	self._clb_XYValue.Items.Add(rpM)
		# TransactionManager.Instance.TransactionTaskDone()		

		# TransactionManager.Instance.EnsureInTransaction(doc)
		TransactionManager.Instance.EnsureInTransaction(doc)
		activeView = doc.ActiveView
		iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
		sketchPlane = SketchPlane.Create(doc, iRefPlane)
		doc.ActiveView.SketchPlane = sketchPlane
		condition = True
		pointsXY = []
		n = 0

		msg = "Pick Points on Current Section plane, hit ESC when finished."
		TaskDialog.Show("^---Ai An Banh Mi Khong??---^", msg)

		while condition:
			try:
				# logger('Line383:', n)
				
				pt=uidoc.Selection.PickPoint()
				n += 1
				pointsXY.append(pt)
		
			except :
				condition = False
		doc.Delete(sketchPlane.Id)	
		for j in pointsXY:
			rpM = Autodesk.DesignScript.Geometry.Point.ByCoordinates(j.X, j.Y, j.Z)
			self._clb_XYValue.Items.Add(rpM)
		TransactionManager.Instance.TransactionTaskDone()			
		pass
		
	def Clb_XYValueSelectedIndexChanged(self, sender, e):

		pass

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
		TaskDialog.Show("^---Ai An Banh Mi Khong??---^", msg)

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
			rpM = Autodesk.DesignScript.Geometry.Point.ByCoordinates(j.X, j.Y, j.Z)
			self._clb_ZValue.Items.Add(rpM)
		TransactionManager.Instance.TransactionTaskDone()			
		pass

	def Clb_ZValueSelectedIndexChanged(self, sender, e):
		pass

	def Cbb_PipingSystemTypeSelectedIndexChanged(self, sender, e):
		# self._cbb_PipingSystemType.selectedIndex = 0	
		pass

	def Cbb_PipeTypeSelectedIndexChanged(self, sender, e):
		pass

	def Cbb_RefLevelSelectedIndexChanged(self, sender, e):
		pass

	def Txb_DiameterTextChanged(self, sender, e):
		pass
	
	def Btt_OKClick(self, sender, e):
		pipingSystem = []
		for i in self._cbb_PipingSystemType.SelectedItem:
			pipingSystem.append(i)
		# stl = len(pipingSystem)
		all_PipingSystem = getAllPipingSystems(doc)
		# all_PipingSystemsName = getAllPipingSystemsName(doc)
		# sel_PipingSystemIdx = all_PipingSystemsName.index(pipingSystem)
		# sel_pipingSystem = all_PipingSystem[sel_PipingSystemIdx]
		auto_PipingSystem = all_PipingSystem[0]

		"""______________________________________________________"""
		pipeType = []
		for j in self._cbb_PipeType.SelectedItem:
			pipeType.append(j)
		# ptl = len(pipeType)
		all_PipeTypes = getAllPipeTypes(doc)
		# all_PipeTypesName = getAllPipeTypesName(doc)
		# sel_PipeTypeIdx = all_PipeTypesName.index(pipeType)
		# sel_PipeType = all_PipeTypes[sel_PipeTypeIdx]
		auto_PipeType = all_PipeTypes[0]
		"""______________________________________________________"""
		# refLevel = []
		# for k in self._cbb_RefLevel.SelectedItem:
		# 	refLevel.append(k)
		refLevel = levelsCollector[0]
		"""_______________________________________________________"""
		diameter = []
		for m in self._txb_Diameter.Text:
			diameter.append(int(m))
		"""_______________________________________________________"""
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
		elements = []

		TransactionManager.Instance.EnsureInTransaction(doc)
		for i,x in enumerate(firstPoint):
			try:
				levelId = refLevel.Id
				sysTypeId = auto_PipingSystem.Id
				pipeTypeId = auto_PipeType.Id
				diam = diameter[0]
				diam1 = diam/304.8
				pipe = Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,sysTypeId,pipeTypeId,levelId,x.ToXyz(),secondPoint[i].ToXyz())
				
				param = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
				param.SetValueString(diam1.ToString())
			
				elements.append(pipe.ToDSType(False))	
			except:
				elements.append(None)

		TransactionManager.Instance.TransactionTaskDone()
		self.Close()
		pass

	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass

f = MainForm()
Application.Run(f)