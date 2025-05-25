"""Copyright by: vudinhduybm@gmail.com"""
import clr
import sys 
import System   
import math
import os
clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *
clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB import SaveAsOptions
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB.Plumbing import*
clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI import*
from Autodesk.Revit.UI.Selection import ISelectionFilter
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
from System.Windows.Forms import OpenFileDialog
'''___'''
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
'''___'''
#region _def
def getAllfloors(doc):
    collector = FilteredElementCollector(doc).OfClass(ViewPlan)
    floors = [view for view in collector if not view.IsTemplate]
    return floors
def pickPoints(doc):
	TransactionManager.Instance.EnsureInTransaction(doc)
	activeView = doc.ActiveView
	iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
	sketchPlane = SketchPlane.Create(doc, iRefPlane)
	doc.ActiveView.SketchPlane = sketchPlane
	condition = True
	pointsList = []
	dynPList = []
	n = 0
	msg = "Pick Point, hit ESC when finished."
	TaskDialog.Show("^------^", msg)
	while condition:
		try:
			pt=uidoc.Selection.PickPoint()
			n += 1
			pointsList.append(pt)				
		except :
			condition = False
	doc.Delete(sketchPlane.Id)	
	for j in pointsList:
		dynP = j.ToPoint()
		dynPList.append(dynP)
	TransactionManager.Instance.TransactionTaskDone()			
	return pointsList[0]
#endregion
#region _MainForm
allfloorsView = getAllfloors(doc)

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()	
	def InitializeComponent(self):
		self._grb_LinkCadPath = System.Windows.Forms.GroupBox()
		self._btt_LinkCad = System.Windows.Forms.Button()
		self._clb_LinkCad = System.Windows.Forms.CheckedListBox()
		self._ckb_AllNoneCad = System.Windows.Forms.CheckBox()
		self._txb_TotalCad = System.Windows.Forms.TextBox()
		self._lbl_TotalCad = System.Windows.Forms.Label()
		self._btt_Reset = System.Windows.Forms.Button()
		self._grb_FloorView = System.Windows.Forms.GroupBox()
		self._clb_FloorView = System.Windows.Forms.CheckedListBox()
		self._ckb_AllNonefloorView = System.Windows.Forms.CheckBox()
		self._lbl_TotalfloorView = System.Windows.Forms.Label()
		self._txb_TotalView = System.Windows.Forms.TextBox()
		self._btt_Run = System.Windows.Forms.Button()
		self._btt_Cancle = System.Windows.Forms.Button()
		self._lbl_vitaminD = System.Windows.Forms.Label()
		self._grb_floorPoint = System.Windows.Forms.GroupBox()
		self._btt_ResetPickPoints = System.Windows.Forms.Button()
		self._lbb_TotalPoints = System.Windows.Forms.Label()
		self._txb_ToTalPoints = System.Windows.Forms.TextBox()
		self._ckb_AllNonePoints = System.Windows.Forms.CheckBox()
		self._clb_FloorPoints = System.Windows.Forms.CheckedListBox()
		self._btt_PickPoints = System.Windows.Forms.Button()
		self._grb_LinkCadPath.SuspendLayout()
		self._grb_FloorView.SuspendLayout()
		self._grb_floorPoint.SuspendLayout()
		self.SuspendLayout()
		self.filePathMap = {}
		# 
		# grb_LinkCadPath
		# 
		self._grb_LinkCadPath.Controls.Add(self._btt_Reset)
		self._grb_LinkCadPath.Controls.Add(self._lbl_TotalCad)
		self._grb_LinkCadPath.Controls.Add(self._txb_TotalCad)
		self._grb_LinkCadPath.Controls.Add(self._ckb_AllNoneCad)
		self._grb_LinkCadPath.Controls.Add(self._clb_LinkCad)
		self._grb_LinkCadPath.Controls.Add(self._btt_LinkCad)
		self._grb_LinkCadPath.Location = System.Drawing.Point(636, 12)
		self._grb_LinkCadPath.Name = "grb_LinkCadPath"
		self._grb_LinkCadPath.Size = System.Drawing.Size(318, 270)
		self._grb_LinkCadPath.TabIndex = 0
		self._grb_LinkCadPath.TabStop = False
		# 
		# btt_LinkCad
		# 
		self._btt_LinkCad.AutoSize = True
		self._btt_LinkCad.BackColor = System.Drawing.SystemColors.ButtonHighlight
		self._btt_LinkCad.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_LinkCad.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_LinkCad.ForeColor = System.Drawing.Color.Red
		self._btt_LinkCad.Location = System.Drawing.Point(6, 16)
		self._btt_LinkCad.Name = "btt_LinkCad"
		self._btt_LinkCad.Size = System.Drawing.Size(138, 29)
		self._btt_LinkCad.TabIndex = 0
		self._btt_LinkCad.Text = "Select Link Cad"
		self._btt_LinkCad.UseVisualStyleBackColor = False
		self._btt_LinkCad.Click += self.Btt_LinkCadClick
		# 
		# clb_LinkCad
		# 
		self._clb_LinkCad.DisplayMember = 'Name'
		self._clb_LinkCad.AllowDrop = True
		self._clb_LinkCad.BackColor = System.Drawing.SystemColors.MenuBar
		self._clb_LinkCad.CheckOnClick = True
		self._clb_LinkCad.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_LinkCad.FormattingEnabled = True
		self._clb_LinkCad.HorizontalScrollbar = True
		self._clb_LinkCad.Location = System.Drawing.Point(6, 86)
		self._clb_LinkCad.Name = "clb_LinkCad"
		self._clb_LinkCad.Size = System.Drawing.Size(306, 169)
		self._clb_LinkCad.TabIndex = 1
		self._clb_LinkCad.SelectedIndexChanged += self.Clb_LinkCadSelectedIndexChanged
		# 
		# ckb_AllNoneCad
		# 
		self._ckb_AllNoneCad.AutoSize = True
		self._ckb_AllNoneCad.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._ckb_AllNoneCad.Location = System.Drawing.Point(10, 58)
		self._ckb_AllNoneCad.Name = "ckb_AllNoneCad"
		self._ckb_AllNoneCad.Size = System.Drawing.Size(89, 22)
		self._ckb_AllNoneCad.TabIndex = 2
		self._ckb_AllNoneCad.Text = "All/None"
		self._ckb_AllNoneCad.UseVisualStyleBackColor = True
		self._ckb_AllNoneCad.CheckedChanged += self.Ckb_AllNoneCadCheckedChanged
		# 
		# txb_TotalCad
		# 
		self._txb_TotalCad.BackColor = System.Drawing.SystemColors.Menu
		self._txb_TotalCad.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_TotalCad.Location = System.Drawing.Point(256, 56)
		self._txb_TotalCad.Name = "txb_TotalCad"
		self._txb_TotalCad.Size = System.Drawing.Size(56, 24)
		self._txb_TotalCad.TabIndex = 3
		self._txb_TotalCad.TextChanged += self.Txb_TotalCadTextChanged
		# 
		# lbl_TotalCad
		# 
		self._lbl_TotalCad.AutoSize = True
		self._lbl_TotalCad.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lbl_TotalCad.Location = System.Drawing.Point(178, 58)
		self._lbl_TotalCad.Name = "lbl_TotalCad"
		self._lbl_TotalCad.Size = System.Drawing.Size(78, 19)
		self._lbl_TotalCad.TabIndex = 4
		self._lbl_TotalCad.Text = "Total Cad"
		# 
		# btt_Reset
		# 
		self._btt_Reset.AutoSize = True
		self._btt_Reset.BackColor = System.Drawing.SystemColors.ButtonHighlight
		self._btt_Reset.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_Reset.ForeColor = System.Drawing.Color.Red
		self._btt_Reset.Location = System.Drawing.Point(225, 16)
		self._btt_Reset.Name = "btt_Reset"
		self._btt_Reset.Size = System.Drawing.Size(93, 29)
		self._btt_Reset.TabIndex = 5
		self._btt_Reset.Text = "Reset"
		self._btt_Reset.UseVisualStyleBackColor = False
		self._btt_Reset.Click += self.Btt_ResetClick
		# 
		# grb_floorView
		# 
		self._grb_FloorView.Controls.Add(self._txb_TotalView)
		self._grb_FloorView.Controls.Add(self._lbl_TotalfloorView)
		self._grb_FloorView.Controls.Add(self._ckb_AllNonefloorView)
		self._grb_FloorView.Controls.Add(self._clb_FloorView)
		self._grb_FloorView.Cursor = System.Windows.Forms.Cursors.Default
		self._grb_FloorView.Location = System.Drawing.Point(9, 12)
		self._grb_FloorView.Name = "grb_floorView"
		self._grb_FloorView.Size = System.Drawing.Size(275, 270)
		self._grb_FloorView.TabIndex = 1
		self._grb_FloorView.TabStop = False
		# 
		# clb_floorView
		# 
		self._clb_FloorView.DisplayMember = 'Name'
		self._clb_FloorView.AllowDrop = True
		self._clb_FloorView.BackColor = System.Drawing.SystemColors.MenuBar
		self._clb_FloorView.CheckOnClick = True
		self._clb_FloorView.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_FloorView.FormattingEnabled = True
		self._clb_FloorView.HorizontalScrollbar = True
		self._clb_FloorView.Location = System.Drawing.Point(6, 86)
		self._clb_FloorView.Name = "clb_floorView"
		self._clb_FloorView.Size = System.Drawing.Size(257, 169)
		self._clb_FloorView.Sorted = True
		self._clb_FloorView.TabIndex = 6
		self._clb_FloorView.SelectedIndexChanged += self.Clb_floorViewSelectedIndexChanged
		self._clb_FloorView.Items.AddRange(System.Array[System.Object](allfloorsView))		
		self._clb_FloorView.SelectedIndex = 0		
		# 
		# ckb_AllNonefloorView
		# 
		self._ckb_AllNonefloorView.AutoSize = True
		self._ckb_AllNonefloorView.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._ckb_AllNonefloorView.Location = System.Drawing.Point(11, 58)
		self._ckb_AllNonefloorView.Name = "ckb_AllNonefloorView"
		self._ckb_AllNonefloorView.Size = System.Drawing.Size(89, 22)
		self._ckb_AllNonefloorView.TabIndex = 6
		self._ckb_AllNonefloorView.Text = "All/None"
		self._ckb_AllNonefloorView.UseVisualStyleBackColor = True
		self._ckb_AllNonefloorView.CheckedChanged += self.Ckb_AllNonefloorViewCheckedChanged
		# 
		# lbl_TotalfloorView
		# 
		self._lbl_TotalfloorView.AutoSize = True
		self._lbl_TotalfloorView.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lbl_TotalfloorView.Location = System.Drawing.Point(123, 58)
		self._lbl_TotalfloorView.Name = "lbl_TotalfloorView"
		self._lbl_TotalfloorView.Size = System.Drawing.Size(85, 19)
		self._lbl_TotalfloorView.TabIndex = 6
		self._lbl_TotalfloorView.Text = "Total View"
		# 
		# txb_TotalView
		# 
		self._txb_TotalView.BackColor = System.Drawing.SystemColors.Menu
		self._txb_TotalView.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_TotalView.Location = System.Drawing.Point(207, 56)
		self._txb_TotalView.Name = "txb_TotalView"
		self._txb_TotalView.Size = System.Drawing.Size(56, 24)
		self._txb_TotalView.TabIndex = 6
		self._txb_TotalView.TextChanged += self.Txb_TotalViewTextChanged
		# 
		# btt_Run
		# 
		self._btt_Run.AutoSize = True
		self._btt_Run.BackColor = System.Drawing.SystemColors.ButtonHighlight
		self._btt_Run.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_Run.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_Run.ForeColor = System.Drawing.Color.Red
		self._btt_Run.Location = System.Drawing.Point(778, 297)
		self._btt_Run.Name = "btt_Run"
		self._btt_Run.Size = System.Drawing.Size(85, 29)
		self._btt_Run.TabIndex = 6
		self._btt_Run.Text = "RUN"
		self._btt_Run.UseVisualStyleBackColor = False
		self._btt_Run.Click += self.Btt_RunClick
		# 
		# btt_Cancle
		# 
		self._btt_Cancle.AutoSize = True
		self._btt_Cancle.BackColor = System.Drawing.SystemColors.ButtonHighlight
		self._btt_Cancle.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_Cancle.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_Cancle.ForeColor = System.Drawing.Color.Red
		self._btt_Cancle.Location = System.Drawing.Point(869, 297)
		self._btt_Cancle.Name = "btt_Cancle"
		self._btt_Cancle.Size = System.Drawing.Size(85, 29)
		self._btt_Cancle.TabIndex = 7
		self._btt_Cancle.Text = "CANCLE"
		self._btt_Cancle.UseVisualStyleBackColor = False
		self._btt_Cancle.Click += self.Btt_CancleClick
		# 
		# lbl_vitaminD
		# 
		self._lbl_vitaminD.AutoSize = True
		self._lbl_vitaminD.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lbl_vitaminD.Location = System.Drawing.Point(15, 315)
		self._lbl_vitaminD.Name = "lbl_vitaminD"
		self._lbl_vitaminD.Size = System.Drawing.Size(25, 13)
		self._lbl_vitaminD.TabIndex = 7
		self._lbl_vitaminD.Text = "@D"
		# 
		# grb_floorPoint
		# 
		self._grb_floorPoint.Controls.Add(self._btt_ResetPickPoints)
		self._grb_floorPoint.Controls.Add(self._lbb_TotalPoints)
		self._grb_floorPoint.Controls.Add(self._txb_ToTalPoints)
		self._grb_floorPoint.Controls.Add(self._ckb_AllNonePoints)
		self._grb_floorPoint.Controls.Add(self._clb_FloorPoints)
		self._grb_floorPoint.Controls.Add(self._btt_PickPoints)
		self._grb_floorPoint.Location = System.Drawing.Point(301, 12)
		self._grb_floorPoint.Name = "grb_floorPoint"
		self._grb_floorPoint.Size = System.Drawing.Size(318, 270)
		self._grb_floorPoint.TabIndex = 6
		self._grb_floorPoint.TabStop = False
		# 
		# btt_ResetPickPoints
		# 
		self._btt_ResetPickPoints.AutoSize = True
		self._btt_ResetPickPoints.BackColor = System.Drawing.SystemColors.ButtonHighlight
		self._btt_ResetPickPoints.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_ResetPickPoints.ForeColor = System.Drawing.Color.Red
		self._btt_ResetPickPoints.Location = System.Drawing.Point(225, 16)
		self._btt_ResetPickPoints.Name = "btt_ResetPickPoints"
		self._btt_ResetPickPoints.Size = System.Drawing.Size(93, 29)
		self._btt_ResetPickPoints.TabIndex = 5
		self._btt_ResetPickPoints.Text = "Reset"
		self._btt_ResetPickPoints.UseVisualStyleBackColor = False
		self._btt_ResetPickPoints.Click += self.Btt_ResetPickPointsClick
		# 
		# lbb_TotalPoints
		# 
		self._lbb_TotalPoints.AutoSize = True
		self._lbb_TotalPoints.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lbb_TotalPoints.Location = System.Drawing.Point(162, 58)
		self._lbb_TotalPoints.Name = "lbb_TotalPoints"
		self._lbb_TotalPoints.Size = System.Drawing.Size(94, 19)
		self._lbb_TotalPoints.TabIndex = 4
		self._lbb_TotalPoints.Text = "Total Points"
		# 
		# txb_ToTalPoints
		# 
		self._txb_ToTalPoints.BackColor = System.Drawing.SystemColors.Menu
		self._txb_ToTalPoints.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_ToTalPoints.Location = System.Drawing.Point(256, 56)
		self._txb_ToTalPoints.Name = "txb_ToTalPoints"
		self._txb_ToTalPoints.Size = System.Drawing.Size(56, 24)
		self._txb_ToTalPoints.TabIndex = 3
		self._txb_ToTalPoints.TextChanged += self.Txb_ToTalPointsTextChanged
		# 
		# ckb_AllNonePoints
		# 
		self._ckb_AllNonePoints.AutoSize = True
		self._ckb_AllNonePoints.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._ckb_AllNonePoints.Location = System.Drawing.Point(10, 58)
		self._ckb_AllNonePoints.Name = "ckb_AllNonePoints"
		self._ckb_AllNonePoints.Size = System.Drawing.Size(89, 22)
		self._ckb_AllNonePoints.TabIndex = 2
		self._ckb_AllNonePoints.Text = "All/None"
		self._ckb_AllNonePoints.UseVisualStyleBackColor = True
		self._ckb_AllNonePoints.CheckedChanged += self.Ckb_AllNonePointsCheckedChanged
		# 
		# clb_floorPoints
		# 
		self._clb_FloorPoints.AllowDrop = True
		self._clb_FloorPoints.BackColor = System.Drawing.SystemColors.MenuBar
		self._clb_FloorPoints.CheckOnClick = True
		self._clb_FloorPoints.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_FloorPoints.FormattingEnabled = True
		self._clb_FloorPoints.HorizontalScrollbar = True
		self._clb_FloorPoints.Location = System.Drawing.Point(6, 86)
		self._clb_FloorPoints.Name = "clb_floorPoints"
		self._clb_FloorPoints.Size = System.Drawing.Size(306, 169)
		self._clb_FloorPoints.TabIndex = 1
		self._clb_FloorPoints.SelectedIndexChanged += self.Clb_floorPointsSelectedIndexChanged
		# 
		# btt_PickPoints
		# 
		self._btt_PickPoints.AutoSize = True
		self._btt_PickPoints.BackColor = System.Drawing.SystemColors.ButtonHighlight
		self._btt_PickPoints.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_PickPoints.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_PickPoints.ForeColor = System.Drawing.Color.Red
		self._btt_PickPoints.Location = System.Drawing.Point(6, 16)
		self._btt_PickPoints.Name = "btt_PickPoints"
		self._btt_PickPoints.Size = System.Drawing.Size(138, 29)
		self._btt_PickPoints.TabIndex = 0
		self._btt_PickPoints.Text = "Pick Points"
		self._btt_PickPoints.UseVisualStyleBackColor = False
		self._btt_PickPoints.Click += self.Btt_PickPointsClick
		# 
		# MainForm
		# 
		self.BackColor = System.Drawing.SystemColors.Menu
		self.ClientSize = System.Drawing.Size(966, 338)
		self.Controls.Add(self._grb_floorPoint)
		self.Controls.Add(self._lbl_vitaminD)
		self.Controls.Add(self._btt_Cancle)
		self.Controls.Add(self._btt_Run)
		self.Controls.Add(self._grb_FloorView)
		self.Controls.Add(self._grb_LinkCadPath)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		# self.Icon = resources.GetObject("$this.Icon")
		self.Name = "MainForm"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.Text = "Link Cad To floor View_v2"
		self.TopMost = True
		self._grb_LinkCadPath.ResumeLayout(False)
		self._grb_LinkCadPath.PerformLayout()
		self._grb_FloorView.ResumeLayout(False)
		self._grb_FloorView.PerformLayout()
		self._grb_floorPoint.ResumeLayout(False)
		self._grb_floorPoint.PerformLayout()
		self.ResumeLayout(False)
		self.PerformLayout()
	def Clb_floorViewSelectedIndexChanged(self, sender, e):
		varZ = self._clb_FloorView.CheckedItems.Count
		n = 0
		if varZ != 0:
			for i in range(varZ):
				n += 1
				self._txb_TotalView.Text = str(n)		
		else:
			self._txb_TotalView.Text = str(0)						
		pass		
	def Ckb_AllNonefloorViewCheckedChanged(self, sender, e):
		var = self._clb_FloorView.Items.Count
		rangers = range(var)
		for i in rangers:
			if self._ckb_AllNonefloorView.Checked == True:
				self._clb_FloorView.SetItemChecked(i, True)
				self._txb_TotalView.Text = str(var)
			else:
				self._clb_FloorView.SetItemChecked(i, False)
				self._txb_TotalView.Text = str(0)
		pass			
	def Txb_TotalViewTextChanged(self, sender, e):
		pass	
	def Btt_LinkCadClick(self, sender, e):	
		def __init__(self):
			# Dictionary to store file name and corresponding file path
			self.filePathMap = {}	
		openDialog = OpenFileDialog()
		openDialog.Multiselect = False
		openDialog.Filter = "DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
		if openDialog.ShowDialog() == DialogResult.OK:  # Check if the user confirmed the selection
			filePaths = openDialog.FileNames  # This will be an array of selected file paths
			for filePath in filePaths:  # Process each file path
				fileName = os.path.basename(filePath)  # Extract the file name
				self._clb_LinkCad.Items.Add(fileName)  
				# self.linkFilePath = filePath
				self.filePathMap[fileName] = filePath
			self._ckb_AllNoneCad.Checked = True	
		pass
	def Btt_ResetClick(self, sender, e):
		self._clb_LinkCad.Items.Clear()
		pass
	def Clb_LinkCadSelectedIndexChanged(self, sender, e):
		varZ = self._clb_LinkCad.CheckedItems.Count
		n = 0
		if varZ != 0:
			for i in range(varZ):
				n += 1
				self._txb_TotalCad.Text = str(n)		
		else:
			self._txb_TotalCad.Text = str(0)						
		pass
	def Txb_TotalCadTextChanged(self, sender, e):
		pass
	def Ckb_AllNoneCadCheckedChanged(self, sender, e):
		var = self._clb_LinkCad.Items.Count
		rangers = range(var)
		for i in rangers:
			if self._ckb_AllNoneCad.Checked == True:
				self._clb_LinkCad.SetItemChecked(i, True)
				self._txb_TotalCad.Text = str(var)
			else:
				self._clb_LinkCad.SetItemChecked(i, False)
				self._txb_TotalCad.Text = str(0)
		pass		
	def Btt_PickPointsClick(self, sender, e):
		TransactionManager.Instance.EnsureInTransaction(doc)
		activeView = doc.ActiveView
		iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
		sketchPlane = SketchPlane.Create(doc, iRefPlane)
		doc.ActiveView.SketchPlane = sketchPlane
		condition = True
		floorPointsList = []
		n = 0
		msg = "Pick Points on Current floor plane, hit ESC when finished."
		TaskDialog.Show("^------^", msg)
		while condition:
			try:
				revitPoint = uidoc.Selection.PickPoint()
				n += 1
				floorPointsList.append(revitPoint)				
			except :
				condition = False
		doc.Delete(sketchPlane.Id)	
		for p in floorPointsList:
			dynPoint = p.ToPoint()
			self._clb_FloorPoints.Items.Add(dynPoint)
		self._ckb_AllNonePoints.Checked = True	
		TransactionManager.Instance.TransactionTaskDone()	
		pass
	def Btt_ResetPickPointsClick(self, sender, e):
		self._clb_FloorPoints.Items.Clear()
		pass
	def Ckb_AllNonePointsCheckedChanged(self, sender, e):
		var = self._clb_FloorPoints.Items.Count
		rangers = range(var)
		for i in rangers:
			if self._ckb_AllNonePoints.Checked == True:
				self._clb_FloorPoints.SetItemChecked(i, True)
				self._txb_ToTalPoints.Text = str(var)
			else:
				self._clb_FloorPoints.SetItemChecked(i, False)
				self._txb_ToTalPoints.Text = str(0)		
		pass
	def Txb_ToTalPointsTextChanged(self, sender, e):
		pass
	def Clb_floorPointsSelectedIndexChanged(self, sender, e):
		varZ = self._clb_FloorPoints.CheckedItems.Count
		n = 0
		if varZ != 0:
			for i in range(varZ):
				n += 1
				self._txb_ToTalPoints.Text = str(n)		
		else:
			self._txb_ToTalPoints.Text = str(0)			
		pass
	def Btt_RunClick(self, sender, e):
		floorViews_count = self._clb_FloorView.CheckedItems.Count
		floorViewsPoints_count = self._clb_FloorPoints.CheckedItems.Count
		linkCad_count = self._clb_LinkCad.CheckedItems.Count
		if floorViews_count == floorViewsPoints_count == linkCad_count:
			floorViews = []
			floorPoints = []
			linkCads = []
			for _floorView in self._clb_FloorView.CheckedItems:
				floorViews.append(_floorView)
			for _floorPoint in self._clb_FloorPoints.CheckedItems:
				floorPoints.append(_floorPoint)
			for _linkCadName in self._clb_LinkCad.CheckedItems:
				# if _linkCadName == os.path.exists(filePath) :
				# 	linkCads.append(_linkCadName)
				if _linkCadName in self.filePathMap:
					linkCads.append(self.filePathMap[_linkCadName])				
			for floorView, floorPoint, linkCad in zip(floorViews, floorPoints, linkCads):
				options = DWGImportOptions()
				options.AutoCorrectAlmostVHLines = True
				options.CustomScale = 1
				options.OrientToView = True
				options.ThisViewOnly = False
				options.VisibleLayersOnly = False
				options.ColorMode = ImportColorMode.Preserved
				options.Placement = ImportPlacement.Origin
				options.Unit = ImportUnit.Millimeter
				linkedElem = clr.Reference[ElementId]()
				TransactionManager.Instance.EnsureInTransaction(doc)
				doc.Link(linkCad, options, floorView, linkedElem)
				TransactionManager.Instance.TransactionTaskDone()
				importInst = doc.GetElement(linkedElem.Value)
				TransactionManager.Instance.EnsureInTransaction(doc)
				importInst.Pinned = False
				TransactionManager.Instance.TransactionTaskDone()
				CADLink = doc.GetElement(importInst.GetTypeId())
				pointOrigin = XYZ(0,0,0)
				moveVector = floorPoint.ToRevitType() - pointOrigin     
				TransactionManager.Instance.EnsureInTransaction(doc)
				elementTransform = Transform.CreateTranslation(moveVector)
				importInst.Location.Move(moveVector)
				TransactionManager.Instance.TransactionTaskDone()	
		pass
	def Btt_CancleClick(self, sender, e):
		self.Close()
		pass
#endregion
f = MainForm()
Application.Run(f)