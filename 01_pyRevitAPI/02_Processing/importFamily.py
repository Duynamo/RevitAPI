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
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#region__function
class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._grb_directFolder = System.Windows.Forms.GroupBox()
		self._btt_selectFolder = System.Windows.Forms.Button()
		self._txb_directFolder = System.Windows.Forms.TextBox()
		self._gbr_Family = System.Windows.Forms.GroupBox()
		self._lb_FVC = System.Windows.Forms.Label()
		self._btt_IMPORT = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._clb_desFamily = System.Windows.Forms.CheckedListBox()
		self._grb_directFolder.SuspendLayout()
		self._gbr_Family.SuspendLayout()
		self.SuspendLayout()
		# 
		# grb_directFolder
		# 
		self._grb_directFolder.Controls.Add(self._txb_directFolder)
		self._grb_directFolder.Controls.Add(self._btt_selectFolder)
		self._grb_directFolder.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._grb_directFolder.ForeColor = System.Drawing.Color.Red
		self._grb_directFolder.Location = System.Drawing.Point(13, 13)
		self._grb_directFolder.Name = "grb_directFolder"
		self._grb_directFolder.Size = System.Drawing.Size(555, 67)
		self._grb_directFolder.TabIndex = 0
		self._grb_directFolder.TabStop = False
		self._grb_directFolder.Text = "Folder"
		# 
		# btt_selectFolder
		# 
		self._btt_selectFolder.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_selectFolder.ForeColor = System.Drawing.Color.Blue
		self._btt_selectFolder.Location = System.Drawing.Point(437, 26)
		self._btt_selectFolder.Name = "btt_selectFolder"
		self._btt_selectFolder.Size = System.Drawing.Size(112, 25)
		self._btt_selectFolder.TabIndex = 0
		self._btt_selectFolder.Text = "Select Folder"
		self._btt_selectFolder.UseVisualStyleBackColor = True
		self._btt_selectFolder.Click += self.Btt_selectFolderClick
		# 
		# txb_directFolder
		# 
		self._txb_directFolder.Location = System.Drawing.Point(4, 26)
		self._txb_directFolder.Name = "txb_directFolder"
		self._txb_directFolder.Size = System.Drawing.Size(424, 27)
		self._txb_directFolder.TabIndex = 1
		self._txb_directFolder.TextChanged += self.Txb_directFolderTextChanged
		# 
		# gbr_Family
		# 
		self._gbr_Family.Controls.Add(self._clb_desFamily)
		self._gbr_Family.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._gbr_Family.ForeColor = System.Drawing.Color.Red
		self._gbr_Family.Location = System.Drawing.Point(13, 86)
		self._gbr_Family.Name = "gbr_Family"
		self._gbr_Family.Size = System.Drawing.Size(428, 142)
		self._gbr_Family.TabIndex = 1
		self._gbr_Family.TabStop = False
		self._gbr_Family.Text = "Family"
		self._gbr_Family.Enter += self.Gbr_FamilyEnter
		# 
		# lb_FVC
		# 
		self._lb_FVC.Font = System.Drawing.Font("Meiryo UI", 4.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_FVC.Location = System.Drawing.Point(12, 231)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(47, 23)
		self._lb_FVC.TabIndex = 2
		self._lb_FVC.Text = "@FVC"
		self._lb_FVC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# btt_IMPORT
		# 
		self._btt_IMPORT.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_IMPORT.ForeColor = System.Drawing.Color.Blue
		self._btt_IMPORT.Location = System.Drawing.Point(450, 172)
		self._btt_IMPORT.Name = "btt_IMPORT"
		self._btt_IMPORT.Size = System.Drawing.Size(112, 25)
		self._btt_IMPORT.TabIndex = 3
		self._btt_IMPORT.Text = "IMPORT"
		self._btt_IMPORT.UseVisualStyleBackColor = True
		self._btt_IMPORT.Click += self.Btt_IMPORTClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Blue
		self._btt_CANCLE.Location = System.Drawing.Point(450, 203)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(112, 25)
		self._btt_CANCLE.TabIndex = 3
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# clb_desFamily
		# 
		# self._clb_desFamily.DisplayMember = 'Name'
		self._clb_desFamily.CheckOnClick = True
		self._clb_desFamily.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_desFamily.FormattingEnabled = True
		self._clb_desFamily.Location = System.Drawing.Point(7, 27)
		self._clb_desFamily.Name = "clb_desFamily"
		self._clb_desFamily.Size = System.Drawing.Size(421, 112)
		self._clb_desFamily.TabIndex = 0
		self._clb_desFamily.SelectedIndexChanged += self.Clb_desFamilySelectedIndexChanged
		# self._clb_desFamily.Items.AddRange(System.Array[System.Object](self.familyFiles))
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(580, 257)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_IMPORT)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._gbr_Family)
		self.Controls.Add(self._grb_directFolder)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.Text = "Import Family"
		self.TopMost = True
		self._grb_directFolder.ResumeLayout(False)
		self._grb_directFolder.PerformLayout()
		self._gbr_Family.ResumeLayout(False)
		self.ResumeLayout(False)
#endrerion
#region__biding
	def Txb_directFolderTextChanged(self, sender, e):
		pass

	def Gbr_FamilyEnter(self, sender, e):
		pass

	def Btt_selectFolderClick(self, sender, e):
		fileDialog = FolderBrowserDialog()
		fileDialog.ShowDialog()
		self._txb_directFolder.Text = fileDialog.SelectedPath
		self.populateFamiliesInList()
		pass
	def Clb_desFamilySelectedIndexChanged(self, sender, e):			
		pass
	def populateFamiliesInList(self):
		self._clb_desFamily.Items.Clear()
		directory = self._txb_directFolder.Text
		if os.path.isdir(directory):
			familyFiles = [f for f in os.listdir(directory) if f.endswith('.rfa')]
			self._clb_desFamily.Items.AddRange(System.Array[System.Object](familyFiles))	
		pass
	def Btt_IMPORTClick(self, sender, e):    
		selectedFamilies = [item for item in self._clb_desFamily.CheckedItems]
		directory = self._txb_directFolder.Text
		importedFamilies = []
		booleans = []
		elementList = []
		# Start a transaction to load the families
		TransactionManager.Instance.EnsureInTransaction(doc)
		for familyFile in selectedFamilies:
			try:
				familyPath = os.path.join(directory, familyFile)
				loadedFamily = clr.Reference[Family]()
				doc.LoadFamily(familyPath, loadedFamily)
				booleans.append(True)
			except: 
				booleans.append(False)
		TransactionManager.Instance.TransactionTaskDone()	
		collector = FilteredElementCollector(doc).OfClass(Family)
		for item in collector.ToElements():
			if item.Name in [os.path.splitext(fam)[0] for fam in selectedFamilies]:  # Compare by family names
				typeList = list()
				for famTypeId in item.GetFamilySymbolIds():
					typeList.append(doc.GetElement(famTypeId).ToDSType(True))
				elementList.append(typeList)
		TransactionManager.Instance.TransactionTaskDone()
		self.Close()  # Close the form
		pass
	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass
#endregion

f = MainForm()
Application.Run(f)


