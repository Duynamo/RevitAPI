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
def getAllPipeFittingsInPJ(doc):
    pipeFittingsName = []
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsElementType().ToElements()
    for fitting in collector:
        fittingNameParam = fitting.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if fittingNameParam:
            fittingName = fittingNameParam.AsString()
            pipeFittingsName.append(fittingName)
    return collector
def getAllPipeAccessoriesInPJ(doc):
    accessoriesName = []
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsElementType().ToElements()
    for a in collector:
        accessoryNameParam = a.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if accessoryNameParam:
            accessoryName = accessoryNameParam.AsString()
            accessoriesName.append(accessoryName)
    return collector
def getAllGenericModelInPJ(doc):
    genericModelsName = []
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsElementType().ToElements()
    for g in collector:
        # Use SYMBOL_NAME_PARAM to get the type name
        genericModelsNameParam = g.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if genericModelsNameParam:
            genericModelName = genericModelsNameParam.AsString()
            genericModelsName.append(genericModelName)
    return collector
#endregion
pipeFittingsCollector = getAllPipeFittingsInPJ(doc)
pipeAccessoriesCollector = getAllPipeAccessoriesInPJ(doc)
genericModelCollector = getAllGenericModelInPJ(doc)
#region__UI
class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._grb_exportFamily = System.Windows.Forms.GroupBox()
		self._clb_pipeFittings = System.Windows.Forms.CheckedListBox()
		self._clb_PipeAccessories = System.Windows.Forms.CheckedListBox()
		self._clb_GenericModels = System.Windows.Forms.CheckedListBox()
		self._lb_FVC = System.Windows.Forms.Label()
		self._btt_EXPORT = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._grb_DirectFolderLink = System.Windows.Forms.GroupBox()
		self._btt_selectFolder = System.Windows.Forms.Button()
		self._txb_folder = System.Windows.Forms.TextBox()
		self._lb_pipeFittings = System.Windows.Forms.Label()
		self._lb_pipeAccessories = System.Windows.Forms.Label()
		self._lb_genericModel = System.Windows.Forms.Label()
		self._grb_exportFamily.SuspendLayout()
		self._grb_DirectFolderLink.SuspendLayout()
		self.SuspendLayout()
		# 
		# grb_exportFamily
		# 
		self._grb_exportFamily.Controls.Add(self._lb_genericModel)
		self._grb_exportFamily.Controls.Add(self._lb_pipeAccessories)
		self._grb_exportFamily.Controls.Add(self._lb_pipeFittings)
		self._grb_exportFamily.Controls.Add(self._clb_GenericModels)
		self._grb_exportFamily.Controls.Add(self._clb_PipeAccessories)
		self._grb_exportFamily.Controls.Add(self._clb_pipeFittings)
		self._grb_exportFamily.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
		self._grb_exportFamily.ForeColor = System.Drawing.Color.Red
		self._grb_exportFamily.Location = System.Drawing.Point(12, 21)
		self._grb_exportFamily.Name = "grb_exportFamily"
		self._grb_exportFamily.Size = System.Drawing.Size(576, 209)
		self._grb_exportFamily.TabIndex = 0
		self._grb_exportFamily.TabStop = False
		self._grb_exportFamily.Text = "Family"
		# 
		# clb_pipeFittings
		# 
		self._clb_pipeFittings.DisplayMember = 'Name'
		self._clb_pipeFittings.CheckOnClick = True
		self._clb_pipeFittings.CheckOnClick = True
		self._clb_pipeFittings.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_pipeFittings.ForeColor = System.Drawing.Color.Blue
		self._clb_pipeFittings.FormattingEnabled = True
		self._clb_pipeFittings.Location = System.Drawing.Point(6, 52)
		self._clb_pipeFittings.Name = "clb_pipeFittings"
		self._clb_pipeFittings.Size = System.Drawing.Size(182, 124)
		self._clb_pipeFittings.TabIndex = 0
		self._clb_pipeFittings.SelectedIndexChanged += self.PipeFittingsSelectedIndexChanged
		self._clb_pipeFittings.Items.AddRange(System.Array[System.Object](pipeFittingsCollector))
		# 
		# clb_PipeAccessories
		# 
		self._clb_PipeAccessories.DisplayMember = 'Name'
		self._clb_PipeAccessories.CheckOnClick = True
		self._clb_PipeAccessories.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_PipeAccessories.ForeColor = System.Drawing.Color.Black
		self._clb_PipeAccessories.FormattingEnabled = True
		self._clb_PipeAccessories.Location = System.Drawing.Point(194, 52)
		self._clb_PipeAccessories.Name = "clb_PipeAccessories"
		self._clb_PipeAccessories.Size = System.Drawing.Size(182, 124)
		self._clb_PipeAccessories.TabIndex = 1
		self._clb_PipeAccessories.SelectedIndexChanged += self.PipeAccessoriesSelectedIndexChanged
		self._clb_PipeAccessories.Items.AddRange(System.Array[System.Object](pipeAccessoriesCollector))
		# 
		# clb_GenericModels
		# 
		self._clb_GenericModels.DisplayMember = 'Name'
		self._clb_GenericModels.CheckOnClick = True
		self._clb_GenericModels.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_GenericModels.ForeColor = System.Drawing.Color.Green
		self._clb_GenericModels.FormattingEnabled = True
		self._clb_GenericModels.Location = System.Drawing.Point(382, 52)
		self._clb_GenericModels.Name = "clb_GenericModels"
		self._clb_GenericModels.Size = System.Drawing.Size(182, 124)
		self._clb_GenericModels.TabIndex = 1
		self._clb_GenericModels.SelectedIndexChanged += self.GenericModelsSelectedIndexChanged
		self._clb_GenericModels.Items.AddRange(System.Array[System.Object](genericModelCollector))
		# 
		# lb_FVC
		# 
		self._lb_FVC.Font = System.Drawing.Font("Meiryo UI", 4.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_FVC.ForeColor = System.Drawing.Color.Black
		self._lb_FVC.Location = System.Drawing.Point(18, 324)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(48, 16)
		self._lb_FVC.TabIndex = 1
		self._lb_FVC.Text = "@FVC"
		# 
		# btt_EXPORT
		# 
		self._btt_EXPORT.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_EXPORT.ForeColor = System.Drawing.Color.Red
		self._btt_EXPORT.Location = System.Drawing.Point(399, 295)
		self._btt_EXPORT.Name = "btt_EXPORT"
		self._btt_EXPORT.Size = System.Drawing.Size(85, 45)
		self._btt_EXPORT.TabIndex = 2
		self._btt_EXPORT.Text = "EXPORT"
		self._btt_EXPORT.UseVisualStyleBackColor = True
		self._btt_EXPORT.Click += self.Btt_EXPORTClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
		self._btt_CANCLE.Location = System.Drawing.Point(490, 295)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(85, 45)
		self._btt_CANCLE.TabIndex = 2
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# grb_DirectFolderLink
		# 
		self._grb_DirectFolderLink.Controls.Add(self._txb_folder)
		self._grb_DirectFolderLink.Controls.Add(self._btt_selectFolder)
		self._grb_DirectFolderLink.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._grb_DirectFolderLink.ForeColor = System.Drawing.Color.Red
		self._grb_DirectFolderLink.Location = System.Drawing.Point(12, 236)
		self._grb_DirectFolderLink.Name = "grb_DirectFolderLink"
		self._grb_DirectFolderLink.Size = System.Drawing.Size(573, 53)
		self._grb_DirectFolderLink.TabIndex = 3
		self._grb_DirectFolderLink.TabStop = False
		self._grb_DirectFolderLink.Text = "Folder"
		# 
		# btt_selectFolder
		# 
		self._btt_selectFolder.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_selectFolder.ForeColor = System.Drawing.Color.Blue
		self._btt_selectFolder.Location = System.Drawing.Point(6, 24)
		self._btt_selectFolder.Name = "btt_selectFolder"
		self._btt_selectFolder.Size = System.Drawing.Size(86, 23)
		self._btt_selectFolder.TabIndex = 0
		self._btt_selectFolder.Text = "Select Folder"
		self._btt_selectFolder.UseVisualStyleBackColor = True
		self._btt_selectFolder.Click += self.Btt_selectFolderClick
		# 
		# txb_folder
		# 
		self._txb_folder.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_folder.Location = System.Drawing.Point(99, 21)
		self._txb_folder.Name = "txb_folder"
		self._txb_folder.Size = System.Drawing.Size(465, 24)
		self._txb_folder.TabIndex = 1
		self._txb_folder.TextChanged += self.Txb_folderTextChanged
		# 
		# lb_pipeFittings
		# 
		self._lb_pipeFittings.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_pipeFittings.ForeColor = System.Drawing.Color.Black
		self._lb_pipeFittings.Location = System.Drawing.Point(6, 33)
		self._lb_pipeFittings.Name = "lb_pipeFittings"
		self._lb_pipeFittings.Size = System.Drawing.Size(168, 16)
		self._lb_pipeFittings.TabIndex = 4
		self._lb_pipeFittings.Text = "Pipe Fittings"
		# 
		# lb_pipeAccessories
		# 
		self._lb_pipeAccessories.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_pipeAccessories.ForeColor = System.Drawing.Color.Black
		self._lb_pipeAccessories.Location = System.Drawing.Point(194, 33)
		self._lb_pipeAccessories.Name = "lb_pipeAccessories"
		self._lb_pipeAccessories.Size = System.Drawing.Size(168, 16)
		self._lb_pipeAccessories.TabIndex = 4
		self._lb_pipeAccessories.Text = "Pipe Accessories"
		# 
		# lb_genericModel
		# 
		self._lb_genericModel.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_genericModel.ForeColor = System.Drawing.Color.Black
		self._lb_genericModel.Location = System.Drawing.Point(382, 33)
		self._lb_genericModel.Name = "lb_genericModel"
		self._lb_genericModel.Size = System.Drawing.Size(168, 16)
		self._lb_genericModel.TabIndex = 4
		self._lb_genericModel.Text = "Generic Models"		
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(597, 348)
		self.Controls.Add(self._grb_DirectFolderLink)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_EXPORT)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._grb_exportFamily)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.Text = "Export Family"
		self.TopMost = True
		self._grb_exportFamily.ResumeLayout(False)
		self._grb_DirectFolderLink.ResumeLayout(False)
		self._grb_DirectFolderLink.PerformLayout()
		self.ResumeLayout(False)
#endrerion
#region__biding
	def PipeFittingsSelectedIndexChanged(self, sender, e):
		pass
	def PipeAccessoriesSelectedIndexChanged(self, sender, e):
		pass
	def GenericModelsSelectedIndexChanged(self, sender, e):
		pass
	def Txb_folderTextChanged(self, sender, e):
		pass
	def Btt_selectFolderClick(self, sender, e):
		fileDialog = FolderBrowserDialog()
		fileDialog.ShowDialog()
		self._txb_folder.Text = fileDialog.SelectedPath
		pass
	def Btt_EXPORTClick(self, sender, e):
		selectedFamilyInstances = []
		selectedFamilyNames = []

		# Collect selected family instances from UI checkboxes
		for f in self._clb_pipeFittings.CheckedItems:
			selectedFamilyInstances.append(f)
		for f in self._clb_PipeAccessories.CheckedItems:
			selectedFamilyInstances.append(f)
		for f in self._clb_GenericModels.CheckedItems:
			selectedFamilyInstances.append(f)

		# Get the direct folder path from the UI
		directFolder = self._txb_folder.Text
		
		# Start a transaction to export families
		TransactionManager.Instance.EnsureInTransaction(doc)
		
		for familyInstance in selectedFamilyInstances:
			# Get the family element directly from the instance
			family = familyInstance.Family

			if family is not None:
				# Retrieve the family name for export
				familyName = family.Name

				# Create the file path for saving
				file_name = os.path.join(directFolder, familyName + ".rfa")
				
				# Ensure the file is saved with overwrite option
				save_options = SaveAsOptions()
				save_options.OverwriteExistingFile = True  

				try:
					# Export the family to the desired location as a .rfa file
					family.Document.SaveAs(file_name, save_options)
					selectedFamilyNames.append(familyName)  # Store exported family name for confirmation
				except Exception as e:
					# print(f"Failed to export family {familyName}: {str(e)}")
					pass

		# End the transaction after exporting
		TransactionManager.Instance.TransactionTaskDone()

		# Close the form after exporting
		self.Close()
		pass
	# def Btt_EXPORTClick(self, sender, e):
	# 	selectedFamily = []
	# 	selectedFamilyName = []
	# 	for f in self._clb_pipeFittings.CheckedItems:
	# 		selectedFamily.append(f)
	# 		familyNameParam = f.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
	# 		if familyNameParam:
	# 			familyName = familyNameParam.AsString()
	# 			selectedFamilyName.append(familyName)
	# 	for f in self._clb_PipeAccessories.CheckedItems:
	# 		selectedFamily.append(f)
	# 		familyNameParam = f.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
	# 		if familyNameParam:
	# 			familyName = familyNameParam.AsString()
	# 			selectedFamilyName.append(familyName)
	# 	for f in self._clb_GenericModels.CheckedItems:
	# 		selectedFamily.append(f)
	# 		familyNameParam = f.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
	# 		if familyNameParam:
	# 			familyName = familyNameParam.AsString()
	# 			selectedFamilyName.append(familyName)		
	# 	directFolder = self._txb_folder.Text
	# 	for family, familyName in zip( selectedFamily,selectedFamilyName) :
	# 		file_name = os.path.join(directFolder, familyName + ".rfa")
	# 		save_options = SaveAsOptions()
	# 		save_options.OverwriteExistingFile = True  
	# 		try:
	# 			family.Document.SaveAs(file_name, save_options)
	# 		except Exception as e:
	# 			pass
	# 	self.Close()		
	# 	pass			
	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass
#endregion

f = MainForm()
Application.Run(f)


