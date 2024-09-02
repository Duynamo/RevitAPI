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
		self._grb_exportFamily.SuspendLayout()
		self._grb_DirectFolderLink.SuspendLayout()
		self.SuspendLayout()
		# 
		# grb_exportFamily
		# 
		self._grb_exportFamily.Controls.Add(self._clb_GenericModels)
		self._grb_exportFamily.Controls.Add(self._clb_PipeAccessories)
		self._grb_exportFamily.Controls.Add(self._clb_pipeFittings)
		self._grb_exportFamily.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
		self._grb_exportFamily.ForeColor = System.Drawing.Color.Red
		self._grb_exportFamily.Location = System.Drawing.Point(12, 21)
		self._grb_exportFamily.Name = "grb_exportFamily"
		self._grb_exportFamily.Size = System.Drawing.Size(576, 166)
		self._grb_exportFamily.TabIndex = 0
		self._grb_exportFamily.TabStop = False
		self._grb_exportFamily.Text = "Family"
		# 
		# clb_pipeFittings
		# 
		self._clb_pipeFittings.DisplayMember = 'Name'
		self._clb_pipeFittings.CheckOnClick = True
		self._clb_pipeFittings.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_pipeFittings.ForeColor = System.Drawing.Color.Blue
		self._clb_pipeFittings.FormattingEnabled = True
		self._clb_pipeFittings.Location = System.Drawing.Point(6, 26)
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
		self._clb_PipeAccessories.Location = System.Drawing.Point(194, 26)
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
		self._clb_GenericModels.ForeColor = System.Drawing.Color.Red
		self._clb_GenericModels.FormattingEnabled = True
		self._clb_GenericModels.Location = System.Drawing.Point(382, 26)
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
		self._lb_FVC.Location = System.Drawing.Point(18, 287)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(48, 16)
		self._lb_FVC.TabIndex = 1
		self._lb_FVC.Text = "@FVC"
		# 
		# btt_EXPORT
		# 
		self._btt_EXPORT.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_EXPORT.ForeColor = System.Drawing.Color.Red
		self._btt_EXPORT.Location = System.Drawing.Point(399, 258)
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
		self._btt_CANCLE.Location = System.Drawing.Point(490, 258)
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
		self._grb_DirectFolderLink.Location = System.Drawing.Point(12, 194)
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
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(597, 311)
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
		pass

	def Btt_EXPORTClick(self, sender, e):
		pass

	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass
#endregion

f = MainForm()
Application.Run(f)


