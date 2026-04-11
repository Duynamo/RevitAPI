import clr
import sys 
import System   
import math
import os
import time

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
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app   = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view  = doc.ActiveView

#region__wrapper
class FamilySymbolWrapper(object):
    def __init__(self, symbol):
        self.Symbol = symbol
        try:
            family_name = symbol.Family.Name
        except:
            family_name = ""
        try:
            type_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        except:
            try:
                type_name = symbol.Name
            except:
                type_name = ""
        if family_name and type_name:
            self.Name = "{} : {}".format(family_name, type_name)
        elif family_name:
            self.Name = family_name
        elif type_name:
            self.Name = type_name
        else:
            self.Name = "ID_{}".format(symbol.Id.IntegerValue)
    def __str__(self):
        return self.Name
#endregion

#region__function
def getAllPipeFittingsInPJ(doc):
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsElementType().ToElements()
    result = []
    for f in collector:
        try:
            result.append(FamilySymbolWrapper(f))
        except:
            pass
    return result

def getAllPipeAccessoriesInPJ(doc):
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsElementType().ToElements()
    result = []
    for a in collector:
        try:
            result.append(FamilySymbolWrapper(a))
        except:
            pass
    return result

def getAllGenericModelInPJ(doc):
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsElementType().ToElements()
    result = []
    for g in collector:
        try:
            result.append(FamilySymbolWrapper(g))
        except:
            pass
    return result

def setHorizontalExtent(clb):
    max_width = 0
    font      = clb.Font
    graphics  = clb.CreateGraphics()
    for i in range(clb.Items.Count):
        item_text  = str(clb.Items[i].Name) if hasattr(clb.Items[i], 'Name') else str(clb.Items[i])
        text_width = int(graphics.MeasureString(item_text, font).Width) + 20
        if text_width > max_width:
            max_width = text_width
    graphics.Dispose()
    clb.HorizontalExtent = max_width
#endregion

pipeFittingsCollector    = getAllPipeFittingsInPJ(doc)
pipeAccessoriesCollector = getAllPipeAccessoriesInPJ(doc)
genericModelCollector    = getAllGenericModelInPJ(doc)

#region__UI
class MainForm(Form):
	def __init__(self):
		self._lastClickIndex_pipeFittings    = -1
		self._lastClickIndex_pipeAccessories = -1
		self._lastClickIndex_genericModels   = -1
		self._isExporting                    = False
		self._cancelRequested                = False
		self.InitializeComponent()

	def InitializeComponent(self):
		self._grb_exportFamily        = System.Windows.Forms.GroupBox()
		self._clb_pipeFittings        = System.Windows.Forms.CheckedListBox()
		self._clb_PipeAccessories     = System.Windows.Forms.CheckedListBox()
		self._clb_GenericModels       = System.Windows.Forms.CheckedListBox()
		self._lb_FVC                  = System.Windows.Forms.Label()
		self._btt_EXPORT              = System.Windows.Forms.Button()
		self._btt_CANCLE              = System.Windows.Forms.Button()
		self._grb_DirectFolderLink    = System.Windows.Forms.GroupBox()
		self._btt_selectFolder        = System.Windows.Forms.Button()
		self._txb_folder              = System.Windows.Forms.TextBox()
		self._lb_pipeFittings         = System.Windows.Forms.Label()
		self._lb_pipeAccessories      = System.Windows.Forms.Label()
		self._lb_genericModel         = System.Windows.Forms.Label()
		self._lb_status               = System.Windows.Forms.Label()
		self._progressBar             = System.Windows.Forms.ProgressBar()
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
		self._grb_exportFamily.Font        = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
		self._grb_exportFamily.ForeColor   = System.Drawing.Color.Red
		self._grb_exportFamily.Location    = System.Drawing.Point(12, 21)
		self._grb_exportFamily.Name        = "grb_exportFamily"
		self._grb_exportFamily.Size        = System.Drawing.Size(576, 209)
		self._grb_exportFamily.TabIndex    = 0
		self._grb_exportFamily.TabStop     = False
		self._grb_exportFamily.Text        = "Family"
		# 
		# clb_pipeFittings
		# 
		self._clb_pipeFittings.DisplayMember      = 'Name'
		self._clb_pipeFittings.CheckOnClick        = False          # Shift+Click tự quản lý
		self._clb_pipeFittings.HorizontalScrollbar = True
		self._clb_pipeFittings.Font                = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_pipeFittings.ForeColor           = System.Drawing.Color.Blue
		self._clb_pipeFittings.FormattingEnabled   = True
		self._clb_pipeFittings.Location            = System.Drawing.Point(6, 52)
		self._clb_pipeFittings.Name                = "clb_pipeFittings"
		self._clb_pipeFittings.Size                = System.Drawing.Size(182, 124)
		self._clb_pipeFittings.TabIndex            = 0
		self._clb_pipeFittings.Items.AddRange(System.Array[System.Object](pipeFittingsCollector))
		self._clb_pipeFittings.MouseDown          += self.clb_pipeFittings_MouseDown
		# 
		# clb_PipeAccessories
		# 
		self._clb_PipeAccessories.DisplayMember      = 'Name'
		self._clb_PipeAccessories.CheckOnClick        = False
		self._clb_PipeAccessories.HorizontalScrollbar = True
		self._clb_PipeAccessories.Font                = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_PipeAccessories.ForeColor           = System.Drawing.Color.Black
		self._clb_PipeAccessories.FormattingEnabled   = True
		self._clb_PipeAccessories.Location            = System.Drawing.Point(194, 52)
		self._clb_PipeAccessories.Name                = "clb_PipeAccessories"
		self._clb_PipeAccessories.Size                = System.Drawing.Size(182, 124)
		self._clb_PipeAccessories.TabIndex            = 1
		self._clb_PipeAccessories.Items.AddRange(System.Array[System.Object](pipeAccessoriesCollector))
		self._clb_PipeAccessories.MouseDown          += self.clb_PipeAccessories_MouseDown
		# 
		# clb_GenericModels
		# 
		self._clb_GenericModels.DisplayMember      = 'Name'
		self._clb_GenericModels.CheckOnClick        = False
		self._clb_GenericModels.HorizontalScrollbar = True
		self._clb_GenericModels.Font                = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_GenericModels.ForeColor           = System.Drawing.Color.Green
		self._clb_GenericModels.FormattingEnabled   = True
		self._clb_GenericModels.Location            = System.Drawing.Point(382, 52)
		self._clb_GenericModels.Name                = "clb_GenericModels"
		self._clb_GenericModels.Size                = System.Drawing.Size(182, 124)
		self._clb_GenericModels.TabIndex            = 1
		self._clb_GenericModels.Items.AddRange(System.Array[System.Object](genericModelCollector))
		self._clb_GenericModels.MouseDown          += self.clb_GenericModels_MouseDown
		# 
		# lb_FVC
		# 
		self._lb_FVC.Font      = System.Drawing.Font("Meiryo UI", 4.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_FVC.ForeColor = System.Drawing.Color.Black
		self._lb_FVC.Location  = System.Drawing.Point(18, 370)
		self._lb_FVC.Name      = "lb_FVC"
		self._lb_FVC.Size      = System.Drawing.Size(48, 16)
		self._lb_FVC.TabIndex  = 1
		self._lb_FVC.Text      = "@FVC"
		# 
		# btt_EXPORT
		# 
		self._btt_EXPORT.Font                    = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_EXPORT.ForeColor               = System.Drawing.Color.Red
		self._btt_EXPORT.Location                = System.Drawing.Point(399, 338)
		self._btt_EXPORT.Name                    = "btt_EXPORT"
		self._btt_EXPORT.Size                    = System.Drawing.Size(85, 45)
		self._btt_EXPORT.TabIndex                = 2
		self._btt_EXPORT.Text                    = "EXPORT"
		self._btt_EXPORT.UseVisualStyleBackColor = True
		self._btt_EXPORT.Click                  += self.Btt_EXPORTClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Font                    = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.ForeColor               = System.Drawing.Color.Red
		self._btt_CANCLE.Location                = System.Drawing.Point(490, 338)
		self._btt_CANCLE.Name                    = "btt_CANCLE"
		self._btt_CANCLE.Size                    = System.Drawing.Size(85, 45)
		self._btt_CANCLE.TabIndex                = 2
		self._btt_CANCLE.Text                    = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click                  += self.Btt_CANCLEClick
		# 
		# grb_DirectFolderLink
		# 
		self._grb_DirectFolderLink.Controls.Add(self._txb_folder)
		self._grb_DirectFolderLink.Controls.Add(self._btt_selectFolder)
		self._grb_DirectFolderLink.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._grb_DirectFolderLink.ForeColor = System.Drawing.Color.Red
		self._grb_DirectFolderLink.Location  = System.Drawing.Point(12, 236)
		self._grb_DirectFolderLink.Name      = "grb_DirectFolderLink"
		self._grb_DirectFolderLink.Size      = System.Drawing.Size(573, 53)
		self._grb_DirectFolderLink.TabIndex  = 3
		self._grb_DirectFolderLink.TabStop   = False
		self._grb_DirectFolderLink.Text      = "Folder"
		# 
		# btt_selectFolder
		# 
		self._btt_selectFolder.Font                    = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_selectFolder.ForeColor               = System.Drawing.Color.Blue
		self._btt_selectFolder.Location                = System.Drawing.Point(6, 24)
		self._btt_selectFolder.Name                    = "btt_selectFolder"
		self._btt_selectFolder.Size                    = System.Drawing.Size(86, 23)
		self._btt_selectFolder.TabIndex                = 0
		self._btt_selectFolder.Text                    = "Select Folder"
		self._btt_selectFolder.UseVisualStyleBackColor = True
		self._btt_selectFolder.Click                  += self.Btt_selectFolderClick
		# 
		# txb_folder
		# 
		self._txb_folder.Font        = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_folder.Location    = System.Drawing.Point(99, 21)
		self._txb_folder.Name        = "txb_folder"
		self._txb_folder.Size        = System.Drawing.Size(465, 24)
		self._txb_folder.TabIndex    = 1
		self._txb_folder.TextChanged += self.Txb_folderTextChanged
		# 
		# lb_pipeFittings
		# 
		self._lb_pipeFittings.Font      = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_pipeFittings.ForeColor = System.Drawing.Color.Black
		self._lb_pipeFittings.Location  = System.Drawing.Point(6, 33)
		self._lb_pipeFittings.Name      = "lb_pipeFittings"
		self._lb_pipeFittings.Size      = System.Drawing.Size(168, 16)
		self._lb_pipeFittings.TabIndex  = 4
		self._lb_pipeFittings.Text      = "Pipe Fittings"
		# 
		# lb_pipeAccessories
		# 
		self._lb_pipeAccessories.Font      = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_pipeAccessories.ForeColor = System.Drawing.Color.Black
		self._lb_pipeAccessories.Location  = System.Drawing.Point(194, 33)
		self._lb_pipeAccessories.Name      = "lb_pipeAccessories"
		self._lb_pipeAccessories.Size      = System.Drawing.Size(168, 16)
		self._lb_pipeAccessories.TabIndex  = 4
		self._lb_pipeAccessories.Text      = "Pipe Accessories"
		# 
		# lb_genericModel
		# 
		self._lb_genericModel.Font      = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_genericModel.ForeColor = System.Drawing.Color.Black
		self._lb_genericModel.Location  = System.Drawing.Point(382, 33)
		self._lb_genericModel.Name      = "lb_genericModel"
		self._lb_genericModel.Size      = System.Drawing.Size(168, 16)
		self._lb_genericModel.TabIndex  = 4
		self._lb_genericModel.Text      = "Generic Models"
		#
		# lb_status
		#
		self._lb_status.Font      = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_status.ForeColor = System.Drawing.Color.Gray
		self._lb_status.Location  = System.Drawing.Point(12, 298)
		self._lb_status.Name      = "lb_status"
		self._lb_status.Size      = System.Drawing.Size(573, 16)
		self._lb_status.TabIndex  = 5
		self._lb_status.Text      = "Ready"
		#
		# progressBar
		#
		self._progressBar.Location = System.Drawing.Point(12, 316)
		self._progressBar.Name     = "progressBar"
		self._progressBar.Size     = System.Drawing.Size(573, 14)
		self._progressBar.Minimum  = 0
		self._progressBar.Maximum  = 100
		self._progressBar.Value    = 0
		self._progressBar.Style    = System.Windows.Forms.ProgressBarStyle.Continuous
		# 
		# MainForm
		# 
		self.ClientSize      = System.Drawing.Size(597, 395)
		self.Controls.Add(self._grb_DirectFolderLink)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_EXPORT)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._lb_status)
		self.Controls.Add(self._progressBar)
		self.Controls.Add(self._grb_exportFamily)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name            = "MainForm"
		self.Text            = "Export Family"
		self.TopMost         = True
		self.FormClosing    += self.MainForm_FormClosing
		self._grb_exportFamily.ResumeLayout(False)
		self._grb_DirectFolderLink.ResumeLayout(False)
		self._grb_DirectFolderLink.PerformLayout()
		self.ResumeLayout(False)
		setHorizontalExtent(self._clb_pipeFittings)
		setHorizontalExtent(self._clb_PipeAccessories)
		setHorizontalExtent(self._clb_GenericModels)
#endregion

#region__binding
	# ── Shift+Click ──────────────────────────────────────────
	def _handleShiftClick(self, clb, e, lastIndexRef):
		index = clb.IndexFromPoint(e.X, e.Y)
		if index < 0:
			return lastIndexRef
		isShift = (System.Windows.Forms.Control.ModifierKeys == System.Windows.Forms.Keys.Shift)
		if isShift and lastIndexRef >= 0:
			target_state = clb.GetItemChecked(lastIndexRef)
			start = min(lastIndexRef, index)
			end   = max(lastIndexRef, index)
			for i in range(start, end + 1):
				clb.SetItemChecked(i, target_state)
			return lastIndexRef
		else:
			current_state = clb.GetItemChecked(index)
			clb.SetItemChecked(index, not current_state)
			return index

	def clb_pipeFittings_MouseDown(self, sender, e):
		self._lastClickIndex_pipeFittings = self._handleShiftClick(
			self._clb_pipeFittings, e, self._lastClickIndex_pipeFittings)

	def clb_PipeAccessories_MouseDown(self, sender, e):
		self._lastClickIndex_pipeAccessories = self._handleShiftClick(
			self._clb_PipeAccessories, e, self._lastClickIndex_pipeAccessories)

	def clb_GenericModels_MouseDown(self, sender, e):
		self._lastClickIndex_genericModels = self._handleShiftClick(
			self._clb_GenericModels, e, self._lastClickIndex_genericModels)

	def Txb_folderTextChanged(self, sender, e):
		pass

	def Btt_selectFolderClick(self, sender, e):
		fileDialog = FolderBrowserDialog()
		if fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
			self._txb_folder.Text = fileDialog.SelectedPath

	# ── FormClosing: chặn đóng khi đang export ───────────────
	def MainForm_FormClosing(self, sender, e):
		if self._isExporting:
			result = MessageBox.Show(
				"Export is in progress.\nDo you want to cancel and stop exporting?",
				"Cancel Export",
				MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning)
			if result == System.Windows.Forms.DialogResult.Yes:
				self._cancelRequested     = True
				self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
				self._lb_status.Text      = "Cancelling... waiting for current family to finish."
			e.Cancel = True  # giữ form, vòng lặp tự đóng sau khi break

	# ── EXPORT ───────────────────────────────────────────────
	def Btt_EXPORTClick(self, sender, e):
		directFolder = self._txb_folder.Text.strip()
		if not directFolder:
			MessageBox.Show("Please select an output folder first.",
				"No Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			return
		if not os.path.isdir(directFolder):
			MessageBox.Show("The selected folder does not exist.",
				"Invalid Folder", MessageBoxButtons.OK, MessageBoxIcon.Error)
			return
		if self._isExporting:
			return

		# Thu thập family được chọn (loại trùng lặp bằng Id)
		selectedFamilies = {}
		def process_items(items):
			for wrapper in items:
				symbol = wrapper.Symbol
				if not isinstance(symbol, FamilySymbol):
					continue
				try:
					family = symbol.Family
					if not family.IsEditable:
						continue
					fid = family.Id.IntegerValue
					if fid not in selectedFamilies:
						selectedFamilies[fid] = family
				except:
					pass

		process_items(self._clb_pipeFittings.CheckedItems)
		process_items(self._clb_PipeAccessories.CheckedItems)
		process_items(self._clb_GenericModels.CheckedItems)

		families = list(selectedFamilies.values())
		if not families:
			MessageBox.Show("No families selected.", "Empty Selection",
				MessageBoxButtons.OK, MessageBoxIcon.Warning)
			return

		# Reset trạng thái
		self._isExporting         = True
		self._cancelRequested     = False
		self._btt_EXPORT.Enabled  = False
		self._btt_CANCLE.Text     = "CANCEL"
		self._btt_CANCLE.ForeColor = System.Drawing.Color.OrangeRed
		self._progressBar.Maximum = len(families)
		self._progressBar.Value   = 0

		save_options = SaveAsOptions()
		save_options.OverwriteExistingFile = True

		success  = []
		failed   = []
		total    = len(families)

		# Throttle DoEvents
		DOEVENTS_INTERVAL = 0.3
		last_doevents     = time.time()

		for i, family in enumerate(families):
			# Kiểm tra cancel trước mỗi family
			if self._cancelRequested:
				failed.append("Cancelled at {}/{}".format(i, total))
				break

			family_name = ""
			try:
				family_name = family.Name
				file_path   = os.path.join(directFolder, family_name + ".rfa")
				family_doc  = doc.EditFamily(family)
				family_doc.SaveAs(file_path, save_options)
				family_doc.Close(False)
				success.append(family_name)
			except Exception as ex:
				failed.append("{} → {}".format(family_name or "?", str(ex)))

			# Throttle: cập nhật UI theo interval
			now = time.time()
			if now - last_doevents >= DOEVENTS_INTERVAL:
				self._progressBar.Value   = i + 1
				self._lb_status.ForeColor = System.Drawing.Color.DarkOrange
				self._lb_status.Text      = "[{}/{}] {}".format(i + 1, total, family_name)
				System.Windows.Forms.Application.DoEvents()
				last_doevents = now

		# Cập nhật lần cuối
		self._progressBar.Value = len(success)

		# Khôi phục UI
		self._isExporting          = False
		self._cancelRequested      = False
		self._btt_EXPORT.Enabled   = True
		self._btt_CANCLE.Text      = "CANCLE"
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
		self._btt_CANCLE.Enabled   = True

		# Hiển thị kết quả
		if failed:
			self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
			self._lb_status.Text = "Done — {} exported, {} failed/cancelled.".format(
				len(success), len(failed))
		else:
			self._lb_status.ForeColor = System.Drawing.Color.Green
			self._lb_status.Text = "Done — {} families exported successfully!".format(len(success))

		# Gỡ FormClosing để có thể đóng bình thường, rồi re-attach
		self.FormClosing -= self.MainForm_FormClosing
		self.FormClosing += self.MainForm_FormClosing

	# ── CANCEL / CLOSE ───────────────────────────────────────
	def Btt_CANCLEClick(self, sender, e):
		if self._isExporting:
			result = MessageBox.Show(
				"Export is in progress.\nDo you want to cancel and stop exporting?",
				"Cancel Export",
				MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning)
			if result == System.Windows.Forms.DialogResult.Yes:
				self._cancelRequested     = True
				self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
				self._lb_status.Text      = "Cancelling... waiting for current family to finish."
				self._btt_CANCLE.Enabled  = False
		else:
			self.FormClosing -= self.MainForm_FormClosing
			self.Close()
#endregion

f = MainForm()
Application.Run(f)