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
def getAllSections(doc):
    collector = FilteredElementCollector(doc).OfClass(ViewSection)
    sections = [view for view in collector if not view.IsTemplate]
    return sections
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

allSectionsView = getAllSections(doc)

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._cbb_LinkCadToSectionView = System.Windows.Forms.ComboBox()
		self._lbl_Section = System.Windows.Forms.Label()
		self._btt_PickPoint = System.Windows.Forms.Button()
		self._btt_SelectCadLink = System.Windows.Forms.Button()
		self._btt_RUN = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._lbb_vitaminD = System.Windows.Forms.Label()
		self._txb_CadLink = System.Windows.Forms.TextBox()
		self._txb_PickPoint = System.Windows.Forms.TextBox()
		self.SuspendLayout()
		# 
		# cbb_LinkCadToSectionView
		# 
		self._cbb_LinkCadToSectionView.DisplayMember = 'Name'
		self._cbb_LinkCadToSectionView.AllowDrop = True
		self._cbb_LinkCadToSectionView.BackColor = System.Drawing.SystemColors.InactiveCaption
		self._cbb_LinkCadToSectionView.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._cbb_LinkCadToSectionView.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		self._cbb_LinkCadToSectionView.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._cbb_LinkCadToSectionView.ForeColor = System.Drawing.Color.Red
		self._cbb_LinkCadToSectionView.FormattingEnabled = True
		self._cbb_LinkCadToSectionView.Location = System.Drawing.Point(143, 186)
		self._cbb_LinkCadToSectionView.Name = "cbb_LinkCadToSectionView"
		self._cbb_LinkCadToSectionView.Size = System.Drawing.Size(337, 27)
		self._cbb_LinkCadToSectionView.Sorted = True
		self._cbb_LinkCadToSectionView.TabIndex = 0
		self._cbb_LinkCadToSectionView.SelectedIndexChanged += self.Cbb_LinkCadToSectionViewSelectedIndexChanged
		self._cbb_LinkCadToSectionView.Items.AddRange(System.Array[System.Object](allSectionsView))		
		self._cbb_LinkCadToSectionView.SelectedIndex = 0
		# 
		# lbl_Section
		# 
		self._lbl_Section.AutoSize = True
		self._lbl_Section.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lbl_Section.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lbl_Section.Location = System.Drawing.Point(12, 189)
		self._lbl_Section.Name = "lbl_Section"
		self._lbl_Section.Size = System.Drawing.Size(112, 21)
		self._lbl_Section.TabIndex = 1
		self._lbl_Section.Text = "Section View"
		self._lbl_Section.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# btt_PickPoint
		# 
		self._btt_PickPoint.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_PickPoint.ForeColor = System.Drawing.SystemColors.ControlText
		self._btt_PickPoint.Location = System.Drawing.Point(12, 106)
		self._btt_PickPoint.Name = "btt_PickPoint"
		self._btt_PickPoint.Size = System.Drawing.Size(110, 31)
		self._btt_PickPoint.TabIndex = 2
		self._btt_PickPoint.Text = "Pick Point"
		self._btt_PickPoint.UseVisualStyleBackColor = True
		self._btt_PickPoint.Click += self.Btt_PickPointClick
		# 
		# btt_SelectCadLink
		# 
		self._btt_SelectCadLink.AutoSize = True
		self._btt_SelectCadLink.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_SelectCadLink.Location = System.Drawing.Point(12, 27)
		self._btt_SelectCadLink.Name = "btt_SelectCadLink"
		self._btt_SelectCadLink.Size = System.Drawing.Size(110, 29)
		self._btt_SelectCadLink.TabIndex = 4
		self._btt_SelectCadLink.Text = "Cad Link"
		self._btt_SelectCadLink.UseVisualStyleBackColor = True
		self._btt_SelectCadLink.Click += self.Btt_SelectCadLinkClick
		# 
		# btt_RUN
		# 
		self._btt_RUN.AutoSize = True
		self._btt_RUN.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_RUN.Location = System.Drawing.Point(310, 280)
		self._btt_RUN.Name = "btt_RUN"
		self._btt_RUN.Size = System.Drawing.Size(75, 29)
		self._btt_RUN.TabIndex = 6
		self._btt_RUN.Text = "RUN"
		self._btt_RUN.UseVisualStyleBackColor = True
		self._btt_RUN.Click += self.Btt_RUNClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.AutoSize = True
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.Location = System.Drawing.Point(399, 280)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(81, 29)
		self._btt_CANCLE.TabIndex = 7
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# lbb_vitaminD
		# 
		self._lbb_vitaminD.AutoSize = True
		self._lbb_vitaminD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lbb_vitaminD.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lbb_vitaminD.Location = System.Drawing.Point(12, 294)
		self._lbb_vitaminD.Name = "lbb_vitaminD"
		self._lbb_vitaminD.Size = System.Drawing.Size(66, 15)
		self._lbb_vitaminD.TabIndex = 8
		self._lbb_vitaminD.Text = "@vitaminD"
		self._lbb_vitaminD.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		# 
		# txb_CadLink
		# 
		self._txb_CadLink.BackColor = System.Drawing.Color.Silver
		self._txb_CadLink.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_CadLink.ForeColor = System.Drawing.Color.Red
		self._txb_CadLink.Location = System.Drawing.Point(143, 109)
		self._txb_CadLink.Name = "txb_CadLink"
		self._txb_CadLink.Size = System.Drawing.Size(337, 27)
		self._txb_CadLink.TabIndex = 1
		self._txb_CadLink.TextChanged += self.Txb_CadLinkTextChanged
		# 
		# txb_PickPoint
		# 
		self._txb_PickPoint.BackColor = System.Drawing.Color.Silver
		self._txb_PickPoint.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_PickPoint.ForeColor = System.Drawing.Color.Red
		self._txb_PickPoint.Location = System.Drawing.Point(143, 29)
		self._txb_PickPoint.Name = "txb_PickPoint"
		self._txb_PickPoint.Size = System.Drawing.Size(337, 27)
		self._txb_PickPoint.TabIndex = 0
		self._txb_PickPoint.TextChanged += self.Txb_PickPointTextChanged
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(495, 325)
		self.Controls.Add(self._txb_PickPoint)
		self.Controls.Add(self._txb_CadLink)
		self.Controls.Add(self._lbb_vitaminD)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_RUN)
		self.Controls.Add(self._btt_SelectCadLink)
		self.Controls.Add(self._btt_PickPoint)
		self.Controls.Add(self._lbl_Section)
		self.Controls.Add(self._cbb_LinkCadToSectionView)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.Text = "Link Cad To Section View"
		self.TopMost = True
		self.TransparencyKey = System.Drawing.Color.White
		self.Load += self.MainFormLoad
		self.ResumeLayout(False)
		self.PerformLayout()


	def Btt_SelectCadLinkClick(self, sender, e):
		openDialog = OpenFileDialog()
		openDialog.Multiselect = False
		openDialog.Filter = "DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
		openDialog.ShowDialog()
		filePaths = openDialog.FileNames
		_sectionView = getAllSections(doc)[1]
		self.selSectionView = _sectionView			
		pass

	def Txb_CadLinkTextChanged(self, sender, e):
		pass

	def Btt_PickPointClick(self, sender, e):
		sectionViewPoint = pickPoints(doc)
		pass

	def Txb_PickedPointTextChanged(self, sender, e):
		pass

	def Cbb_LinkCadToSectionViewSelectedIndexChanged(self, sender, e):
		
		
		# self.selSectionView = _sectionView				
		pass

	def Btt_RUNClick(self, sender, e):

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
		doc.Link(filePaths[0], options, sectionView, linkedElem)
		TransactionManager.Instance.TransactionTaskDone()
		importInst = doc.GetElement(linkedElem.Value)
		TransactionManager.Instance.EnsureInTransaction(doc)
		importInst.Pinned = False
		TransactionManager.Instance.TransactionTaskDone()
		CADLink = doc.GetElement(importInst.GetTypeId())
		pointOrigin = XYZ(0,0,0)
		moveVector = sectionViewPoint - pointOrigin     
		TransactionManager.Instance.EnsureInTransaction(doc)
		elementTransform = Transform.CreateTranslation(moveVector)
		importInst.Location.Move(moveVector)
		TransactionManager.Instance.TransactionTaskDone()		
		pass

	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass


	def Txb_PickPointTextChanged(self, sender, e):
		pass

	def MainFormLoad(self, sender, e):
		pass
f = MainForm()
Application.Run(f)

