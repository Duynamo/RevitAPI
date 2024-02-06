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

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._btt_PickPoints = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._btt_OK = System.Windows.Forms.Button()
		self._btt_Markup = System.Windows.Forms.Button()
		self._clb_pickedPoints = System.Windows.Forms.CheckedListBox()
		self._ckb_AllNone = System.Windows.Forms.CheckBox()
		self.SuspendLayout()
		# 
		# btt_PickPoints
		# 
		self._btt_PickPoints.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_PickPoints.FlatAppearance.BorderSize = 5
		self._btt_PickPoints.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue
		self._btt_PickPoints.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow
		self._btt_PickPoints.Location = System.Drawing.Point(12, 12)
		self._btt_PickPoints.Name = "btt_PickPoints"
		self._btt_PickPoints.Size = System.Drawing.Size(95, 46)
		self._btt_PickPoints.TabIndex = 0
		self._btt_PickPoints.Text = "Pick Points"
		self._btt_PickPoints.UseVisualStyleBackColor = True
		self._btt_PickPoints.Click += self.Btt_PickPointsClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_CANCLE.FlatAppearance.BorderSize = 5
		self._btt_CANCLE.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue
		self._btt_CANCLE.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow
		self._btt_CANCLE.Location = System.Drawing.Point(244, 290)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(112, 55)
		self._btt_CANCLE.TabIndex = 0
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# btt_OK
		# 
		self._btt_OK.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_OK.FlatAppearance.BorderSize = 5
		self._btt_OK.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue
		self._btt_OK.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow
		self._btt_OK.Location = System.Drawing.Point(126, 290)
		self._btt_OK.Name = "btt_OK"
		self._btt_OK.Size = System.Drawing.Size(112, 55)
		self._btt_OK.TabIndex = 0
		self._btt_OK.Text = "OK"
		self._btt_OK.UseVisualStyleBackColor = True
		self._btt_OK.Click += self.Btt_OKClick
		# 
		# btt_Markup
		# 
		self._btt_Markup.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_Markup.FlatAppearance.BorderSize = 5
		self._btt_Markup.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue
		self._btt_Markup.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow
		self._btt_Markup.Location = System.Drawing.Point(15, 290)
		self._btt_Markup.Name = "btt_Markup"
		self._btt_Markup.Size = System.Drawing.Size(105, 55)
		self._btt_Markup.TabIndex = 1
		self._btt_Markup.Text = "Markup"
		self._btt_Markup.UseVisualStyleBackColor = True
		self._btt_Markup.Click += self.Btt_MarkupClick
		# 
		# clb_pickedPoints
		# 
		self._clb_pickedPoints.AllowDrop = True
		self._clb_pickedPoints.CheckOnClick = True
		self._clb_pickedPoints.Font = System.Drawing.Font("MS UI Gothic", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_pickedPoints.FormattingEnabled = True
		self._clb_pickedPoints.HorizontalScrollbar = True
		self._clb_pickedPoints.Location = System.Drawing.Point(170, 13)
		self._clb_pickedPoints.Name = "clb_pickedPoints"
		self._clb_pickedPoints.Size = System.Drawing.Size(195, 157)
		self._clb_pickedPoints.TabIndex = 2
		self._clb_pickedPoints.SelectedIndexChanged += self.Clb_pickedPointsSelectedIndexChanged
		# 
		# ckb_AllNone
		# 
		self._ckb_AllNone.Location = System.Drawing.Point(271, 176)
		self._ckb_AllNone.Name = "ckb_AllNone"
		self._ckb_AllNone.Size = System.Drawing.Size(94, 24)
		self._ckb_AllNone.TabIndex = 4
		self._ckb_AllNone.Text = "All/None"
		self._ckb_AllNone.UseVisualStyleBackColor = True
		self._ckb_AllNone.CheckedChanged += self.Ckb_AllNoneCheckedChanged
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(377, 404)
		self.Controls.Add(self._ckb_AllNone)
		self.Controls.Add(self._clb_pickedPoints)
		self.Controls.Add(self._btt_Markup)
		self.Controls.Add(self._btt_OK)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_PickPoints)
		self.Name = "MainForm"
		self.Text = "MarkupPickedPoints"
		self.TopMost = True
		self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.ResumeLayout(False)


	def Btt_PickPointsClick(self, sender, e):
		TransactionManager.Instance.EnsureInTransaction(doc)
		activeView = doc.ActiveView
		iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
		sketchPlane = SketchPlane.Create(doc, iRefPlane)
		doc.ActiveView.SketchPlane = sketchPlane
		condition = True
		points = []
		n = 0
		msg = "Pick Points on Current Section plane, hit ESC when finished."
		# TaskDialog.Show("^------^", msg)
		while condition:
			try:
				# logger('Line383:', n)
				pt=uidoc.Selection.PickPoint()
				n += 1
				points.append(pt)				
			except :
				condition = False
		doc.Delete(sketchPlane.Id)	
		for j in points:
			rpM = Autodesk.DesignScript.Geometry.Point.ByCoordinates(j.X*304.8, j.Y*304.8, j.Z*304.8)
			self._clb_pickedPoints.Items.Add(rpM)
		self._ckb_AllNone.Checked = True	
		TransactionManager.Instance.TransactionTaskDone()					

		pass

	def Btt_MarkupClick(self, sender, e):
		txtLocation = self._clb_pickedPoints.CheckedItems
		defaultTextTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
		noteWidth = 0.2
		minWidth = TextNote.GetMinimumAllowedWidth(doc, defaultTextTypeId)
		maxWidth = TextNote.GetMaximumAllowedWidth(doc, defaultTextTypeId)
		if noteWidth < minWidth :
			noteWidth = minWidth
		else :
			noteWidth = maxWidth	
		textNoteOpts = TextNoteOptions(defaultTextTypeId)
		textNoteOpts.HorizontalAlignment.Center
		
		textNoteType = None
		textNoteTypes = FilteredElementCollector(doc).OfClass(TextNoteType).ToElements()
		desiredTextSize = 200  # Set your desired text size here
		for tnt in textNoteTypes:
			if tnt.get_Parameter(BuiltInParameter.TEXT_SIZE).AsDouble() == desiredTextSize:
				textNoteType = tnt
				break
		TransactionManager.Instance.EnsureInTransaction(doc)
		if textNoteTypes is None:
            # If the desired TextNoteType doesn't exist, create a new one
			newType = TextNoteType.Create(doc, "Custom Text Note Type", desiredTextSize)
			textNoteType = newType.Id
			textNoteOpts.set_TypeId(textNoteType)
		TransactionManager.Instance.TransactionTaskDone()

		for index, point in enumerate (txtLocation):
			XYZ_point = XYZ((point.X)/304.8, (point.Y)/304.8, (point.Z)/304.8)
			txt = "Point " + str( index + 1)
			textNote = TextNote.Create(doc, view.Id, XYZ_point, noteWidth, txt, textNoteOpts)
          
		existedTextNoteTypes = FilteredElementCollector(doc).OfClass(TextNoteType).ToElements()
		desTextNoteType = existedTextNoteTypes[0]
		TransactionManager.Instance.EnsureInTransaction(doc)
		desTextNoteType.get_Parameter(BuiltInParameter.TEXT_SIZE).Set(20/ 304.8)  
		desTextNoteType.get_Parameter(BuiltInParameter.LINE_COLOR).Set(255)  
		TransactionManager.Instance.TransactionTaskDone()
		pass		
	
	
	def Ckb_AllNoneCheckedChanged(self, sender, e):
		var = self._clb_pickedPoints.Items.Count
		rangers = range(var)
		for i in rangers:
			if self._ckb_AllNone.Checked == True:
				self._clb_pickedPoints.SetItemChecked(i, True)
				# self._total_XYValue.Text = str(var)
			else:
				self._clb_pickedPoints.SetItemChecked(i, False)
				# self._total_XYValue.Text = str(0)		
		pass

	def Clb_pickedPointsSelectedIndexChanged(self, sender, e):
		pass

	def Btt_OKClick(self, sender, e):
		pass

	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass
		
f = MainForm()
Application.Run(f)	