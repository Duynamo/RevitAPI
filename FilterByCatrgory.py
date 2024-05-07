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

"""_________"""

cateList = [
    "StructuralColumns",
    "StructuralFraming",
    "Doors",
    "PipeCurves",
    "DuctCurves",
    "PipeAccessory",
    "DuctAccessory",
    "Roofs",
    "Floors"
]
"""___"""
class selectionFilter(ISelectionFilter):
    def __init__(self, category):
        self.category = category
    def AllowElement(self, element):
        if element.Category.Name in [c.Name for c in self.category]:
            return True
        else:
            return False
    def AllowReference(reference, point):
        return False
    
"""___"""

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._clb_Categories = System.Windows.Forms.CheckedListBox()
		self._lb_Categories = System.Windows.Forms.Label()
		self._lb_FVC = System.Windows.Forms.Label()
		self._btt_SelectElement = System.Windows.Forms.Button()
		self._btt_Run = System.Windows.Forms.Button()
		self.SuspendLayout()
		# 
		# clb_Categories
		# 
		self._clb_Categories.DisplayMember = "Name"
		self._clb_Categories.AllowDrop = True
		self._clb_Categories.BackColor = System.Drawing.SystemColors.Window
		self._clb_Categories.CheckOnClick = True
		self._clb_Categories.FormattingEnabled = True
		self._clb_Categories.HorizontalScrollbar = True
		self._clb_Categories.Location = System.Drawing.Point(119, 34)
		self._clb_Categories.Name = "clb_Categories"
		self._clb_Categories.Size = System.Drawing.Size(236, 106)
		self._clb_Categories.Sorted = True
		self._clb_Categories.TabIndex = 0
		self._clb_Categories.SelectedIndexChanged += self.Clb_CategoriesSelectedIndexChanged
		self._clb_Categories.Items.AddRange(System.Array[System.Object](cateList))
		# 
		# lb_Categories
		# 
		self._lb_Categories.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_Categories.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_Categories.ForeColor = System.Drawing.Color.Blue
		self._lb_Categories.Location = System.Drawing.Point(2, 70)
		self._lb_Categories.Name = "lb_Categories"
		self._lb_Categories.Size = System.Drawing.Size(111, 26)
		self._lb_Categories.TabIndex = 1
		self._lb_Categories.Text = "Categories"
		# 
		# lb_FVC
		# 
		self._lb_FVC.Font = System.Drawing.Font("Meiryo UI", 4.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_FVC.Location = System.Drawing.Point(12, 272)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(69, 14)
		self._lb_FVC.TabIndex = 2
		self._lb_FVC.Text = "@FVC"
		# 
		# btt_SelectElement
		# 
		self._btt_SelectElement.BackColor = System.Drawing.Color.FromArgb(255, 192, 192)
		self._btt_SelectElement.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_SelectElement.FlatAppearance.BorderColor = System.Drawing.Color.Blue
		self._btt_SelectElement.FlatAppearance.BorderSize = 5
		self._btt_SelectElement.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(192, 192, 255)
		self._btt_SelectElement.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(128, 255, 128)
		self._btt_SelectElement.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_SelectElement.Location = System.Drawing.Point(179, 245)
		self._btt_SelectElement.Name = "btt_SelectElement"
		self._btt_SelectElement.Size = System.Drawing.Size(85, 38)
		self._btt_SelectElement.TabIndex = 4
		self._btt_SelectElement.Text = "Select"
		self._btt_SelectElement.UseVisualStyleBackColor = False
		self._btt_SelectElement.Click += self.Btt_SelectElementClick
		# 
		# btt_Run
		# 
		self._btt_Run.BackColor = System.Drawing.Color.FromArgb(255, 192, 192)
		self._btt_Run.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_Run.FlatAppearance.BorderColor = System.Drawing.Color.Blue
		self._btt_Run.FlatAppearance.BorderSize = 5
		self._btt_Run.FlatAppearance.MouseDownBackColor = System.Drawing.Color.FromArgb(128, 128, 255)
		self._btt_Run.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(128, 255, 128)
		self._btt_Run.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
		self._btt_Run.Location = System.Drawing.Point(270, 245)
		self._btt_Run.Name = "btt_Run"
		self._btt_Run.Size = System.Drawing.Size(85, 38)
		self._btt_Run.TabIndex = 5
		self._btt_Run.Text = "CANCLE"
		self._btt_Run.UseVisualStyleBackColor = False
		self._btt_Run.Click += self.Btt_RunClick
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(367, 295)
		self.Controls.Add(self._btt_Run)
		self.Controls.Add(self._btt_SelectElement)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._lb_Categories)
		self.Controls.Add(self._clb_Categories)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.Text = "SelectionByCategory"
		self.TopMost = True
		self.ResumeLayout(False)


	def Clb_CategoriesSelectedIndexChanged(self, sender, e):
		pass

	def Btt_SelectElementClick(self, sender, e):
		selectedCatName = [item for item in self._clb_Categories.CheckedItems]
		
		modelCate = []
		categories = doc.Settings.Categories
		for c in categories:
			if c.CategoryType == CategoryType.Model:
				category_name = c.Name
				if category_name in selectedCatName:
					modelCate.append(c)
		selFilter = selectionFilter(modelCate)
		ele = uidoc.Selection.PickElementsByRectangle(selFilter, 'please drag to select elements')
		IDS = List[ElementId]()
		for e in ele:
			IDS.Add(e.Id)
		TransactionManager.Instance.EnsureInTransaction(doc)
		isolatedEles = view.IsolateElementsTemporary(IDS)
		TransactionManager.Instance.TransactionTaskDone()
		pass

	def Btt_RunClick(self, sender, e):
            

		
		self.Close()
		pass
	

f = MainForm()
Application.Run(f)