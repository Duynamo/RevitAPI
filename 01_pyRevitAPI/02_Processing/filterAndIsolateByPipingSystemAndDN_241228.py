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
'''___'''
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
'''___'''
#region _def
def getAllPipingSystemsInActiveView(doc):  
    # Collect all piping systems in the active view  
    collector = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PipingSystem)  
    pipingSystems = collector.ToElements()  
    # Create a set to store unique system names  
    pipingSystemsName = set()  
    for system in pipingSystems:  
        systemName = system.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()  
        if systemName:  # Check if systemName is not None  
            pipingSystemsName.add(systemName)  # Add system name to the set  
    return list(pipingSystemsName)  # Convert the set back to a list before returning
#endregion
'''___'''
allPipingSystemInActiveView = getAllPipingSystemsInActiveView(doc)
'''___'''
class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._grb_PipingSystem = System.Windows.Forms.GroupBox()
		self._lb_AllPipingSystemInActiveView = System.Windows.Forms.Label()
		self._clb_AllPipingSystemInActiveView = System.Windows.Forms.CheckedListBox()
		self._grb_Diameter = System.Windows.Forms.GroupBox()
		self._clb_Diameter = System.Windows.Forms.CheckedListBox()
		self._label1 = System.Windows.Forms.Label()
		self._ckb_AllNonePPSystem = System.Windows.Forms.CheckBox()
		self._ckb_AllNoneDiameter = System.Windows.Forms.CheckBox()
		self._btt_Isolate = System.Windows.Forms.Button()
		self._btt_Cancle = System.Windows.Forms.Button()
		self._lb_vitaminD = System.Windows.Forms.Label()
		self._grb_PipingSystem.SuspendLayout()
		self._grb_Diameter.SuspendLayout()
		self.SuspendLayout()
		# 
		# grb_PipingSystem
		# 
		self._grb_PipingSystem.Controls.Add(self._ckb_AllNonePPSystem)
		self._grb_PipingSystem.Controls.Add(self._clb_AllPipingSystemInActiveView)
		self._grb_PipingSystem.Controls.Add(self._lb_AllPipingSystemInActiveView)
		self._grb_PipingSystem.Location = System.Drawing.Point(12, 24)
		self._grb_PipingSystem.Name = "grb_PipingSystem"
		self._grb_PipingSystem.Size = System.Drawing.Size(400, 235)
		self._grb_PipingSystem.TabIndex = 0
		self._grb_PipingSystem.TabStop = False
		# 
		# lb_AllPipingSystemInActiveView
		# 
		self._lb_AllPipingSystemInActiveView.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_AllPipingSystemInActiveView.ForeColor = System.Drawing.SystemColors.MenuText
		self._lb_AllPipingSystemInActiveView.Location = System.Drawing.Point(6, 13)
		self._lb_AllPipingSystemInActiveView.Name = "lb_AllPipingSystemInActiveView"
		self._lb_AllPipingSystemInActiveView.Size = System.Drawing.Size(277, 22)
		self._lb_AllPipingSystemInActiveView.TabIndex = 0
		self._lb_AllPipingSystemInActiveView.Text = "All Piping System In Active View"
		# 
		# clb_AllPipingSystemInActiveView
		# 
		self._clb_AllPipingSystemInActiveView.Sorted = True
		self._clb_AllPipingSystemInActiveView.BackColor = System.Drawing.SystemColors.Control
		self._clb_AllPipingSystemInActiveView.CheckOnClick = True
		self._clb_AllPipingSystemInActiveView.Font = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_AllPipingSystemInActiveView.ForeColor = System.Drawing.Color.Black
		self._clb_AllPipingSystemInActiveView.FormattingEnabled = True
		self._clb_AllPipingSystemInActiveView.HorizontalScrollbar = True
		self._clb_AllPipingSystemInActiveView.Location = System.Drawing.Point(6, 47)
		self._clb_AllPipingSystemInActiveView.Name = "clb_AllPipingSystemInActiveView"
		self._clb_AllPipingSystemInActiveView.Size = System.Drawing.Size(387, 184)
		self._clb_AllPipingSystemInActiveView.TabIndex = 1
		self._clb_AllPipingSystemInActiveView.SelectedIndexChanged += self.Clb_AllPipingSystemInActiveViewSelectedIndexChanged
		self._clb_AllPipingSystemInActiveView.Items.AddRange(System.Array[System.Object](allPipingSystemInActiveView))		
		self._clb_AllPipingSystemInActiveView.SelectedIndex = 0	
		# 
		# grb_Diameter
		# 
		self._grb_Diameter.Controls.Add(self._ckb_AllNoneDiameter)
		self._grb_Diameter.Controls.Add(self._clb_Diameter)
		self._grb_Diameter.Controls.Add(self._label1)
		self._grb_Diameter.Location = System.Drawing.Point(12, 265)
		self._grb_Diameter.Name = "grb_Diameter"
		self._grb_Diameter.Size = System.Drawing.Size(228, 176)
		self._grb_Diameter.TabIndex = 1
		self._grb_Diameter.TabStop = False
		# 
		# clb_Diameter
		# 
		self._clb_Diameter.BackColor = System.Drawing.SystemColors.Control
		self._clb_Diameter.CheckOnClick = True
		self._clb_Diameter.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._clb_Diameter.ForeColor = System.Drawing.Color.Red
		self._clb_Diameter.FormattingEnabled = True
		self._clb_Diameter.HorizontalScrollbar = True
		self._clb_Diameter.Location = System.Drawing.Point(6, 34)
		self._clb_Diameter.Name = "clb_Diameter"
		self._clb_Diameter.Size = System.Drawing.Size(213, 130)
		self._clb_Diameter.TabIndex = 5
		self._clb_Diameter.SelectedIndexChanged += self.Clb_DiameterSelectedIndexChanged
		# 
		# label1
		# 
		self._label1.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._label1.ForeColor = System.Drawing.SystemColors.MenuText
		self._label1.Location = System.Drawing.Point(6, 9)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(89, 22)
		self._label1.TabIndex = 4
		self._label1.Text = "Diameter"
		# 
		# ckb_AllNonePPSystem
		# 
		self._ckb_AllNonePPSystem.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._ckb_AllNonePPSystem.Location = System.Drawing.Point(289, 13)
		self._ckb_AllNonePPSystem.Name = "ckb_AllNonePPSystem"
		self._ckb_AllNonePPSystem.Size = System.Drawing.Size(104, 24)
		self._ckb_AllNonePPSystem.TabIndex = 2
		self._ckb_AllNonePPSystem.Text = "All/None"
		self._ckb_AllNonePPSystem.UseVisualStyleBackColor = True
		self._ckb_AllNonePPSystem.CheckedChanged += self.Ckb_AllNonePPSystemCheckedChanged
		# 
		# ckb_AllNoneDiameter
		# 
		self._ckb_AllNoneDiameter.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._ckb_AllNoneDiameter.Location = System.Drawing.Point(116, 8)
		self._ckb_AllNoneDiameter.Name = "ckb_AllNoneDiameter"
		self._ckb_AllNoneDiameter.Size = System.Drawing.Size(104, 24)
		self._ckb_AllNoneDiameter.TabIndex = 3
		self._ckb_AllNoneDiameter.Text = "All/None"
		self._ckb_AllNoneDiameter.UseVisualStyleBackColor = True
		self._ckb_AllNoneDiameter.CheckedChanged += self.Ckb_AllNoneDiameterCheckedChanged
		# 
		# btt_Isolate
		# 
		self._btt_Isolate.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_Isolate.ForeColor = System.Drawing.Color.Red
		self._btt_Isolate.Location = System.Drawing.Point(257, 299)
		self._btt_Isolate.Name = "btt_Isolate"
		self._btt_Isolate.Size = System.Drawing.Size(93, 27)
		self._btt_Isolate.TabIndex = 2
		self._btt_Isolate.Text = "ISOLATE"
		self._btt_Isolate.UseVisualStyleBackColor = True
		self._btt_Isolate.Click += self.Btt_IsolateClick
		# 
		# btt_Cancle
		# 
		self._btt_Cancle.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_Cancle.ForeColor = System.Drawing.Color.Red
		self._btt_Cancle.Location = System.Drawing.Point(320, 349)
		self._btt_Cancle.Name = "btt_Cancle"
		self._btt_Cancle.Size = System.Drawing.Size(92, 27)
		self._btt_Cancle.TabIndex = 3
		self._btt_Cancle.Text = "CANCLE"
		self._btt_Cancle.UseVisualStyleBackColor = True
		self._btt_Cancle.Click += self.Btt_CancleClick
		# 
		# lb_vitaminD
		# 
		self._lb_vitaminD.AutoSize = True
		self._lb_vitaminD.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_vitaminD.Location = System.Drawing.Point(376, 424)
		self._lb_vitaminD.Name = "lb_vitaminD"
		self._lb_vitaminD.Size = System.Drawing.Size(31, 17)
		self._lb_vitaminD.TabIndex = 4
		self._lb_vitaminD.Text = "@D"
		self._lb_vitaminD.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(419, 453)
		self.Controls.Add(self._lb_vitaminD)
		self.Controls.Add(self._btt_Cancle)
		self.Controls.Add(self._grb_Diameter)
		self.Controls.Add(self._grb_PipingSystem)
		self.Controls.Add(self._btt_Isolate)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.Text = "IsolateByPSysAndDN"
		self.TopMost = True
		self._grb_PipingSystem.ResumeLayout(False)
		self._grb_Diameter.ResumeLayout(False)
		self.ResumeLayout(False)
		self.PerformLayout()



	def Clb_AllPipingSystemInActiveViewSelectedIndexChanged(self, sender, e):
		pass

	def Ckb_AllNonePPSystemCheckedChanged(self, sender, e):
		pass

	def Clb_DiameterSelectedIndexChanged(self, sender, e):
		pass

	def Ckb_AllNoneDiameterCheckedChanged(self, sender, e):
		pass

	def Btt_IsolateClick(self, sender, e):
		
		pass

	def Btt_CancleClick(self, sender, e):
		self.Close()
		pass
	
f = MainForm()
Application.Run(f)