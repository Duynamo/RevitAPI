import clr
import sys 
import System   


clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB import MEPSystemType
# from Autodesk.Revit.DB import MechanicalSystem


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

def getAllPipingSystems(doc):
	collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem)
	pipingSystems = collector.ToElements()
	pipingSystemsName = []
	for system in pipingSystems:
		systemName = system.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
		pipingSystemsName.append(systemName)
	return pipingSystems

pipingSystems = getAllPipingSystems(doc)
# pipingSystemsCollector = [item for item in lst if item is not None]

def getAllPipeTypes(doc):
	collector1 = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves)
	pipeTypes = collector1.ToElements()
	# pipeTypesName = []
	# for pipeType in pipeTypes:
	# 	pipeTypeName = pipeType.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
	# 	pipeTypesName.append(pipeTypeName)
	return pipeTypes

lst1 = getAllPipeTypes(doc)
# pipeTypesCollector = list(set(lst_PipeType))
# pipeTypesCollector = [item for item in lst1 if item is not None]
pipeTypesCollector = []
[pipeTypesCollector.append(i) for i in lst1 if i not in pipeTypesCollector]
# for i in lst1:
# 	if i not in pipeTypesCollector:
# 		pipeTypesCollector.append(i)


def create_piping_system(doc, system_name):
    # Get the Piping System Type
    system_type = FilteredElementCollector(doc).OfClass(MEPSystemType).OfCategory(BuiltInCategory.OST_PipingSystem).FirstElement()

    # Check if the system type is found
    if system_type is not None:
        # Create a Mechanical System
        new_system = MechanicalSystem.Create(doc, system_type.Id)

        # Set the name of the system
        new_system.Name = system_name

        # Commit the transaction
        doc.Regenerate()
        return new_system
    return None

newSystemName1 = "vudinhduy"
newSystemName2 = "dadasdadadad"
newSystemName = []
[newSystemName.extend(i for i in (newSystemName1, newSystemName2))]
newSystemType = create_piping_system(newSystemName)



class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._cbb = System.Windows.Forms.ComboBox()
		self.SuspendLayout()
		# 
		# cbb
		# 
		self._cbb.DisplayMember = "Name"
		self._cbb.FormattingEnabled = True
		self._cbb.Location = System.Drawing.Point(57, 65)
		self._cbb.Name = "cbb"
		self._cbb.Size = System.Drawing.Size(397, 23)
		self._cbb.TabIndex = 0
		self._cbb.SelectedIndexChanged += self.CbbSelectedIndexChanged
		self._cbb.Items.AddRange(System.Array[System.Object](newSystemType)) 
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(554, 320)
		self.Controls.Add(self._cbb)
		self.Name = "MainForm"
		self.Text = "test"
		self.ResumeLayout(False)


	def CbbSelectedIndexChanged(self, sender, e):
		pass



f = MainForm()
Application.Run(f)

# lstt = [1,1,1,1,2,34,5]
# als = []
# [als.append(i) for i in lstt if i not in als]
# print(als)