"""Copyright by: vudinhduybm@gmail.com"""
import clr 
import sys 
import System   
clr.AddReference("ProtoGeometry")
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('System')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
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

from System.Windows.Forms import OpenFileDialog

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

#Region __def
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
	msg = "Pick Points on Current Section plane, hit ESC when finished."
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
sectionViewPoint = pickPoints(doc)
openDialog = OpenFileDialog()
openDialog.Multiselect = False
openDialog.Filter = "DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
openDialog.ShowDialog()
# fileName = openDialog.FileNames
'''___'''
filePaths = openDialog.FileNames
sectionView = getAllSections(doc)[1]
# Import options
options = DWGImportOptions()
options.AutoCorrectAlmostVHLines = True
options.CustomScale = 1
options.OrientToView = True
options.ThisViewOnly = False
options.VisibleLayersOnly = False
options.ColorMode = ImportColorMode.Preserved
options.Placement = ImportPlacement.Origin
options.Unit = ImportUnit.Millimeter

# Create ElementId / .NET object
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
# Apply the transformation
TransactionManager.Instance.EnsureInTransaction(doc)
elementTransform = Transform.CreateTranslation(moveVector)
importInst.Location.Move(moveVector)
TransactionManager.Instance.TransactionTaskDone()
'''___'''

OUT = moveVector


