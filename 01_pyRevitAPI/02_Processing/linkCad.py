#Based on a script by Konrad Sobon
#Additions by Alban de Chasteigner 2018

import clr
# Import DocumentManager and TransactionManager
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

# Import RevitAPI
clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

filePaths = IN[0] if isinstance(IN[0],list) else [IN[0]]
views = UnwrapElement(IN[1]) if isinstance(IN[1],list) else [UnwrapElement(IN[1])]
customscale = IN[2]
colormode = IN[3]
placement = IN[4]
unit = IN[5]
allView = IN[6]
viewsplaced,outName,CADlinktype,importinstance = [],[],[],[]

options = DWGImportOptions()
options.AutoCorrectAlmostVHLines = True
options.OrientToView = True
if allView :options.ThisViewOnly = False
else: options.ThisViewOnly = True
options.VisibleLayersOnly = True
options.CustomScale = customscale
if colormode == None: ImportColorMode.Preserved
else: options.ColorMode = colormode
if placement == None: ImportPlacement.Shared
else :options.Placement= placement
if unit == None : ImportUnit.Default
else : options.Unit = unit

# Create ElementId / .NET object
linkedElem = clr.Reference[ElementId]()

for view in range(len(views)):
	TransactionManager.Instance.EnsureInTransaction(doc)
	doc.Link(filePaths[view], options, views[view], linkedElem)
	TransactionManager.Instance.TransactionTaskDone()
	viewsplaced.append(views[view])
	importinst = doc.GetElement(linkedElem.Value)
	importinstance.append(importinst)
	CADLink = doc.GetElement(importinst.GetTypeId())
	CADlinktype.append(CADLink)
	outName.append(Element.Name.GetValue(CADLink))
	
if isinstance(IN[0], list): OUT = viewsplaced,outName,CADlinktype,importinstance
else: OUT = viewsplaced[0],outName[0],CADlinktype[0],importinstance[0]