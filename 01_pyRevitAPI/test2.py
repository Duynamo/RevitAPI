"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

from math import *

sys.path.append('C:\Program Files\Autodesk\Revit 2022\AddIns\DynamoForRevit\IronPython.StdLib.2.7.9')
from duynamoLibrary import *

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*

clr.AddReference("DSCoreNodes")
from DSCore.List import Flatten

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
#endregion

#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

#region ___someFunctions
def uwList(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)
def allPipesInActiveView():
	pipesCollector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
	return pipesCollector
def getAllConnectedParts(pipes):
	changedPipes =[]
	for p in pipes:
		conns = list(p.ConnectorManager.Connectors.GetEnumerator())
		lk_CELengthParam = p.LookupParameter('CE_Length')
		for c in conns:
			connectedItems = []
			lk_connsLengthParams = []
			sum_connsLength =[]
			if c.IsConnected:
				for ref in c.AllRefs:
					if ref.Owner.Id != p.Id:
						connectedItems.append(ref.Owner)
						for item in connectedItems:
							lk_connsLengthParam = item.LookupParameter('Cons_Length').AsDouble()*304.8
							lk_connsLengthParams.append(lk_connsLengthParam)
						sum_connsLength = sum(lk_connsLengthParams)


	return changedPipes
#endregion


OUT = allPipesInActiveView()