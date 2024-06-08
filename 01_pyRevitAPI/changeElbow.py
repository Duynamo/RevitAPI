"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

from math import cos,sin,tan,radians

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

def get_elbow_fitting_between_pipes(doc, pipeLst):
	conns = []
	ref = []
	fittingsLst = []
	for p in pipeLst:
		connLst = []
		refs = []
		fittings = []
		try:
			connectors = p.MEPModel.ConnectorManager.Connectors
		except:
			try:
				connectors = p.ConnectorManager.Connectors
			except:
				conns.append(None)
		for conn in connectors:
			connLst.append(conn)
			for c in conn.AllRefs:
				refs.append(c.Owner)
				for item in refs:
					IDS = List[ElementId]()
					IDS.Add(item.Id)
					fittings = FilteredElementCollector(doc, IDS).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()
		conns.append(connLst)
		ref.append(refs)
		fittingsLst.append(fittings)
	return conns, ref , fittingsLst