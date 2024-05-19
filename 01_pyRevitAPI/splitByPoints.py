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
def toList(input):
      if isinstance(input, list):
            return UnwrapElement(input)
      else:
            return [UnwrapElement(input)]

def pointOnLine(p,a,b):
	if not (p-a).CrossProduct(b-a).IsAlmostEqualTo(XYZ.Zero):
		return False
	if a.X!=b.X:
		if a.X<p.X<b.X:
			return True
		if a.X>p.X>b.X:
			return True
	else:
		if a.Y<p.Y<b.Y:
			return True
		if a.Y>p.Y>b.Y:
			return True
	return False
"""_____"""
elements = toList(IN[0])
points = toList(IN[1])
points = [p.ToXyzs() for p in points]
parents,childs,fittings = [],[],[]
children = {}
"""_____"""
TransactionManager.Instance.EnsureInTransaction(doc)
for e, p in zip(elements,points):
	to_check = [e]
	if e.Id in children:
		to_check.extend(children[e.Id])
		
	splitId = None
	for ec in to_check:
		if isinstance(ec,Autodesk.Revit.DB.Plumbing.Pipe):
			try:
				splitId = Autodesk.Revit.DB.Plumbing.PlumbingUtils.BreakCurve(doc, ec.Id, p)
				break
			except:
				pass				
		elif isinstance(ec,Autodesk.Revit.DB.Mechanical.Duct):
			try:
				splitId = Autodesk.Revit.DB.Mechanical.MechanicalUtils.BreakCurve(doc, ec.Id, p)
				break
			except:
				pass

	if splitId:
		split = doc.GetElement(splitId)
		if hasattr(split,"ConnectorManager"):
			newPipeConnectors = split.ConnectorManager.Connectors
			#Check
			connA = None
			connB = None
			for c in ec.ConnectorManager.Connectors:
				pc = c.Origin
				#get connectorB near to connectorA
				nearest = [x for x in newPipeConnectors if pc.DistanceTo(x.Origin) < 0.01]
				#If exists assign
				if nearest:
					connA = c
					connB = nearest[0]
			#Create fitting
			try:
				fittings.append(doc.Create.NewUnionFitting(connA, connB))
			except:pass
		
		if e.Id in children:
			children[e.Id].append(split)
		else:
			children[e.Id] = [split]
		parents.append(ec)
		childs.append(split)
	else:
		parents.append(None)
		childs.append(None)
TransactionManager.Instance.TransactionTaskDone()
			
OUT = parents,childs,fittings