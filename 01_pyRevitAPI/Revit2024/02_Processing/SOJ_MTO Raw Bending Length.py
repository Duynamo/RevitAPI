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
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
'''___'''

def getAllRawBendingPipeInActiveView():
    desPipes = []
    od = []
    collector = FilteredElementCollector(doc, view.Id).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
    for ele in collector:
        eleCheckParam = ele.LookupParameter('FU_element Name').AsString()
        if eleCheckParam!= None and eleCheckParam == 'Raw Bending  HDPE':
            desPipes.append(ele)
            ele_Volume = ele.LookupParameter('Volume').AsDouble()
            pipeOD = ele.LookupParameter('FU_RawBending_OD').AsDouble()
            od.append(pipeOD)
            rawBendingPipeLength = (ele_Volume ) / pow((pipeOD*0.5),2)/3.14 #convert to millimeters
            ele_MTO_Length_param = ele.LookupParameter('FU_MTO Length')
            TransactionManager.Instance.EnsureInTransaction(doc)
            ele_MTO_Length_param.Set(rawBendingPipeLength)
            TransactionManager.Instance.TransactionTaskDone()
        else: pass
    return desPipes, od


OUT =  getAllRawBendingPipeInActiveView()

'''___'''