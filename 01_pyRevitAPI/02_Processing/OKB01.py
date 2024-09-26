import clr
import sys 
import System   
import math
import os

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB import SaveAsOptions
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
#region __function
cateList = [
    "Pipe Fittings"
    ]

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

#endregion 

#region __code here
categories = [BuiltInCategory.OST_PipeFitting]
#region __pipe list
susPipes = []
susPipes_visible_1end = []
susPipes_visible_1mid = []
susPipes_visible_1end1mid = []
susPipes_visible_2end = []
susPipes_visible_2end1mid = []
susPipes_visible_0end0mid = []
#endregion
#region __sus pipes parameters for BOM
_FVC_PartName = None
_FVC_PartNumber = None
_FVC_PartMaterial = None
_FVC_PartDimension = None
_FVC_Note= None
_FVC_CoatingSpec= None
_FVC_F_M1 = None
_FVC_F_M2 = None
_FVC_F_B = None

#endregion
#region __sus pipes parameters other
_FVC_PartType = None
_FVC_Diameter = None
#endregion

for c in categories:
    collector = FilteredElementCollector(doc, view.Id).OfCategory(c).WhereElementIsNotElementType().ToElements()
for ele in collector:
    typeId = ele.GetTypeId()
    if typeId != ElementId.InvalidElementId:  
        elementType = doc.GetElement(typeId)
        if elementType is not None:
            elementName = elementType.FamilyName 
            if elementName and elementName == 'SUS_PIPE_DN':  
                susPipes.append(ele)
    
for p in susPipes:
    checkVisible1 = p.LookupParameter('FVC_Visible1').AsValueString()
    checkVisible2 = p.LookupParameter('FVC_Visible2').AsValueString()
    checkVisible3 = p.LookupParameter('FVC_Visible3').AsValueString()
    _FVC_PartName_param = p.LookupParameter('FVC_Part Name')
    _FVC_PartNumber = p.LookupParameter('FVC_Part Number')
    _FVC_PartMaterial = p.LookupParameter('FVC_Part Material')
    # _FVC_PartStandard = p.LookupParameter('FVC_Part Standard')
    _FVC_Note = p.LookupParameter('FVC_Note')
    _FVC_CoatingSpec = p.LookupParameter('FVC_CoatingSpec')
    _FVC_F_M1 = p.LookupParameter('FVC_F_M1')
    _FVC_F_M2 = p.LookupParameter('FVC_F_M2')
    _FVC_F_B = p.LookupParameter('FVC_F_B')
    _FVC_PartStandard = p.LookupParameter('FVC_Part Standard')
    susPipe_DN_param = p.LookupParameter('DN').AsValueString()
    susPipe_t_param = p.LookupParameter('t').AsValueString()
    TransactionManager.Instance.EnsureInTransaction(doc)
    _FVC_PartStandard.Set('D'+str( susPipe_DN_param)+' x '+str( susPipe_t_param)+'mm')
    _FVC_PartMaterial.Set('vitaminD')
    TransactionManager.Instance.TransactionTaskDone()
#region __group sus pipe by visible status  
    #check susPipes_visible_1end
    if checkVisible1 == 'Yes' and checkVisible2 == 'No' and checkVisible3 == 'No' :
        susPipes_visible_1end.append(p)
    elif checkVisible1 == 'No' and checkVisible2 == 'Yes' and checkVisible3 == 'No' :
        susPipes_visible_1end.append(p)      
    #check susPipes_visible_2end
    elif checkVisible1 == 'Yes' and checkVisible2 == 'Yes' and checkVisible3 == 'No' :
        susPipes_visible_1mid.append(p)
    #check susPipes_visible_1end1mid    
    elif checkVisible1 == 'Yes' and checkVisible2 == 'No' and checkVisible3 == 'Yes' :
        susPipes_visible_1end1mid.append(p)
    elif checkVisible1 == 'No' and checkVisible2 == 'Yes' and checkVisible3 == 'Yes' :
        susPipes_visible_1end1mid.append(p)
     #check susPipes_visible_2end1mid   
    elif checkVisible1 == 'Yes' and checkVisible2 == 'Yes' and checkVisible3 == 'Yes' :
        susPipes_visible_2end1mid.append(p)
    #check susPipes_visible_0end0mid    
    elif checkVisible1 == 'No' and checkVisible2 == 'No' and checkVisible3 == 'No' :
        susPipes_visible_0end0mid.append(p)
#endregion

# IDS = List[ElementId]()
# for i in flat_categoriesFilter:
# 	IDS.Add(i.Id)



#endregion


OUT = susPipes, susPipes_visible_1end ,susPipes_visible_1mid,susPipes_visible_1end1mid,susPipes_visible_2end,susPipes_visible_2end1mid,susPipes_visible_0end0mid