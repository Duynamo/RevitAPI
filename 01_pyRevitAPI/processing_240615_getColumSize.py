"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections
import duynamoLibrary as dLib

sys.path.append('')
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
def duplicate_and_resize_concrete_columns(doc, dimensions_list):
    # Collect all structural column types
    colCollector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralColumns).WhereElementIsElementType().ToElements()
    
    # Initialize a list to store the new column types
    new_column_types = []

    TransactionManager.Instance.EnsureInTransaction(doc)
    
    try:
        for c in colCollector:
            # Get the material parameter
            matParam1 = c.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
            
            if matParam1:
                matName = matParam1.AsValueString()
                # Check if the material name contains "Concrete"
                if matName and 'Concrete' in matName:
                    # Loop through the list of dimensions
                    for dims in dimensions_list:
                        b, h = dims
                        # Duplicate the column type
                        new_col_type_id = c.Duplicate(f"{c.Name} {b}x{h}")
                        new_col_type = doc.GetElement(new_col_type_id)
                        
                        # Set the new dimensions
                        b_param = new_col_type.LookupParameter('b')
                        h_param = new_col_type.LookupParameter('h')
                        
                        if b_param and h_param:
                            b_param.Set(b)
                            h_param.Set(h)
                            
                            # Add to the new column types list
                            new_column_types.append(new_col_type)
    except Exception as e:
        pass

    TransactionManager.Instance.TransactionTaskDone()

    return new_column_types
#endregion


eleList   = uwList(IN[0])

# Do some action in a Transaction
TransactionManager.Instance.EnsureInTransaction(doc)
cate = BuiltInCategory.OST_StructuralColumns
colCollector = FilteredElementCollector(doc).OfCategory(cate).WhereElementIsElementType()
filteredColumns = []
name = []

for c in colCollector:
    name = c.Family.LookupParameter('Material for Model Behavior').AsValueString()
    if name is not None and 'Concrete' in name :
        filteredColumns.append(c)

baseCol = filteredColumns[0]

TransactionManager.Instance.TransactionTaskDone()

OUT = filteredColumns 


def duplicateColumns(baseColumn, bList, hList):
    newColumns = []
    sizeList = []
    ss = []
    for b, h in zip(bList, hList):
        size = "{} x {}mm".format(b, h)
        sizeList.append(size)
    TransactionManager.Instance.EnsureInTransaction(doc)
    try:
        for b, h, size in zip(bList, hList, sizeList):
            idList = List[ElementId]([c.Id for c in baseColumn])
            copyIds = ElementTransformUtils.CopyElements(doc, idList, doc, Transform.Identity, CopyPasteOptions())
            for id in copyIds:
                newCol = doc.GetElement(id)
                newCol.Name = size
                bParam = newCol.LookupParameter('b')
                hParam = newCol.LookupParameter('h')
                if bParam and hParam:
                    bParam.Set(b/304.8)
                    hParam.Set(h/304.8)
                newColumns.append(newCol)
    except Exception as e:
        pass
    TransactionManager.Instance.TransactionTaskDone()
    return newColumns