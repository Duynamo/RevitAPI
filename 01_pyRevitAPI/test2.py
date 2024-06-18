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

def placeFamilyInstances(doc, listCol, listPoints, level1):
    newInstances = []
    TransactionManager.Instance.EnsureInTransaction(doc)
    try:
        for symbol, point in zip(listCol, listPoints):
            if isinstance(symbol, FamilySymbol):
                if not symbol.IsActive:
                    symbol.Activate()
                    doc.Regenerate()
                
                newInst = doc.Create.NewFamilyInstance(XYZ(point.X, point.Y, point.Z), symbol, level, StructuralType.NonStructural)
                newInstances.append(newInst)

    except Exception as e:
        pass
    return newInstances


def duplicateColumns(baseColumn, bList, hList):
    newColumns = []
    sizeList = []
    existingSizes = set()

    for b, h in zip(bList, hList):
        size = "{} x {}mm".format(int(b), int(h))
        sizeList.append(size)
    for col in baseColumn:
        existingSizes.add(col.Name)
		
    TransactionManager.Instance.EnsureInTransaction(doc)
    try:
        for b, h, size in zip(bList, hList, sizeList):
            if size in existingSizes:
                continue
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



def getSizeString(colList):
    bParam = c.LookupParameter('b')
    hParam = c.LookupParameter('h')
    if bParam and hParam:
        b = int(round(bParam.AsDouble() * 304.8))
        h = int(round(hParam.AsDouble() * 304.8))
    return None


def duplicateColumns(baseColumn, bList, hList):
    newColumns = []
    sizeList = []
    ss = []
    for b, h in zip(bList, hList):
        size = "{} x {}mm".format(int(b), int(h))
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


def allConcreteColumns(doc):
    cate = BuiltInCategory.OST_StructuralColumns
    colCollector = FilteredElementCollector(doc).OfCategory(cate).WhereElementIsElementType()
    allConCols = []
    for c in colCollector:
        material_param = c.Family.LookupParameter('Material for Model Behavior')
        if material_param and 'Concrete' in material_param.AsValueString():
            allConCols.append(c)
    return allConCols

def duplicateColumns(doc, bList, hList):
    newColumns = []
    sizeList = []
    baseColumn = allConcreteColumns(doc)
    
    # Create a dictionary to store existing sizes in the family with corresponding columns
    existingSizes = {}
    for col in baseColumn:
        bParam = col.LookupParameter('b')
        hParam = col.LookupParameter('h')
        if bParam and hParam:
            size = "{} x {}mm".format(int(bParam.AsDouble() * 304.8), int(hParam.AsDouble() * 304.8))
            existingSizes[size] = col

    # Generate the list of sizes based on bList and hList
    for b, h in zip(bList, hList):
        size = "{} x {}mm".format(int(b), int(h))
        sizeList.append(size)

    TransactionManager.Instance.EnsureInTransaction(doc)
    try:
        for b, h, size in zip(bList, hList, sizeList):
            # Check if the size already exists
            if size not in existingSizes:
                # Duplicate the base column
                idList = List[ElementId]([c.Id for c in baseColumn])
                copyIds = ElementTransformUtils.CopyElements(doc, idList, doc, Transform.Identity, CopyPasteOptions())
                
                for id in copyIds:
                    newCol = doc.GetElement(id)
                    newCol.Name = size
                    bParam = newCol.LookupParameter('b')
                    hParam = newCol.LookupParameter('h')
                    if bParam and hParam:
                        bParam.Set(b / 304.8)  # Convert from mm to feet
                        hParam.Set(h / 304.8)  # Convert from mm to feet
                    newColumns.append(newCol)
                    #existingSizes[size] = newCol  # Add the new size to the dictionary of existing sizes
    except Exception as e:
        pass
    TransactionManager.Instance.TransactionTaskDone()
    return newColumns