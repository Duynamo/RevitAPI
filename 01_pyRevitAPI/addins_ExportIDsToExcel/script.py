"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
# -*- coding: utf-8 -*-
import clr
import sys
import System
import math
import collections
import os

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import * 
from Autodesk.Revit.DB.Structure import *
import Autodesk.Revit.Exceptions as RevitExceptions

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI.Selection import ISelectionFilter, Selection
from Autodesk.Revit.UI import *

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

clr.AddReference('System.Windows.Forms')
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms.DataVisualization")

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *
#endregion

"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""

def alert(message, title="Notification"):
    System.Windows.Forms.MessageBox.Show(message, title, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

def get_running_excel():
    try:
        from System.Runtime.InteropServices import Marshal
        return Marshal.GetActiveObject("Excel.Application")
    except Exception as e:
        try:
            clr.AddReference("Microsoft.VisualBasic")
            import Microsoft.VisualBasic
            return Microsoft.VisualBasic.Interaction.GetObject(None, "Excel.Application")
        except Exception as inner_e:
            raise Exception("Framework Error: " + str(e) + " | Core Fallback Error: " + str(inner_e))

def get_or_open_excel(file_path):
    excel = None
    try:
        excel = get_running_excel()
    except:
        pass
        
    if not excel:
        excel_type = System.Type.GetTypeFromProgID("Excel.Application")
        if excel_type is None:
            raise Exception("Excel is not installed or COM registry is missing.")
        excel = System.Activator.CreateInstance(excel_type)
        
    excel.Visible = True
    
    workbook = None
    if excel.Workbooks.Count > 0:
        for i in range(1, excel.Workbooks.Count + 1):
            wb = excel.Workbooks(i)
            if wb.FullName.lower() == file_path.lower():
                workbook = wb
                break
                
    if not workbook:
        workbook = excel.Workbooks.Open(file_path)
        
    return excel, workbook

# --- Execution ---

OUT = "No operation has been completed yet"

dialog = OpenFileDialog()
dialog.Filter = "Files Excel|*.xlsx;*.xls;*.xlsm|All Files|*.*"
dialog.Title = "Select Excel file to export ElementID"

if dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
    excel_file = dialog.FileName
    
    try:
        excel_app, workbook = get_or_open_excel(excel_file)
        
        # Find ElementID sheet
        ws = None
        for i in range(1, workbook.Worksheets.Count + 1):
            if workbook.Worksheets(i).Name == "ElementID":
                ws = workbook.Worksheets(i)
                break
        
        if ws is None:
            alert("Could not find a Worksheet named 'ElementID' in your file.", "Error")
        else:
            # Find Column Index
            col_index = -1
            for i in range(1, 40):
                val = ws.Cells(1, i).Value2
                if val is not None and str(val).strip() == "ElementID":
                    col_index = i
                    break
                    
            if col_index == -1:
                alert("Could not find the 'ElementID' header in the first row.", "Error")
            else:
                msg = ("LIVE EXPORT INSTRUCTIONS:\n\n"
                       "- Click to select an Element. Its ID will be INSTANTLY exported to 'Sheet: ElementID' -> Column 'ElementID'.\n"
                       "- The Excel file remains open. If you select the wrong Element, simply delete the target cell or row directly in Excel.\n"
                       "- The script will automatically adapt to your deletions and write the next ID into the correct empty cell.\n"
                       "- PRESS 'ESC' AT ANY TIME TO FINISH THE PROCESS.\n")
                       
                alert(msg, "Instructions")
                
                exported_ids = []
                
                # PickObject Loop
                while True:
                    try:
                        ref = uidoc.Selection.PickObject(Selection.ObjectType.Element, "Select Element... [PRESS ESC to finish]")
                        el_id = ref.ElementId
                        
                        try:
                            id_val = el_id.Value
                        except:
                            id_val = el_id.IntegerValue
                        
                        
                        # Fix: Constantly recalculate the last row so the user can freely delete rows in Excel
                        try:
                            last_row = ws.Cells(ws.Rows.Count, col_index).End(-4162).Row
                        except:
                            last_row = 1
                            for i in range(1, 10000):
                                if ws.Cells(i, col_index).Value2 is not None:
                                    last_row = i
                                else:
                                    if i - last_row > 5:
                                        break
                                        
                        # Check for duplicates LIVE targeting only the active Excel column
                        existing_ids = set()
                        if last_row >= 2:
                            try:
                                vals = ws.Range(ws.Cells(2, col_index), ws.Cells(last_row, col_index)).Value2
                                if vals is not None:
                                    if hasattr(vals, '__iter__'):
                                        for item in vals:
                                            if item is not None:
                                                try:
                                                    existing_ids.add(int(float(item)))
                                                except:
                                                    pass
                                    else:
                                        try:
                                            existing_ids.add(int(float(vals)))
                                        except:
                                            pass
                            except:
                                pass
                                
                        if int(id_val) in existing_ids:
                            alert("This Element ID ({}) already exists in the Excel file!".format(id_val), "Duplicate Warning")
                            continue
                                        
                        current_row = last_row + 1
                        if current_row < 2:
                            current_row = 2
                                
                        # Write seamlessly
                        ws.Cells(current_row, col_index).Value2 = id_val
                        
                        exported_ids.append(id_val)
                        
                    except RevitExceptions.OperationCanceledException:
                        break
                    except Exception as e:
                        alert("An error occurred during selection: " + str(e), "Error")
                        break
                        
                alert("Success! Added {} IDs to Excel.".format(len(exported_ids)), "Completed")
                OUT = "Exported {} IDs successfully".format(len(exported_ids))
                
    except Exception as e:
        alert("COM Application Error:\n" + str(e), "Error")
        OUT = "Failed: " + str(e)
