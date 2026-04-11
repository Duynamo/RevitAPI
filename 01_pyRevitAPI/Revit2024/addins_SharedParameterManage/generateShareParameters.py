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
try:
    from Autodesk.Revit.DB.Structure import *
except:
    pass

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI.Selection import ISelectionFilter
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

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
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

# --- Excel Late Binding ---
def get_hidden_excel(file_path):
    # Always spawn a brand new hidden background process to ensure absolutely zero flashes or popup
    excel_type = System.Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise Exception("Excel is not installed or COM registry is missing.")
        
    excel = System.Activator.CreateInstance(excel_type)
    
    # Strictly enforce hidden UI and disable updates to prevent ghost window popups
    excel.Visible = False
    excel.ScreenUpdating = False
    excel.DisplayAlerts = False
    
    try:
        workbook = excel.Workbooks.Open(file_path)
        return excel, workbook
    except Exception as e:
        try:
            excel.Quit()
        except:
            pass
        raise e

def select_excel_file():
    dialog = OpenFileDialog()
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dialog.Title = "Select Excel file"
    if dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
        return dialog.FileName
    return None

def read_excel_data(file_path, sheet_name):
    """Read data from 'SharedParameter' sheet via Late Binding silently in the background."""
    if not os.path.exists(file_path):
        return None, "Error: File '%s' does not exist" % file_path
    
    excel = None
    workbook = None
    try:
        excel, workbook = get_hidden_excel(file_path)
        
        sheet = None
        for i in range(1, workbook.Worksheets.Count + 1):
            if str(workbook.Worksheets(i).Name) == str(sheet_name):
                sheet = workbook.Worksheets(i)
                break
                
        if sheet is None:
            return None, "Sheet '%s' does not exist" % sheet_name
        
        used_range = sheet.UsedRange
        last_row = used_range.Rows.Count
        last_col = used_range.Columns.Count
        
        if last_row < 1:
            return None, "Sheet '%s' is empty" % sheet_name
            
        target_col_index = -1
        for i in range(1, last_col + 1):
            cell_val = sheet.Cells(1, i).Value2
            if cell_val is not None and str(cell_val).strip() == "Shared Parameters Name":
                target_col_index = i
                break
                
        if target_col_index == -1:
            return None, "Column 'Shared Parameters Name' not found in the first row."
            
        param_names = []
        for r in range(2, last_row + 1):
            val = sheet.Cells(r, target_col_index).Value2
            if val is not None and str(val).strip() != "":
                param_names.append(str(val).strip())
                
        return param_names, None
        
    except Exception as e:
        return None, "Error reading Excel file: " + str(e)
    finally:
        # Guarantee memory flush and background process termination
        if workbook is not None:
            try:
                workbook.Close(False)
            except:
                pass
        if excel is not None:
            try:
                # Quit background instance permanently
                excel.Quit()
            except:
                pass

def create_and_bind_shared_parameters(param_names):
    try:
        shared_param_file = app.OpenSharedParameterFile()
        if shared_param_file is None:
            desktop_path = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
            new_file_path = os.path.join(desktop_path, "Temp_SharedParameters.txt")
            
            msg = "Could not find an active Shared Parameters file.\n\nA new temporary file:\n'Temp_SharedParameters.txt'\nwill be created on your Desktop so the script can proceed."
            System.Windows.Forms.MessageBox.Show(msg, "Auto-Create Notification", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            
            # Scaffolding correct Shared Parameter structurally
            with open(new_file_path, "w") as f:
                f.write("# This is a Revit shared parameter file.\n")
                f.write("# Do not edit manually.\n")
                f.write("*META\tVERSION\tMINVERSION\n")
                f.write("META\t2\t1\n")
                f.write("*GROUP\tID\tNAME\n")
                f.write("*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE\n")
                
            app.SharedParametersFilename = new_file_path
            shared_param_file = app.OpenSharedParameterFile()
            
            if shared_param_file is None:
                return False, "Failed to forcefully initialize the temporary Shared Parameters file."
    except Exception as e:
        return False, "Error during Shared Parameters initialization: " + str(e)

    t = Transaction(doc, "Add/Update Shared Parameters")
    t.Start()

    try:
        group_name = "ProjectParameters"
        group = shared_param_file.Groups.get_Item(group_name)
        if group is None:
            group = shared_param_file.Groups.Create(group_name)

        cat_set = app.Create.NewCategorySet()
        for cat in doc.Settings.Categories:
            try:
                if cat.AllowsBoundParameters:
                    cat_set.Insert(cat)
            except:
                pass

        for param_name in param_names:
            definition = group.Definitions.get_Item(param_name)
            if definition is None:
                try:
                    opt = ExternalDefinitionCreationOptions(param_name, SpecTypeId.String.Text)
                except:
                    opt = ExternalDefinitionCreationOptions(param_name, ParameterType.Text)
                    
                opt.Visible = True
                opt.UserModifiable = True
                definition = group.Definitions.Create(opt)

            binding_map = doc.ParameterBindings
            existing_binding = None
            
            it = binding_map.ForwardIterator()
            it.Reset()
            while it.MoveNext():
                if it.Key.Name == definition.Name:
                    existing_binding = it.Current
                    break

            instance_binding = app.Create.NewInstanceBinding(cat_set)

            if existing_binding is None:
                binding_map.Insert(definition, instance_binding, BuiltInParameterGroup.PG_DATA)
            else:
                binding_map.ReInsert(definition, instance_binding, BuiltInParameterGroup.PG_DATA)

        t.Commit()
        return True, "Success !!! Loaded %s parameters." % len(param_names)
    except Exception as e:
        t.RollBack()
        return False, "Error creating parameters: " + str(e)

# ================= THE MAIN RUN LOGIC ================= #
excelPath = select_excel_file()
OUT = "No operation has been completed yet."

if excelPath:
    sheetName = "SharedParameter"
    param_names, error_message = read_excel_data(excelPath, sheetName)
    
    if error_message:
        System.Windows.Forms.MessageBox.Show(error_message, "System Error")
        OUT = error_message
    elif param_names:
        success, message = create_and_bind_shared_parameters(param_names)
        System.Windows.Forms.MessageBox.Show(message, "Notification")
        OUT = message
    else:
        System.Windows.Forms.MessageBox.Show("No Shared Parameters found in the file.", "Empty")
        OUT = "No Shared Parameters found in the file."
else:
    OUT = "Excel file was not selected."