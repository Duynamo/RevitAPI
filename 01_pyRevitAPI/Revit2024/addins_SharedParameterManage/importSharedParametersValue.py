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
#region _function

def get_hidden_excel(file_path):
    # Always spawn a brand new hidden background process to ensure absolutely zero flashes or popup
    excel_type = System.Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise Exception("Excel is not installed or COM registry is missing.")
        
    excel = System.Activator.CreateInstance(excel_type)
    
    # Strictly enforce hidden UI and disable updates
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

def read_excel_data(file_path, sheet_name):
    """Read data from Excel file using strict Background Late Binding."""
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
            return None, "Sheet '%s' does not exist." % sheet_name
        
        used_range = sheet.UsedRange
        last_row = used_range.Rows.Count
        last_col = used_range.Columns.Count
        data = []
        
        if last_row < 1:
            return None, "Sheet '%s' is empty." % sheet_name
        
        headers = []
        for col_idx in range(1, last_col + 1):
            cell = sheet.Cells(1, col_idx)
            header_value = str(cell.Value2) if cell.Value2 is not None else "Column_%s" % col_idx
            headers.append(header_value)
        
        if "ElementID" not in headers:
            return None, "Sheet '%s' is missing 'ElementID' header." % sheet_name
        
        for row_idx in range(2, last_row + 1):
            row_data = {}
            for idx, header in enumerate(headers):
                cell = sheet.Cells(row_idx, idx + 1)
                cell_value = cell.Value2 if cell.Value2 is not None else None
                row_data[header] = cell_value
                
            if row_data.get("ElementID") is not None and str(row_data["ElementID"]).strip() != "":
                data.append(row_data)
        
        return headers, data
    
    except Exception as e:
        return None, "Error reading Excel file: %s" % str(e)
    finally:
        # Guarantee memory flush and background process termination
        if workbook is not None:
            try:
                workbook.Close(False)
            except:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except:
                pass

def get_element_by_id(element_id):
    try:
        if isinstance(element_id, float):
            element_id = int(element_id)
        elif isinstance(element_id, str):
            try:
                element_id = int(float(element_id.strip()))
            except:
                pass
                
        try:
            return doc.GetElement(ElementId(element_id))
        except:
            import System
            return doc.GetElement(ElementId(System.Int64(element_id)))
    except:
        return None

def update_parameter(element, param_name, param_value):
    try:
        param = element.LookupParameter(param_name)
        if not param:
            element_type = doc.GetElement(element.GetTypeId())
            if element_type:
                param = element_type.LookupParameter(param_name)
        if not param:
            return False, "Parameter does not exist"
        if param.IsReadOnly:
            return False, "Parameter is read-only"
        
        TransactionManager.Instance.EnsureInTransaction(doc)
        param.Set(str(param_value) if param_value is not None else "")
        TransactionManager.Instance.TransactionTaskDone()
        return True, "Update OK"
    except:
        return False, "Update Error"
    
def process_excel_data(headers, data):
    TransactionManager.Instance.EnsureInTransaction(doc)
    success_count = 0
    error_count = 0
    
    for row in data:
        if "Error" in row:
            # If the row doesn't have an ID
            error_count += max(0, len(headers) - 1)
            continue
        
        element_id = row.get("ElementID")
        element = get_element_by_id(element_id)
        if not element:
            # If element doesn't exist, all params for it fail
            error_count += max(0, len(headers) - 1)
            continue
        
        for param_name in headers:
            if param_name == "ElementID":
                continue
            param_value = row.get(param_name)
            if param_value is not None:
                success, _ = update_parameter(element, param_name, param_value)
                if success:
                    success_count += 1
                else:
                    error_count += 1
                    
    TransactionManager.Instance.TransactionTaskDone()
    return success_count, error_count

def release_com(obj):
    pass
#endregion

class MainForm(Form):
    def __init__(self):
        self.InitializeComponent()
    
    def InitializeComponent(self):
        # Material Colors
        c_primary = Color.FromArgb(98, 0, 238)      # Purple 500
        c_success = Color.FromArgb(76, 175, 80)     # Green 500
        c_bg_main = Color.FromArgb(245, 245, 246)   # Light Gray Background
        c_surface = Color.White                     # Panel Surface
        c_text_dark = Color.FromArgb(33, 33, 33)    # Near Black Text
        c_btn_cancel = Color.FromArgb(224, 224, 224)# Gray Cancel Button
        
        self.Text = "Import Shared Parameters"
        self.ClientSize = Size(500, 310)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.BackColor = c_bg_main
        self.Font = Font("Segoe UI", 10)
        
        # Panel
        self.panel = Panel()
        self.panel.BackColor = c_surface
        self.panel.Location = Point(20, 20)
        self.panel.Size = Size(460, 200)
        
        # Panel Title
        self.lbl_title = Label()
        self.lbl_title.Text = "Data Source"
        self.lbl_title.Font = Font("Segoe UI", 12, FontStyle.Bold)
        self.lbl_title.ForeColor = c_primary
        self.lbl_title.Location = Point(20, 15)
        self.lbl_title.AutoSize = True
        
        # Layout Metrics
        x_left = 20
        width_left = 160
        x_right = 190
        width_right = 255
        y_row1 = 60
        y_row2 = 115
        h_btn = 35
        h_input = 28
        
        # Select Excel Button (Purple)
        self.btt_linkExcel = Button()
        self.btt_linkExcel.Text = "SELECT EXCEL"
        self.btt_linkExcel.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.btt_linkExcel.BackColor = c_primary
        self.btt_linkExcel.ForeColor = Color.White
        self.btt_linkExcel.FlatStyle = FlatStyle.Flat
        self.btt_linkExcel.FlatAppearance.BorderSize = 0
        self.btt_linkExcel.Location = Point(x_left, y_row1)
        self.btt_linkExcel.Size = Size(width_left, h_btn)
        self.btt_linkExcel.Click += self.Btt_linkExcelClick
        
        # TextBox Link Excel
        self.txb_linkExcel = TextBox()
        self.txb_linkExcel.Location = Point(x_right, y_row1 + 4)
        self.txb_linkExcel.Size = Size(width_right, h_input)
        self.txb_linkExcel.ReadOnly = True
        self.txb_linkExcel.BackColor = Color.FromArgb(250, 250, 250)
        self.txb_linkExcel.BorderStyle = BorderStyle.FixedSingle
        
        # ComboBox Label (Styled as a block to match Select Excel)
        self.lbl_ws = Label()
        self.lbl_ws.Text = "WORKSHEET"
        self.lbl_ws.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.lbl_ws.BackColor = c_primary
        self.lbl_ws.ForeColor = Color.White
        self.lbl_ws.Location = Point(x_left, y_row2)
        self.lbl_ws.Size = Size(width_left, h_btn)
        self.lbl_ws.TextAlign = ContentAlignment.MiddleCenter
        self.lbl_ws.AutoSize = False
        
        # ComboBox
        self.comboBox1 = ComboBox()
        self.comboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        self.comboBox1.Location = Point(x_right, y_row2 + 4)
        self.comboBox1.Size = Size(width_right, h_input)
        self.comboBox1.BackColor = Color.White
        
        # Watermark D
        self.lbl_D = Label()
        self.lbl_D.Font = Font("Segoe UI", 7, FontStyle.Bold)
        self.lbl_D.ForeColor = Color.DarkGray
        self.lbl_D.Location = Point(20, 165)
        self.lbl_D.Text = "@D"
        self.lbl_D.AutoSize = True
        
        # Assemble Panel
        self.panel.Controls.Add(self.lbl_title)
        self.panel.Controls.Add(self.btt_linkExcel)
        self.panel.Controls.Add(self.txb_linkExcel)
        self.panel.Controls.Add(self.lbl_ws)
        self.panel.Controls.Add(self.comboBox1)
        self.panel.Controls.Add(self.lbl_D)
        
        # RUN Button (Green)
        self.btt_Run = Button()
        self.btt_Run.Text = "RUN IMPORT"
        self.btt_Run.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.btt_Run.BackColor = c_success
        self.btt_Run.ForeColor = Color.White
        self.btt_Run.FlatStyle = FlatStyle.Flat
        self.btt_Run.FlatAppearance.BorderSize = 0
        self.btt_Run.Location = Point(215, 240)
        self.btt_Run.Size = Size(140, 42)
        self.btt_Run.Click += self.Btt_RunClick
        
        # CANCEL Button (Gray)
        self.btt_Cancle = Button()
        self.btt_Cancle.Text = "CANCEL"
        self.btt_Cancle.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self.btt_Cancle.BackColor = c_btn_cancel
        self.btt_Cancle.ForeColor = c_text_dark
        self.btt_Cancle.FlatStyle = FlatStyle.Flat
        self.btt_Cancle.FlatAppearance.BorderSize = 0
        self.btt_Cancle.Location = Point(370, 240)
        self.btt_Cancle.Size = Size(110, 42)
        self.btt_Cancle.Click += self.Btt_CancleClick
        
        self.Controls.Add(self.panel)
        self.Controls.Add(self.btt_Run)
        self.Controls.Add(self.btt_Cancle)
        self.TopMost = True

    def Btt_linkExcelClick(self, sender, e):
        dlg = OpenFileDialog()
        dlg.Multiselect = False
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        dlg.Title = "Select Excel file"
        if dlg.ShowDialog() == DialogResult.OK:
            path = dlg.FileName
            self.txb_linkExcel.Text = path
            
            excel = None
            workbook = None
            try:
                excel, workbook = get_hidden_excel(path)
                self.comboBox1.Items.Clear()
                for i in range(1, workbook.Worksheets.Count + 1):
                    ws = workbook.Worksheets(i)
                    self.comboBox1.Items.Add(str(ws.Name))
                if self.comboBox1.Items.Count > 0:
                    self.comboBox1.SelectedIndex = 0
            except Exception as ex:
                MessageBox.Show("Error connecting to Excel: " + str(ex), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            finally:
                if workbook is not None:
                    try: workbook.Close(False)
                    except: pass
                if excel is not None:
                    try: excel.Quit()
                    except: pass

    def Btt_RunClick(self, sender, e):
        sheet_name = self.comboBox1.SelectedItem
        file_path = self.txb_linkExcel.Text
        if not file_path or not sheet_name:
            MessageBox.Show("Please select a valid file and Worksheet.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
            
        headers, data = read_excel_data(file_path, sheet_name)
        if headers is None:
            MessageBox.Show(str(data), "Excel Processing Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        else:
            try:
                # Trả quá trình process cho DocumentManager bên ngoài vòng lặp
                success_count, error_count = process_excel_data(headers, data)
                
                # Show simplified notification
                if error_count > 0:
                    msg = "Completed with some errors.\n\nSuccessfully updated %d parameters.\nFailed updates: %d" % (success_count, error_count)
                    MessageBox.Show(msg, "Import Finished", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                else:
                    msg = "Success!\n\nSuccessfully updated %d parameters." % success_count
                    MessageBox.Show(msg, "Import Finished", MessageBoxButtons.OK, MessageBoxIcon.Information)
            except Exception as ex:
                MessageBox.Show("Error writing to Revit: " + str(ex), "Update Element Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                
        self.Close()

    def Btt_CancleClick(self, sender, e):
        self.Close()

f = MainForm()
Application.Run(f)
