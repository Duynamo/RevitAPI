import os
import clr
import sys
import System
import math

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB.Plumbing import *
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *
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

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel

"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""

#region _function
def select_excel_file():
    """Hiển thị hộp thoại để chọn file Excel."""
    dialog = OpenFileDialog()
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dialog.Title = "Chọn file Excel"
    if dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
        return dialog.FileName
    return None

def read_excel_data(file_path, sheet_name):
    """Đọc dữ liệu từ file Excel bằng Microsoft.Office.Interop.Excel."""
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = False
        workbook = excel.Workbooks.Open(file_path)
        try:
            sheet = workbook.Worksheets[sheet_name]
        except:
            workbook.Close(False)
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            return None, "Sheet '%s' không tồn tại" % sheet_name
        
        # Xác định phạm vi dữ liệu
        used_range = sheet.UsedRange
        last_row = used_range.Rows.Count
        last_col = used_range.Columns.Count
        data = []
        
        # Kiểm tra sheet có dữ liệu không
        if last_row < 1:
            workbook.Close(False)
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            return None, "Sheet '%s' rỗng" % sheet_name
        
        # Đọc tiêu đề từ dòng 1
        headers = []
        for col_idx in range(1, last_col + 1):
            cell = sheet.Cells[1, col_idx]
            header_value = str(cell.Value2) if cell.Value2 is not None else "Column_%s" % col_idx
            headers.append(header_value)
        
        # Kiểm tra có tiêu đề Element_ID không
        if "Element_ID" not in headers:
            workbook.Close(False)
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
            return None, "Sheet '%s' thiếu tiêu đề 'Element_ID'" % sheet_name
        
        # Đọc dữ liệu từ dòng 2
        for row_idx in range(2, last_row + 1):
            row_data = {}
            for idx, header in enumerate(headers):
                cell = sheet.Cells[row_idx, idx + 1]
                cell_value = cell.Value2 if cell.Value2 is not None else None
                row_data[header] = cell_value
            if row_data.get("Element_ID"):
                data.append(row_data)
            else:
                data.append({"Error": "Dòng %s: Thiếu hoặc rỗng Element_ID" % row_idx})
        
        # Đóng Excel
        workbook.Close(False)
        excel.Quit()
        # Giải phóng COM objects
        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        return headers, data
    
    except Exception as e:
        try:
            workbook.Close(False)
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        except:
            pass
        return None, "Lỗi đọc file Excel: %s" % str(e)

def get_element_by_id(element_id):
    """Tìm phần tử Revit theo Element ID."""
    try:
        # Xử lý element_id có thể là số thực hoặc chuỗi
        if isinstance(element_id, float):
            element_id = int(element_id)
        elif isinstance(element_id, str):
            element_id = int(float(element_id.strip()))
        return doc.GetElement(ElementId(element_id))
    except:
        return None

def update_parameter(element, param_name, param_value):
    """Cập nhật tham số kiểu text của phần tử Revit (instance hoặc type)."""
    try:
        # Tìm instance parameter
        param = element.LookupParameter(param_name)
        if not param:
            # Tìm type parameter nếu instance không tồn tại
            element_type = doc.GetElement(element.GetTypeId())
            if element_type:
                param = element_type.LookupParameter(param_name)
        if not param:
            return False, "Tham số '%s' không tồn tại" % param_name
        if param.IsReadOnly:
            return False, "Tham số '%s' là chỉ đọc" % param_name
        TransactionManager.Instance.EnsureInTransaction(doc)
        param.Set(str(param_value) if param_value is not None else "")
        TransactionManager.Instance.TransactionTaskDone()
        return True, "Đã cập nhật tham số '%s' thành '%s'" % (param_name, param_value if param_value is not None else "")
    except Exception as e:
        return False, "Lỗi khi cập nhật tham số '%s': %s" % (param_name, str(e))

def process_excel_data(headers, data):
    """Xử lý dữ liệu Excel và cập nhật mô hình Revit."""
    results = []
    TransactionManager.Instance.EnsureInTransaction(doc)
    
    for row in data:
        if "Error" in row:
            results.append(row["Error"])
            continue
        
        element_id = row.get("Element_ID")
        element = get_element_by_id(element_id)
        if not element:
            pass
            # results.append("Element_ID %s: Không tìm thấy phần tử" % element_id)
            # continue
        
        row_result = "Element_ID %s: " % element_id
        param_results = []
        for param_name in headers[1:]:  # Bỏ Element_ID
            TransactionManager.Instance.EnsureInTransaction(doc)
            param_value = row.get(param_name)
            success, message = update_parameter(element, param_name, param_value)
            TransactionManager.Instance.TransactionTaskDone()
            param_results.append(message)
        
        if param_results:
            row_result += "; ".join(param_results)
        else:
            row_result += "Không có tham số nào được cập nhật"
        results.append(row_result)
    
    TransactionManager.Instance.TransactionTaskDone()
    return results
def get_excel_worksheet_names(file_path):
    """Trả về danh sách tên của tất cả worksheet trong file Excel."""
    try:
        import clr
        clr.AddReference("Microsoft.Office.Interop.Excel")
        from Microsoft.Office.Interop import Excel
        import System.Runtime.InteropServices
        
        excel = Excel.ApplicationClass()
        excel.Visible = False
        workbook = excel.Workbooks.Open(file_path)
        
        # Lấy danh sách tên worksheet
        sheet_names = [str(sheet.Name) for sheet in workbook.Worksheets]
        
        # Đóng Excel
        workbook.Close(False)
        excel.Quit()
        
        # Giải phóng COM objects
        for sheet in workbook.Worksheets:
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        
        return sheet_names
    
    except Exception as e:
        pass
        # try:
        #     workbook.Close(False)
        #     excel.Quit()
        #     for sheet in workbook.Worksheets:
        #         System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        #     System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        #     System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)
        # except:
        #     pass
        # return None
#enregion 

class MainForm(Form):
    def __init__(self):
        self.InitializeComponent()
    
    def InitializeComponent(self):
        # Lấy kích thước màn hình làm việc
        screen_width = Screen.PrimaryScreen.WorkingArea.Width
        screen_height = Screen.PrimaryScreen.WorkingArea.Height
        
        # Tính kích thước UI: 1/4 chiều ngang, 1/2 chiều cao
        form_width = screen_width / 4
        form_height = screen_height / 2
        
        self._btt_Run = System.Windows.Forms.Button()
        self._btt_Cancle = System.Windows.Forms.Button()
        self._grb_data = System.Windows.Forms.GroupBox()
        self._btt_linkExcel = System.Windows.Forms.Button()
        self._txb_linkExcel = System.Windows.Forms.TextBox()
        self._lbl_D = System.Windows.Forms.Label()
        self._lbl_WorkSheetName = System.Windows.Forms.Label()
        self._comboBox1 = System.Windows.Forms.ComboBox()
        self._grb_data.SuspendLayout()
        self.SuspendLayout()
        
        # btt_Run
        self._btt_Run.BackColor = System.Drawing.Color.FromArgb(255, 192, 255)
        self._btt_Run.Font = System.Drawing.Font("Meiryo UI", 15, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_Run.Location = System.Drawing.Point(int(form_width * 0.401), int(form_height * 0.814))
        self._btt_Run.Name = "btt_Run"
        self._btt_Run.Size = System.Drawing.Size(int(form_width * 0.268), int(form_height * 0.131))
        self._btt_Run.TabIndex = 0
        self._btt_Run.Text = "RUN"
        self._btt_Run.UseVisualStyleBackColor = False
        self._btt_Run.Click += self.Btt_RunClick
        
        # btt_Cancle
        self._btt_Cancle.BackColor = System.Drawing.Color.FromArgb(255, 192, 255)
        self._btt_Cancle.Font = System.Drawing.Font("Meiryo UI", 15, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_Cancle.Location = System.Drawing.Point(int(form_width * 0.707), int(form_height * 0.814))
        self._btt_Cancle.Name = "btt_Cancle"
        self._btt_Cancle.Size = System.Drawing.Size(int(form_width * 0.268), int(form_height * 0.131))
        self._btt_Cancle.TabIndex = 1
        self._btt_Cancle.Text = "CANCLE"
        self._btt_Cancle.UseVisualStyleBackColor = False
        self._btt_Cancle.Click += self.Btt_CancleClick
        
        # grb_data
        self._grb_data.Controls.Add(self._comboBox1)
        self._grb_data.Controls.Add(self._lbl_WorkSheetName)
        self._grb_data.Controls.Add(self._txb_linkExcel)
        self._grb_data.Controls.Add(self._btt_linkExcel)
        self._grb_data.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._grb_data.Location = System.Drawing.Point(int(form_width * 0.026), int(form_height * 0.029))
        self._grb_data.Name = "grb_data"
        self._grb_data.Size = System.Drawing.Size(int(form_width * 0.948), int(form_height * 0.691))
        self._grb_data.TabIndex = 2
        self._grb_data.TabStop = False
        self._grb_data.Text = "data"
        
        # btt_linkExcel
        self._btt_linkExcel.BackColor = System.Drawing.Color.FromArgb(255, 192, 192)
        self._btt_linkExcel.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_linkExcel.Location = System.Drawing.Point(int(self._grb_data.Width * 0.013), int(self._grb_data.Height * 0.118))
        self._btt_linkExcel.Name = "btt_linkExcel"
        self._btt_linkExcel.Size = System.Drawing.Size(int(self._grb_data.Width * 0.298), int(self._grb_data.Height * 0.145))
        self._btt_linkExcel.TabIndex = 3
        self._btt_linkExcel.Text = "Select Link"
        self._btt_linkExcel.UseVisualStyleBackColor = False
        self._btt_linkExcel.Click += self.Btt_linkExcelClick
        
        # txb_linkExcel
        self._txb_linkExcel.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._txb_linkExcel.Location = System.Drawing.Point(int(self._grb_data.Width * 0.349), int(self._grb_data.Height * 0.167))
        self._txb_linkExcel.Name = "txb_linkExcel"
        self._txb_linkExcel.Size = System.Drawing.Size(int(self._grb_data.Width * 0.6), int(self._grb_data.Height * 0.064))
        self._txb_linkExcel.TabIndex = 4
        self._txb_linkExcel.TextChanged += self.Txb_linkExcelTextChanged
        
        # lbl_D
        self._lbl_D.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._lbl_D.ForeColor = System.Drawing.Color.Blue
        self._lbl_D.Location = System.Drawing.Point(int(form_width * 0.024), int(form_height * 0.929))
        self._lbl_D.Name = "lbl_D"
        self._lbl_D.Size = System.Drawing.Size(int(form_width * 0.085), int(form_height * 0.051))
        self._lbl_D.TabIndex = 3
        self._lbl_D.Text = "@D"
        self._lbl_D.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        
        # lbl_WorkSheetName
        self._lbl_WorkSheetName.BackColor = System.Drawing.Color.FromArgb(255, 192, 192)
        self._lbl_WorkSheetName.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._lbl_WorkSheetName.Location = System.Drawing.Point(int(self._grb_data.Width * 0.013), int(self._grb_data.Height * 0.344))
        self._lbl_WorkSheetName.Name = "lbl_WorkSheetName"
        self._lbl_WorkSheetName.Size = System.Drawing.Size(int(self._grb_data.Width * 0.298), int(self._grb_data.Height * 0.122))
        self._lbl_WorkSheetName.TabIndex = 5
        self._lbl_WorkSheetName.Text = "Work Sheet"
        self._lbl_WorkSheetName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        
        # comboBox1
        self._comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        self._comboBox1.FormattingEnabled = True
        self._comboBox1.Location = System.Drawing.Point(int(self._grb_data.Width * 0.349), int(self._grb_data.Height * 0.378))
        self._comboBox1.Name = "comboBox1"
        self._comboBox1.Size = System.Drawing.Size(int(self._grb_data.Width * 0.6), int(self._grb_data.Height * 0.087))
        self._comboBox1.TabIndex = 6
        self._comboBox1.UseWaitCursor = True
        self._comboBox1.SelectedIndexChanged += self.ComboBox1SelectedIndexChanged
        
        # MainForm
        self.ClientSize = System.Drawing.Size(int(form_width), int(form_height))
        self.Controls.Add(self._lbl_D)
        self.Controls.Add(self._grb_data)
        self.Controls.Add(self._btt_Cancle)
        self.Controls.Add(self._btt_Run)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.Name = "MainForm"
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        self.Text = "nonParameter"
        self.TopMost = True
        self._grb_data.ResumeLayout(False)
        self._grb_data.PerformLayout()
        self.ResumeLayout(False)

    def Btt_linkExcelClick(self, sender, e):
        openDialog = OpenFileDialog()
        openDialog.Multiselect = False
        openDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        if openDialog.ShowDialog() == DialogResult.OK:
            file_path = openDialog.FileName  # Lấy đường dẫn file duy nhất
            self._txb_linkExcel.Text = file_path  # Hiển thị đường dẫn đầy đủ
            
            # Xóa danh sách cũ trong comboBox1
            self._comboBox1.Items.Clear()
            
            # Lấy danh sách tên worksheet
            sheet_names = get_excel_worksheet_names(file_path)
            if sheet_names :
                for sheet_name in sheet_names:
                    self._comboBox1.Items.Add(sheet_name)
                # Chọn mặc định sheet đầu tiên (nếu có)
                if self._comboBox1.Items.Count > 0:
                    self._comboBox1.SelectedIndex = 0
            else:
                pass
    def Txb_linkExcelTextChanged(self, sender, e):

        pass

    def ComboBox1SelectedIndexChanged(self, sender, e):
        pass

    def Btt_RunClick(self, sender, e):
        pass

    def Btt_CancleClick(self, sender, e):
        self.Close()
        pass

f = MainForm()
Application.Run(f)