"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
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
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB import Line, XYZ, IntersectionResultArray, SetComparisonResult

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI.Selection import ISelectionFilter
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

clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel
import System.Runtime.InteropServices
import System.Runtime.InteropServices as Marshal

#endregion

"""___"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""___"""

try:
    clr.AddReference("IronPython.wpf")
    clr.AddReference('PresentationCore')
    clr.AddReference('PresentationFramework')
except IOError:
    raise
from System.IO import StringReader
from System.Windows.Markup import XamlReader, XamlWriter
from System.Windows import Window, Application, MessageBox, MessageBoxButton, MessageBoxResult

try:
    import wpf
except ImportError:
    raise

# =====================================================================
# Khởi tạo excelFile = None ở module level để tránh NameError
# =====================================================================
excelFile = None

"""___"""
def select_excel_file():
    """Hiển thị hộp thoại để chọn file Excel."""
    dialog = OpenFileDialog()
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dialog.Title = "Chọn file Excel"
    if dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK:
        return dialog.FileName
    return None

# Hàm xuất danh sách ID vào file Excel
def export_to_excel(element_ids, file_path):
    excel = None
    workbook = None
    sheet = None
    try:
        excel = Excel.ApplicationClass()
        excel.Visible = False
        excel.DisplayAlerts = False

        file_exists = os.path.exists(file_path)

        if file_exists:
            workbook = excel.Workbooks.Open(file_path)
        else:
            workbook = excel.Workbooks.Add()

        sheet_name = "ElementID"
        sheet = None
        for ws in workbook.Worksheets:
            if ws.Name == sheet_name:
                sheet = ws
                sheet.UsedRange.ClearContents()
                break
        if sheet is None:
            sheet = workbook.Worksheets.Add()
            sheet.Name = sheet_name

        sheet.Cells[1, 1].Value2 = "No"
        sheet.Cells[1, 2].Value2 = "ElementID"

        for idx, element_id in enumerate(element_ids, start=1):
            sheet.Cells[idx + 1, 1].Value2 = idx
            sheet.Cells[idx + 1, 2].Value2 = element_id.IntegerValue

        # FIX: file đã tồn tại → Save() giữ nguyên format gốc
        #      file mới (Add) → SaveAs với đủ tham số Missing
        if file_exists:
            workbook.Save()
        else:
            missing = System.Type.Missing
            workbook.SaveAs(
                file_path,  # Filename
                51,         # FileFormat: xlOpenXMLWorkbook (.xlsx)
                missing, missing, missing, missing,
                Excel.XlSaveAsAccessMode.xlNoChange,
                missing, missing, missing, missing, missing
            )
        return True

    except Exception as e:
        MessageBox.Show("Export error: {}".format(str(e)), "Error")
        return False

    finally:
        if sheet is not None:
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        if workbook is not None:
            workbook.Close(False)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        if excel is not None:
            excel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel)

def getParameterFromTypeId(elements):
    elements = UnwrapElement(elements)
    typeIds = [i.GetTypeId() for i in elements]
    familyTypes = [doc.GetElement(typeId) for typeId in typeIds]

    families = []
    IDs = []

    if elements:
        IDs = [ele.Id for ele in elements]
    for j in familyTypes:
        allParams = j.Parameters
        for param in allParams:
            if param.Definition.Name == "Family Name":
                families.append(param.AsString())

    return {'family_name': families, 'family_ID': IDs}    

# =====================================================================
# Thêm setter cho tất cả các property của viewModel
# vì update_listview gán vm.gvcNo = i → cần setter, không có
# setter sẽ gây AttributeError crash ngầm
# =====================================================================
class viewModel:
    def __init__(self, gvcNo, gvcFamilyName, gvcSelectID, gvcID):
        self._gvcNo = gvcNo
        self._gvcFamilyName = gvcFamilyName
        self._gvcSelectID = gvcSelectID
        self._gvcID = gvcID

    @property
    def gvcNo(self):
        return self._gvcNo

    @gvcNo.setter                       # <-- FIX: thêm setter
    def gvcNo(self, value):
        self._gvcNo = value

    @property
    def gvcFamilyName(self):
        return self._gvcFamilyName

    @gvcFamilyName.setter
    def gvcFamilyName(self, value):
        self._gvcFamilyName = value

    @property
    def gvcSelectID(self):
        return self._gvcSelectID

    @gvcSelectID.setter
    def gvcSelectID(self, value):
        self._gvcSelectID = value

    @property
    def gvcID(self):
        return self._gvcID

    @gvcID.setter
    def gvcID(self, value):
        self._gvcID = value

"""___"""

class MyWindow(Window):
    def __init__(self):
        self.ui = wpf.LoadComponent(self, r'C:\Users\Laptop\OneDrive\07_Python Dynamo\01_Duuynamo\Duuynamo.extension\vitaminD.tab\Other.panel\edit3.stack\Test260218.pushbutton\Form1.xaml')
        
        self.txt_excel_path = self.ui.FindName("ExcelPath")
        self.lvElements = self.ui.FindName("lvElements")
        
        self.view_models = []
        
        if self.lvElements:
            self.lvElements.ItemsSource = self.view_models

    def CheckBox_Checked(self, sender, e):
        self._select_row_from_checkbox(sender)

    def CheckBox_Unchecked(self, sender, e):
        self._select_row_from_checkbox(sender)

    def CheckBox_MouseDown(self, sender, e):
        self._select_row_from_checkbox(sender)

    def _select_row_from_checkbox(self, checkbox):
        try:
            parent = checkbox
            while parent is not None and not isinstance(parent, System.Windows.Controls.ListViewItem):
                parent = System.Windows.Media.VisualTreeHelper.GetParent(parent)
            
            if isinstance(parent, System.Windows.Controls.ListViewItem):
                # ✅ FIX: Đồng bộ IsSelected theo trạng thái checkbox
                vm = parent.DataContext
                if vm is not None and hasattr(vm, 'gvcSelectID'):
                    parent.IsSelected = vm.gvcSelectID
        except Exception as ex:
            pass

    def update_listview(self):
        if not self.lvElements:
            return
        for i, vm in enumerate(self.view_models, start=1):
            vm.gvcNo = i
        
        self.lvElements.ItemsSource = None
        self.lvElements.ItemsSource = self.view_models
        self.lvElements.Items.Refresh()   

    def lvElements_PreviewMouseLeftButtonDown(self, sender, e):
        try:
            source = e.OriginalSource

            # Nếu click trực tiếp vào CheckBox → bỏ qua
            parent_check = source
            while parent_check is not None:
                if isinstance(parent_check, System.Windows.Controls.CheckBox):
                    return
                parent_check = System.Windows.Media.VisualTreeHelper.GetParent(parent_check)

            # Tìm ListViewItem chứa vị trí click
            item = source
            while item is not None and not isinstance(item, System.Windows.Controls.ListViewItem):
                item = System.Windows.Media.VisualTreeHelper.GetParent(item)

            if item is None:
                return

            # ✅ FIX: Toggle cả IsSelected và gvcSelectID
            vm = item.DataContext
            if vm is not None and hasattr(vm, 'gvcSelectID'):
                if item.IsSelected:
                    item.IsSelected = False          # Bỏ chọn row
                    vm.gvcSelectID = False           # Uncheck checkbox
                else:
                    item.IsSelected = True           # Chọn row
                    vm.gvcSelectID = True            # Check checkbox

                e.Handled = True                     # Chặn event mặc định của ListView
                self.lvElements.Items.Refresh()

        except Exception as ex:
            pass

    def btnSelectExcel(self, sender, event):
        global excelFile
        selected = select_excel_file()

        if selected:
            excelFile = selected
            txt_excel_path = self.ui.FindName("ExcelPath")
            if txt_excel_path:
                txt_excel_path.Text = excelFile
            else:
                print("Không tìm thấy TextBlock 'ExcelPath'")

    def btnPickElements(self, sender, event):
        self.Hide()

        try:
            from Autodesk.Revit.UI.Selection import ObjectType
            picked_ref = uidoc.Selection.PickObject(ObjectType.Element, "Chọn một đối tượng")
            if not picked_ref:
                return

            element = doc.GetElement(picked_ref.ElementId)
            if not element:
                return

            elements = [element]
            dictParams = getParameterFromTypeId(elements)

            if not dictParams.get("family_name") or not dictParams.get("family_ID"):
                MessageBox.Show("Cannot get ID!", "Warning")
                return

            family_name = dictParams["family_name"][0] or "Unknown"
            new_id = dictParams["family_ID"][0]

            if any(vm.gvcID == new_id for vm in self.view_models):
                MessageBox.Show("Duplicated element", "Notification")
                return

            new_no = len(self.view_models) + 1
            vm = viewModel(
                gvcNo=new_no,
                gvcFamilyName=family_name,
                gvcSelectID=False,
                gvcID=new_id
            )
            self.view_models.append(vm)
            self.update_listview()
            
        except Exception as ex:
            MessageBox.Show("Pick error: {}".format(str(ex)), "Error")
        finally:

            self.Show()

    def btnDelete(self, sender, event):
        if not self.lvElements or not self.lvElements.SelectedItems:
            MessageBox.Show("Please select an item first.", "Notification")
            return       
        selected_items = list(self.lvElements.SelectedItems)       
        for item in selected_items:
            if item in self.view_models:
                self.view_models.remove(item)       
        self.update_listview()
        MessageBox.Show("Deleted successfully.", "Notification")

    def btnUp(self, sender, event):
        if not self.lvElements or not self.lvElements.SelectedItems:
            MessageBox.Show("Please select an item first.", "Notification")
            return
        try:
            selected_items = list(self.lvElements.SelectedItems)
            if not selected_items:
                return
            selected_indices = []
            for item in selected_items:
                try:
                    idx = self.view_models.index(item)
                    selected_indices.append(idx)
                except ValueError:
                    continue
            selected_indices.sort()
            if not selected_indices or selected_indices[0] == 0:
                return
            for idx in selected_indices:
                if idx > 0:
                    self.view_models[idx - 1], self.view_models[idx] = self.view_models[idx], self.view_models[idx - 1]
            self.update_listview()
            self.lvElements.SelectedItems.Clear()
            for old_idx in selected_indices:
                new_idx = old_idx - 1 if old_idx > 0 else old_idx
                if 0 <= new_idx < len(self.view_models):
                    self.lvElements.SelectedItems.Add(self.view_models[new_idx])
        except Exception as ex:
            MessageBox.Show("Move up error: {}".format(str(ex)), "Error")

    def btnDown(self, sender, event):
        if not self.lvElements or not self.lvElements.SelectedItems:
            MessageBox.Show("Please select an item first.", "Notification")
            return
        try:
            selected_items = list(self.lvElements.SelectedItems)
            if not selected_items:
                return
            selected_indices = []
            for item in selected_items:
                try:
                    idx = self.view_models.index(item)
                    selected_indices.append(idx)
                except ValueError:
                    continue
            selected_indices.sort(reverse=True)
            last_index = len(self.view_models) - 1
            if not selected_indices or selected_indices[0] == last_index:
                return
            for idx in selected_indices:
                if idx < last_index:
                    self.view_models[idx + 1], self.view_models[idx] = self.view_models[idx], self.view_models[idx + 1]
            self.update_listview()
            self.lvElements.SelectedItems.Clear()
            for old_idx in selected_indices:
                new_idx = old_idx + 1 if old_idx < last_index else old_idx
                if 0 <= new_idx < len(self.view_models):
                    self.lvElements.SelectedItems.Add(self.view_models[new_idx])
        except Exception as ex:
            MessageBox.Show("Error")

    def btnClear(self, sender, event):
        if not self.view_models:
            MessageBox.Show("List is already empty!", "Notification")
            return
        
        confirm = MessageBox.Show(
            "Are you sure you want to clear all items?",
            "Confirmation",
            MessageBoxButton.YesNo
        )
        
        if confirm == MessageBoxResult.Yes:
            self.view_models = []
            self.update_listview()
            MessageBox.Show("List cleared.", "Notification")

    def btnExportID(self, sender, event):
        """Xuất TOÀN BỘ ElementID có trong ListView ra sheet ElementID trong file Excel"""
        global excelFile  # FIX 1: excelFile được khởi tạo = None ở module level
        excel_path = excelFile  # an toàn vì excelFile đã được khởi tạo
        if not excel_path:
            MessageBox.Show("Please select an Excel file first!", "Notification")
            return
        if not self.view_models:
            MessageBox.Show("List is empty!", "Notification")
            return
        all_ids = []
        for vm in self.view_models:
            try:
                # gvcID là ElementId object → dùng trực tiếp
                if isinstance(vm.gvcID, ElementId):
                    all_ids.append(vm.gvcID)
                else:
                    # fallback nếu gvcID là int/string
                    all_ids.append(ElementId(int(vm.gvcID)))
            except Exception as e:
                continue       
        if not all_ids:
            MessageBox.Show("No valid ElementIDs found!", "Notification")
            return
        success = export_to_excel(all_ids, excel_path)
        if success:
            self._collect_result()         # ← THÊM
            
            MessageBox.Show("Export successful", "Notification")
            self.Close()

    def btnCancle(self, sender, event):
        self.Close()

myWindow = MyWindow()
myWindow.ShowDialog()

vms = myWindow.view_models

if vms:
    OUT = [
        [vm.gvcNo         for vm in vms],
        [vm.gvcFamilyName for vm in vms],
        [vm.gvcID.IntegerValue if isinstance(vm.gvcID, ElementId) 
         else int(vm.gvcID) for vm in vms],   # int
        [vm.gvcID         for vm in vms],      # ElementId object
        [vm.gvcSelectID   for vm in vms],
    ]
else:
    OUT = "No data returned"