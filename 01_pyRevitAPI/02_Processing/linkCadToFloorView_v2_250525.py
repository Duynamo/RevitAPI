"""Copyright by: vudinhduybm@gmail.com"""
import clr
import sys 
import System   
import math
import os
clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *
clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import SaveAsOptions
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
from System.Windows.Forms import OpenFileDialog

'''___'''
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
'''___'''

#region _def
def getAllfloors(doc):
    collector = FilteredElementCollector(doc).OfClass(ViewPlan)
    floors = [view for view in collector if not view.IsTemplate]
    # Sắp xếp theo tên (alphabet)
    floors.sort(key=lambda x: x.Name)
    return floors
#endregion

#region _MainForm
allfloorsView = getAllfloors(doc)

class MainForm(Form):
    def __init__(self):
        self.filePathMap = {}  # Initialize filePathMap
        self.last_selected_list = None  # Track last selected list
        self.InitializeComponent()    
    
    def InitializeComponent(self):
        # Lấy kích thước màn hình làm việc
        primary_screen = Screen.PrimaryScreen.WorkingArea
        screen_width = primary_screen.Width
        screen_height = primary_screen.Height

        # Đặt kích thước form: 1/2 chiều rộng và 1/3 chiều cao màn hình
        form_width = int(screen_width * 0.5)
        form_height = int(screen_height * 0.33)
        # Đảm bảo kích thước tối thiểu
        form_width = max(form_width, 600)  # Tối thiểu 600px
        form_height = max(form_height, 300)  # Tối thiểu 300px

        # Tính tỷ lệ để điều chỉnh vị trí và kích thước các thành phần
        width_ratio = form_width / 966.0  # So với kích thước gốc (966)
        height_ratio = form_height / 338.0  # So với kích thước gốc (338)

        self._grb_LinkCadPath = System.Windows.Forms.GroupBox()
        self._btt_LinkCad = System.Windows.Forms.Button()
        self._clb_LinkCad = System.Windows.Forms.CheckedListBox()
        self._ckb_AllNoneCad = System.Windows.Forms.CheckBox()
        self._txb_TotalCad = System.Windows.Forms.TextBox()
        self._lbl_TotalCad = System.Windows.Forms.Label()
        self._btt_Reset = System.Windows.Forms.Button()
        self._grb_FloorView = System.Windows.Forms.GroupBox()
        self._clb_FloorView = System.Windows.Forms.CheckedListBox()
        self._ckb_AllNonefloorView = System.Windows.Forms.CheckBox()
        self._lbl_TotalfloorView = System.Windows.Forms.Label()
        self._txb_TotalView = System.Windows.Forms.TextBox()
        self._btt_Up = System.Windows.Forms.Button()
        self._btt_Down = System.Windows.Forms.Button()
        self._btt_Run = System.Windows.Forms.Button()
        self._btt_Cancle = System.Windows.Forms.Button()
        self._lbl_vitaminD = System.Windows.Forms.Label()
        self._grb_LinkCadPath.SuspendLayout()
        self._grb_FloorView.SuspendLayout()
        self.SuspendLayout()

        # grb_FloorView
        group_box_width = int(form_width * 0.45)
        group_box_height = int(form_height * 0.80)
        self._grb_FloorView.Controls.Add(self._txb_TotalView)
        self._grb_FloorView.Controls.Add(self._lbl_TotalfloorView)
        self._grb_FloorView.Controls.Add(self._ckb_AllNonefloorView)
        self._grb_FloorView.Controls.Add(self._clb_FloorView)
        self._grb_FloorView.Cursor = System.Windows.Forms.Cursors.Default
        self._grb_FloorView.Location = System.Drawing.Point(12, 12)
        self._grb_FloorView.Name = "grb_FloorView"
        self._grb_FloorView.Size = System.Drawing.Size(group_box_width, group_box_height)
        self._grb_FloorView.TabIndex = 1
        self._grb_FloorView.TabStop = False

        # clb_FloorView
        clb_height = int(group_box_height * 0.65)
        self._clb_FloorView.DisplayMember = 'Name'
        self._clb_FloorView.AllowDrop = True
        self._clb_FloorView.BackColor = System.Drawing.SystemColors.MenuBar
        self._clb_FloorView.CheckOnClick = True
        self._clb_FloorView.Font = System.Drawing.Font("Segoe UI", 6, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self._clb_FloorView.FormattingEnabled = True
        self._clb_FloorView.HorizontalScrollbar = True
        self._clb_FloorView.Location = System.Drawing.Point(6, int(86 * height_ratio))
        self._clb_FloorView.Name = "clb_FloorView"
        self._clb_FloorView.Size = System.Drawing.Size(group_box_width - 18, clb_height)
        self._clb_FloorView.Sorted = False
        self._clb_FloorView.TabIndex = 6
        self._clb_FloorView.SelectedIndexChanged += self.Clb_FloorViewSelectedIndexChanged
        self._clb_FloorView.DoubleClick += self.Clb_FloorViewDoubleClick  # Thêm sự kiện DoubleClick
        self._clb_FloorView.Items.AddRange(System.Array[System.Object](allfloorsView))

        # ckb_AllNonefloorView
        self._ckb_AllNonefloorView.AutoSize = True
        self._ckb_AllNonefloorView.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self._ckb_AllNonefloorView.Location = System.Drawing.Point(int(11 * width_ratio), int(58 * height_ratio))
        self._ckb_AllNonefloorView.Name = "ckb_AllNonefloorView"
        self._ckb_AllNonefloorView.Size = System.Drawing.Size(int(89 * width_ratio), int(22 * height_ratio))
        self._ckb_AllNonefloorView.TabIndex = 6
        self._ckb_AllNonefloorView.Text = "All/None"
        self._ckb_AllNonefloorView.UseVisualStyleBackColor = True
        self._ckb_AllNonefloorView.CheckedChanged += self.Ckb_AllNonefloorViewCheckedChanged

        # lbl_TotalfloorView
        self._lbl_TotalfloorView.AutoSize = True
        self._lbl_TotalfloorView.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self._lbl_TotalfloorView.Location = System.Drawing.Point(int(123 * width_ratio), int(58 * height_ratio))
        self._lbl_TotalfloorView.Name = "lbl_TotalfloorView"
        self._lbl_TotalfloorView.Size = System.Drawing.Size(int(85 * width_ratio), int(19 * height_ratio))
        self._lbl_TotalfloorView.TabIndex = 6
        self._lbl_TotalfloorView.Text = "Total View"

        # txb_TotalView
        self._txb_TotalView.BackColor = System.Drawing.SystemColors.Menu
        self._txb_TotalView.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self._txb_TotalView.Location = System.Drawing.Point(int(207 * width_ratio), int(56 * height_ratio))
        self._txb_TotalView.Name = "txb_TotalView"
        self._txb_TotalView.Size = System.Drawing.Size(int(56 * width_ratio), int(24 * height_ratio))
        self._txb_TotalView.TabIndex = 6
        self._txb_TotalView.TextChanged += self.Txb_TotalViewTextChanged

        # grb_LinkCadPath
        linkcad_width = form_width - 12 - (12 + group_box_width + 12)
        self._grb_LinkCadPath.Controls.Add(self._btt_Reset)
        self._grb_LinkCadPath.Controls.Add(self._lbl_TotalCad)
        self._grb_LinkCadPath.Controls.Add(self._txb_TotalCad)
        self._grb_LinkCadPath.Controls.Add(self._ckb_AllNoneCad)
        self._grb_LinkCadPath.Controls.Add(self._clb_LinkCad)
        self._grb_LinkCadPath.Controls.Add(self._btt_LinkCad)
        self._grb_LinkCadPath.Location = System.Drawing.Point(12 + group_box_width + 12, 12)
        self._grb_LinkCadPath.Name = "grb_LinkCadPath"
        self._grb_LinkCadPath.Size = System.Drawing.Size(linkcad_width, group_box_height)
        self._grb_LinkCadPath.TabIndex = 0
        self._grb_LinkCadPath.TabStop = False

        # btt_LinkCad
        linkcad_button_width = int(138 * width_ratio)
        linkcad_button_x = linkcad_width - 6 - linkcad_button_width
        reset_button_width = int(93 * width_ratio)
        reset_button_x = linkcad_button_x - 6 - reset_button_width
        self._btt_LinkCad.AutoSize = True
        self._btt_LinkCad.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_LinkCad.Cursor = System.Windows.Forms.Cursors.AppStarting
        self._btt_LinkCad.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self._btt_LinkCad.ForeColor = System.Drawing.Color.Red
        self._btt_LinkCad.Location = System.Drawing.Point(reset_button_x, int(16 * height_ratio))  # Đặt ở vị trí của Reset
        self._btt_LinkCad.Name = "btt_LinkCad"
        self._btt_LinkCad.Size = System.Drawing.Size(linkcad_button_width, int(29 * height_ratio))
        self._btt_LinkCad.TabIndex = 0
        self._btt_LinkCad.Text = "Select Link Cad"
        self._btt_LinkCad.UseVisualStyleBackColor = False
        self._btt_LinkCad.Click += self.Btt_LinkCadClick

        # btt_Reset
        self._btt_Reset.AutoSize = True
        self._btt_Reset.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_Reset.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self._btt_Reset.ForeColor = System.Drawing.Color.Red
        self._btt_Reset.Location = System.Drawing.Point(linkcad_button_x, int(16 * height_ratio))  # Đặt ở vị trí của LinkCad
        self._btt_Reset.Name = "btt_Reset"
        self._btt_Reset.Size = System.Drawing.Size(reset_button_width, int(29 * height_ratio))
        self._btt_Reset.TabIndex = 5
        self._btt_Reset.Text = "Reset"
        self._btt_Reset.UseVisualStyleBackColor = False
        self._btt_Reset.Click += self.Btt_ResetClick

        # clb_LinkCad
        self._clb_LinkCad.DisplayMember = 'Name'
        self._clb_LinkCad.AllowDrop = True
        self._clb_LinkCad.BackColor = System.Drawing.SystemColors.MenuBar
        self._clb_LinkCad.CheckOnClick = True
        self._clb_LinkCad.Font = System.Drawing.Font("Segoe UI", 6, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self._clb_LinkCad.FormattingEnabled = True
        self._clb_LinkCad.HorizontalScrollbar = True
        self._clb_LinkCad.Location = System.Drawing.Point(6, int(86 * height_ratio))
        self._clb_LinkCad.Name = "clb_LinkCad"
        self._clb_LinkCad.Size = System.Drawing.Size(linkcad_width - 18, clb_height)
        self._clb_LinkCad.TabIndex = 1
        self._clb_LinkCad.SelectedIndexChanged += self.Clb_LinkCadSelectedIndexChanged

        # ckb_AllNoneCad
        self._ckb_AllNoneCad.AutoSize = True
        self._clb_LinkCad.CheckOnClick = True
        self._ckb_AllNoneCad.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self._ckb_AllNoneCad.Location = System.Drawing.Point(int(10 * width_ratio), int(58 * height_ratio))
        self._ckb_AllNoneCad.Name = "ckb_AllNoneCad"
        self._ckb_AllNoneCad.Size = System.Drawing.Size(int(89 * width_ratio), int(22 * height_ratio))
        self._ckb_AllNoneCad.TabIndex = 2
        self._ckb_AllNoneCad.Text = "All/None"
        self._ckb_AllNoneCad.UseVisualStyleBackColor = True
        self._ckb_AllNoneCad.CheckedChanged += self.Ckb_AllNoneCadCheckedChanged

        # txb_TotalCad
        self._txb_TotalCad.BackColor = System.Drawing.SystemColors.Menu
        self._txb_TotalCad.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self._txb_TotalCad.Location = System.Drawing.Point(int(207 * width_ratio), int(56 * height_ratio))
        self._txb_TotalCad.Name = "txb_TotalCad"
        self._txb_TotalCad.Size = System.Drawing.Size(int(56 * width_ratio), int(24 * height_ratio))
        self._txb_TotalCad.TabIndex = 3
        self._txb_TotalCad.TextChanged += self.Txb_TotalCadTextChanged

        # lbl_TotalCad
        self._lbl_TotalCad.AutoSize = True
        self._lbl_TotalCad.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self._lbl_TotalCad.Location = System.Drawing.Point(int(123 * width_ratio), int(58 * height_ratio))
        self._lbl_TotalCad.Name = "lbl_TotalCad"
        self._lbl_TotalCad.Size = System.Drawing.Size(int(78 * width_ratio), int(19 * height_ratio))
        self._lbl_TotalCad.TabIndex = 4
        self._lbl_TotalCad.Text = "Total Cad"

        # btt_Cancle
        button_width = int(85 * width_ratio)
        button_height = int(29 * height_ratio)
        cancle_x = form_width - 12 - button_width
        self._btt_Cancle.AutoSize = True
        self._btt_Cancle.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_Cancle.Cursor = System.Windows.Forms.Cursors.AppStarting
        self._btt_Cancle.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self._btt_Cancle.ForeColor = System.Drawing.Color.Red
        self._btt_Cancle.Location = System.Drawing.Point(cancle_x, form_height - 70)
        self._btt_Cancle.Name = "btt_Cancle"
        self._btt_Cancle.Size = System.Drawing.Size(button_width, button_height)
        self._btt_Cancle.TabIndex = 7
        self._btt_Cancle.Text = "CANCLE"
        self._btt_Cancle.UseVisualStyleBackColor = False
        self._btt_Cancle.Click += self.Btt_CancleClick

        # btt_Run
        run_x = cancle_x - 12 - button_width
        self._btt_Run.AutoSize = True
        self._btt_Run.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_Run.Cursor = System.Windows.Forms.Cursors.AppStarting
        self._btt_Run.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self._btt_Run.ForeColor = System.Drawing.Color.Red
        self._btt_Run.Location = System.Drawing.Point(run_x, form_height - 70)
        self._btt_Run.Name = "btt_Run"
        self._btt_Run.Size = System.Drawing.Size(button_width, button_height)
        self._btt_Run.TabIndex = 6
        self._btt_Run.Text = "RUN"
        self._btt_Run.UseVisualStyleBackColor = False
        self._btt_Run.Click += self.Btt_RunClick

        # btt_Up
        up_width = int(70 * width_ratio)
        up_x = run_x - 12 - up_width
        self._btt_Up.AutoSize = True
        self._btt_Up.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_Up.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self._btt_Up.ForeColor = System.Drawing.Color.Red
        self._btt_Up.Location = System.Drawing.Point(up_x, form_height - 70)
        self._btt_Up.Name = "btt_Up"
        self._btt_Up.Size = System.Drawing.Size(up_width, button_height)
        self._btt_Up.TabIndex = 9
        self._btt_Up.Text = "Up"
        self._btt_Up.UseVisualStyleBackColor = False
        self._btt_Up.Click += self.Btt_UpClick

        # btt_Down
        down_x = up_x - 12 - up_width
        self._btt_Down.AutoSize = True
        self._btt_Down.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_Down.Font = System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self._btt_Down.ForeColor = System.Drawing.Color.Red
        self._btt_Down.Location = System.Drawing.Point(down_x, form_height - 70)
        self._btt_Down.Name = "btt_Down"
        self._btt_Down.Size = System.Drawing.Size(up_width, button_height)
        self._btt_Down.TabIndex = 8
        self._btt_Down.Text = "Down"
        self._btt_Down.UseVisualStyleBackColor = False
        self._btt_Down.Click += self.Btt_DownClick

        # lbl_vitaminD
        self._lbl_vitaminD.AutoSize = True
        self._lbl_vitaminD.Font = System.Drawing.Font("Segoe UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self._lbl_vitaminD.Location = System.Drawing.Point(15, form_height - 40)
        self._lbl_vitaminD.Name = "lbl_vitaminD"
        self._lbl_vitaminD.Size = System.Drawing.Size(int(25 * width_ratio), int(13 * height_ratio))
        self._lbl_vitaminD.TabIndex = 7
        self._lbl_vitaminD.Text = "@D"

        # MainForm
        self.BackColor = System.Drawing.SystemColors.Menu
        self.ClientSize = System.Drawing.Size(form_width, form_height)
        self.Controls.Add(self._grb_FloorView)
        self.Controls.Add(self._grb_LinkCadPath)
        self.Controls.Add(self._lbl_vitaminD)
        self.Controls.Add(self._btt_Down)
        self.Controls.Add(self._btt_Up)
        self.Controls.Add(self._btt_Run)
        self.Controls.Add(self._btt_Cancle)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.Name = "MainForm"
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.Text = "Link Cad To Floor View"
        self.TopMost = True
        self._grb_LinkCadPath.ResumeLayout(False)
        self._grb_LinkCadPath.PerformLayout()
        self._grb_FloorView.ResumeLayout(False)
        self._grb_FloorView.PerformLayout()
        self.ResumeLayout(False)
        self.PerformLayout()

    def Clb_FloorViewDoubleClick(self, sender, e):
        self._clb_FloorView.SelectedIndex = -1  # Bỏ highlight khi double-click
        pass

    def Clb_FloorViewSelectedIndexChanged(self, sender, e):
        varZ = self._clb_FloorView.CheckedItems.Count
        n = 0
        if varZ != 0:
            for i in range(varZ):
                n += 1
                self._txb_TotalView.Text = str(n)
        else:
            self._txb_TotalView.Text = str(0)
        # Cập nhật last_selected_list và bỏ chọn clb_LinkCad
        self.last_selected_list = self._clb_FloorView
        self._clb_LinkCad.SelectedIndex = -1
        pass

    def Ckb_AllNonefloorViewCheckedChanged(self, sender, e):
        var = self._clb_FloorView.Items.Count
        rangers = range(var)
        for i in rangers:
            if self._ckb_AllNonefloorView.Checked == True:
                self._clb_FloorView.SetItemChecked(i, True)
                self._txb_TotalView.Text = str(var)
            else:
                self._clb_FloorView.SetItemChecked(i, False)
                self._txb_TotalView.Text = str(0)
        pass

    def Txb_TotalViewTextChanged(self, sender, e):
        pass

    def Clb_LinkCadSelectedIndexChanged(self, sender, e):
        varZ = self._clb_LinkCad.CheckedItems.Count
        n = 0
        if varZ != 0:
            for i in range(varZ):
                n += 1
                self._txb_TotalCad.Text = str(n)
        else:
            self._txb_TotalCad.Text = str(0)
        # Cập nhật last_selected_list và bỏ chọn clb_FloorView
        self.last_selected_list = self._clb_LinkCad
        self._clb_FloorView.SelectedIndex = -1
        pass

    def Ckb_AllNoneCadCheckedChanged(self, sender, e):
        var = self._clb_LinkCad.Items.Count
        rangers = range(var)
        for i in rangers:
            if self._ckb_AllNoneCad.Checked == True:
                self._clb_LinkCad.SetItemChecked(i, True)
                self._txb_TotalCad.Text = str(var)
            else:
                self._clb_LinkCad.SetItemChecked(i, False)
                self._txb_TotalCad.Text = str(0)
        pass

    def Txb_TotalCadTextChanged(self, sender, e):
        pass

    def move_item_up(self, checklistbox, selected_index, update_count_callback):
        if checklistbox.Items.Count > 1 and selected_index > 0:  # Có ít nhất 2 mục và không phải mục đầu
            # Lưu trạng thái checked và mục
            is_checked = checklistbox.GetItemChecked(selected_index)
            selected_item = checklistbox.Items[selected_index]
            above_item = checklistbox.Items[selected_index - 1]
            above_checked = checklistbox.GetItemChecked(selected_index - 1)
            # Cập nhật filePathMap nếu là clb_LinkCad
            if checklistbox == self._clb_LinkCad:
                selected_path = self.filePathMap.get(str(selected_item))
                above_path = self.filePathMap.get(str(above_item))
                if selected_path and above_path:
                    self.filePathMap[str(selected_item)] = above_path
                    self.filePathMap[str(above_item)] = selected_path
            # Hoán đổi bằng RemoveAt và Insert
            checklistbox.Items.RemoveAt(selected_index)
            checklistbox.Items.Insert(selected_index - 1, selected_item)
            # Cập nhật trạng thái checked
            checklistbox.SetItemChecked(selected_index - 1, is_checked)
            checklistbox.SetItemChecked(selected_index, above_checked)
            # Cập nhật lựa chọn
            checklistbox.SelectedIndex = selected_index - 1
            # Làm mới giao diện
            checklistbox.Refresh()
            # Cập nhật số lượng checked
            update_count_callback(self, System.EventArgs.Empty)
        pass

    def move_item_down(self, checklistbox, selected_index, update_count_callback):
        if checklistbox.Items.Count > 1 and selected_index >= 0 and selected_index < checklistbox.Items.Count - 1:  # Có ít nhất 2 mục và không phải mục cuối
            # Lưu trạng thái checked và mục
            is_checked = checklistbox.GetItemChecked(selected_index)
            selected_item = checklistbox.Items[selected_index]
            below_item = checklistbox.Items[selected_index + 1]
            below_checked = checklistbox.GetItemChecked(selected_index + 1)
            # Cập nhật filePathMap nếu là clb_LinkCad
            if checklistbox == self._clb_LinkCad:
                selected_path = self.filePathMap.get(str(selected_item))
                below_path = self.filePathMap.get(str(below_item))
                if selected_path and below_path:
                    self.filePathMap[str(selected_item)] = below_path
                    self.filePathMap[str(below_item)] = selected_path
            # Hoán đổi bằng RemoveAt và Insert
            checklistbox.Items.RemoveAt(selected_index)
            checklistbox.Items.Insert(selected_index + 1, selected_item)
            # Cập nhật trạng thái checked
            checklistbox.SetItemChecked(selected_index + 1, is_checked)
            checklistbox.SetItemChecked(selected_index, below_checked)
            # Cập nhật lựa chọn
            checklistbox.SelectedIndex = selected_index + 1
            # Làm mới giao diện
            checklistbox.Refresh()
            # Cập nhật số lượng checked
            update_count_callback(self, System.EventArgs.Empty)
        pass

    def Btt_UpClick(self, sender, e):
        # Kiểm tra danh sách nào đang được focus
        if self._clb_FloorView.Focused and self._clb_FloorView.SelectedIndex >= 0:
            self.move_item_up(self._clb_FloorView, self._clb_FloorView.SelectedIndex, self.Clb_FloorViewSelectedIndexChanged)
        elif self._clb_LinkCad.Focused and self._clb_LinkCad.SelectedIndex >= 0:
            self.move_item_up(self._clb_LinkCad, self._clb_LinkCad.SelectedIndex, self.Clb_LinkCadSelectedIndexChanged)
        # Nếu không danh sách nào được focus, dùng last_selected_list
        elif self.last_selected_list == self._clb_FloorView and self._clb_FloorView.SelectedIndex >= 0:
            self.move_item_up(self._clb_FloorView, self._clb_FloorView.SelectedIndex, self.Clb_FloorViewSelectedIndexChanged)
        elif self.last_selected_list == self._clb_LinkCad and self._clb_LinkCad.SelectedIndex >= 0:
            self.move_item_up(self._clb_LinkCad, self._clb_LinkCad.SelectedIndex, self.Clb_LinkCadSelectedIndexChanged)
        else:
            System.Windows.Forms.MessageBox.Show("Vui lòng chọn một mục trong danh sách Floor View hoặc Link CAD.")
        pass

    def Btt_DownClick(self, sender, e):
        # Kiểm tra danh sách nào đang được focus
        if self._clb_FloorView.Focused and self._clb_FloorView.SelectedIndex >= 0:
            self.move_item_down(self._clb_FloorView, self._clb_FloorView.SelectedIndex, self.Clb_FloorViewSelectedIndexChanged)
        elif self._clb_LinkCad.Focused and self._clb_LinkCad.SelectedIndex >= 0:
            self.move_item_down(self._clb_LinkCad, self._clb_LinkCad.SelectedIndex, self.Clb_LinkCadSelectedIndexChanged)
        # Nếu không danh sách nào được focus, dùng last_selected_list
        elif self.last_selected_list == self._clb_FloorView and self._clb_FloorView.SelectedIndex >= 0:
            self.move_item_down(self._clb_FloorView, self._clb_FloorView.SelectedIndex, self.Clb_FloorViewSelectedIndexChanged)
        elif self.last_selected_list == self._clb_LinkCad and self._clb_LinkCad.SelectedIndex >= 0:
            self.move_item_down(self._clb_LinkCad, self._clb_LinkCad.SelectedIndex, self.Clb_LinkCadSelectedIndexChanged)
        else:
            System.Windows.Forms.MessageBox.Show("Vui lòng chọn một mục trong danh sách Floor View hoặc Link CAD.")
        pass

    def Btt_LinkCadClick(self, sender, e):
        openDialog = OpenFileDialog()
        openDialog.Multiselect = True  # Cho phép chọn nhiều file
        openDialog.Filter = "DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
        if openDialog.ShowDialog() == DialogResult.OK:
            filePaths = openDialog.FileNames
            if len(filePaths) > 0:  # Kiểm tra xem có file nào được chọn không
                for filePath in filePaths:
                    fileName = os.path.basename(filePath)
                    self._clb_LinkCad.Items.Add(fileName)
                    self.filePathMap[fileName] = filePath
                self._ckb_AllNoneCad.Checked = True
            else:
                System.Windows.Forms.MessageBox.Show("Không có file nào được chọn.")
        pass

    def Btt_ResetClick(self, sender, e):
        self._clb_LinkCad.Items.Clear()
        self.filePathMap.Clear()
        pass

    def Btt_RunClick(self, sender, e):
        floorViews_count = self._clb_FloorView.CheckedItems.Count
        linkCad_count = self._clb_LinkCad.CheckedItems.Count
        if floorViews_count == linkCad_count:
            floorViews = []
            linkCads = []
            for _floorView in self._clb_FloorView.CheckedItems:
                floorViews.append(_floorView)
            for _linkCadName in self._clb_LinkCad.CheckedItems:
                if _linkCadName in self.filePathMap:
                    linkCads.append(self.filePathMap[_linkCadName])
            for floorView, linkCad in zip(floorViews, linkCads):
                options = DWGImportOptions()
                options.AutoCorrectAlmostVHLines = True
                options.CustomScale = 1
                options.OrientToView = True
                options.ThisViewOnly = False
                options.VisibleLayersOnly = False
                options.ColorMode = ImportColorMode.Preserved
                options.Placement = ImportPlacement.Origin
                options.Unit = ImportUnit.Millimeter
                linkedElem = clr.Reference[ElementId]()
                TransactionManager.Instance.EnsureInTransaction(doc)
                doc.Link(linkCad, options, floorView, linkedElem)
                TransactionManager.Instance.TransactionTaskDone()
                importInst = doc.GetElement(linkedElem.Value)
                TransactionManager.Instance.EnsureInTransaction(doc)
                importInst.Pinned = False
                TransactionManager.Instance.TransactionTaskDone()
        pass

    def Btt_CancleClick(self, sender, e):
        self.Close()
        pass
#endregion

f = MainForm()
Application.Run(f)