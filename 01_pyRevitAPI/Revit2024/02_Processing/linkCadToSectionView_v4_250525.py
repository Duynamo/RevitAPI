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
from System.Windows.Forms import OpenFileDialog

'''___'''
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
'''___'''

#region _def
def getAllSections(doc):
    collector = FilteredElementCollector(doc).OfClass(ViewSection)
    sections = [view for view in collector if not view.IsTemplate]
    return sections

def pickPoints(doc):
    TransactionManager.Instance.EnsureInTransaction(doc)
    activeView = doc.ActiveView
    iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
    sketchPlane = SketchPlane.Create(doc, iRefPlane)
    doc.ActiveView.SketchPlane = sketchPlane
    condition = True
    pointsList = []
    dynPList = []
    n = 0
    msg = "Pick Point, hit ESC when finished."
    TaskDialog.Show("^------^", msg)
    while condition:
        try:
            pt = uidoc.Selection.PickPoint()
            n += 1
            pointsList.append(pt)                
        except:
            condition = False
    doc.Delete(sketchPlane.Id)    
    for j in pointsList:
        dynP = j.ToPoint()
        dynPList.append(dynP)
    TransactionManager.Instance.TransactionTaskDone()            
    return pointsList[0]
#endregion

#region _MainForm
allSectionsView = getAllSections(doc)

class MainForm(Form):
    def __init__(self):
        self.filePathMap = {}  # Initialize filePathMap
        self.last_selected_list = None  # Track last selected list
        self.InitializeComponent()    
    
    def InitializeComponent(self):
        # Lấy kích thước màn hình chính (WorkingArea)
        primary_screen = Screen.PrimaryScreen.WorkingArea
        screen_width = primary_screen.Width
        screen_height = primary_screen.Height

        # Đặt kích thước form: 1/3 chiều rộng và 1/2 chiều cao của WorkingArea
        form_width = int(screen_width * 1/2)
        form_height = int(screen_height * 1/3)
        # Đảm bảo kích thước tối thiểu
        form_width = max(form_width, 800)  # Tối thiểu 800px
        form_height = max(form_height, 400)  # Tối thiểu 400px

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
        self._grb_SectionView = System.Windows.Forms.GroupBox()
        self._clb_SectionView = System.Windows.Forms.CheckedListBox()
        self._ckb_AllNoneSectionView = System.Windows.Forms.CheckBox()
        self._lbl_TotalSecTionView = System.Windows.Forms.Label()
        self._txb_TotalView = System.Windows.Forms.TextBox()
        self._btt_Up = System.Windows.Forms.Button()
        self._btt_Down = System.Windows.Forms.Button()
        self._btt_Run = System.Windows.Forms.Button()
        self._btt_Cancle = System.Windows.Forms.Button()
        self._lbl_vitaminD = System.Windows.Forms.Label()
        self._grb_SectionPoint = System.Windows.Forms.GroupBox()
        self._btt_ResetPickPoints = System.Windows.Forms.Button()
        self._lbb_TotalPoints = System.Windows.Forms.Label()
        self._txb_ToTalPoints = System.Windows.Forms.TextBox()
        self._ckb_AllNonePoints = System.Windows.Forms.CheckBox()
        self._clb_SectionPoints = System.Windows.Forms.CheckedListBox()
        self._btt_PickPoints = System.Windows.Forms.Button()
        self._grb_LinkCadPath.SuspendLayout()
        self._grb_SectionView.SuspendLayout()
        self._grb_SectionPoint.SuspendLayout()
        self.SuspendLayout()

        # grb_SectionView
        group_box_width = int(275 * width_ratio)
        group_box_height = int(270 * height_ratio)
        self._grb_SectionView.Location = System.Drawing.Point(int(9 * width_ratio), int(12 * height_ratio))
        self._grb_SectionView.Size = System.Drawing.Size(group_box_width, group_box_height)
        self._grb_SectionView.Controls.Add(self._txb_TotalView)
        self._grb_SectionView.Controls.Add(self._lbl_TotalSecTionView)
        self._grb_SectionView.Controls.Add(self._ckb_AllNoneSectionView)
        self._grb_SectionView.Controls.Add(self._clb_SectionView)
        self._grb_SectionView.Name = "grb_SectionView"
        self._grb_SectionView.TabIndex = 1
        self._grb_SectionView.TabStop = False

        # clb_SectionView
        clb_height = int(169 * height_ratio)
        self._clb_SectionView.DisplayMember = 'Name'
        self._clb_SectionView.AllowDrop = True
        self._clb_SectionView.BackColor = System.Drawing.SystemColors.MenuBar
        self._clb_SectionView.CheckOnClick = True
        self._clb_SectionView.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._clb_SectionView.FormattingEnabled = True
        self._clb_SectionView.HorizontalScrollbar = True
        self._clb_SectionView.Location = System.Drawing.Point(int(6 * width_ratio), int(86 * height_ratio))
        self._clb_SectionView.Size = System.Drawing.Size(int(257 * width_ratio), clb_height)
        self._clb_SectionView.Sorted = True
        self._clb_SectionView.Name = "clb_SectionView"
        self._clb_SectionView.TabIndex = 6
        self._clb_SectionView.SelectedIndexChanged += self.Clb_SectionViewSelectedIndexChanged
        self._clb_SectionView.Items.AddRange(System.Array[System.Object](allSectionsView))        
        self._clb_SectionView.SelectedIndex = 0

        # ckb_AllNoneSectionView
        self._ckb_AllNoneSectionView.AutoSize = True
        self._ckb_AllNoneSectionView.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._ckb_AllNoneSectionView.Location = System.Drawing.Point(int(11 * width_ratio), int(58 * height_ratio))
        self._ckb_AllNoneSectionView.Size = System.Drawing.Size(int(89 * width_ratio), int(22 * height_ratio))
        self._ckb_AllNoneSectionView.Name = "ckb_AllNoneSectionView"
        self._ckb_AllNoneSectionView.TabIndex = 6
        self._ckb_AllNoneSectionView.Text = "All/None"
        self._ckb_AllNoneSectionView.UseVisualStyleBackColor = True
        self._ckb_AllNoneSectionView.CheckedChanged += self.Ckb_AllNoneSectionViewCheckedChanged

        # lbl_TotalSecTionView
        self._lbl_TotalSecTionView.AutoSize = True
        self._lbl_TotalSecTionView.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._lbl_TotalSecTionView.Location = System.Drawing.Point(int(123 * width_ratio), int(58 * height_ratio))
        self._lbl_TotalSecTionView.Size = System.Drawing.Size(int(85 * width_ratio), int(19 * height_ratio))
        self._lbl_TotalSecTionView.Name = "lbl_TotalSecTionView"
        self._lbl_TotalSecTionView.TabIndex = 6
        self._lbl_TotalSecTionView.Text = "Total View"

        # txb_TotalView
        self._txb_TotalView.BackColor = System.Drawing.SystemColors.Menu
        self._txb_TotalView.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._txb_TotalView.Location = System.Drawing.Point(int(207 * width_ratio), int(56 * height_ratio))
        self._txb_TotalView.Size = System.Drawing.Size(int(56 * width_ratio), int(24 * height_ratio))
        self._txb_TotalView.Name = "txb_TotalView"
        self._txb_TotalView.TabIndex = 6
        self._txb_TotalView.TextChanged += self.Txb_TotalViewTextChanged

        # grb_SectionPoint
        self._grb_SectionPoint.Controls.Add(self._btt_ResetPickPoints)
        self._grb_SectionPoint.Controls.Add(self._lbb_TotalPoints)
        self._grb_SectionPoint.Controls.Add(self._txb_ToTalPoints)
        self._grb_SectionPoint.Controls.Add(self._ckb_AllNonePoints)
        self._grb_SectionPoint.Controls.Add(self._clb_SectionPoints)
        self._grb_SectionPoint.Controls.Add(self._btt_PickPoints)
        self._grb_SectionPoint.Location = System.Drawing.Point(int(301 * width_ratio), int(12 * height_ratio))
        self._grb_SectionPoint.Size = System.Drawing.Size(int(318 * width_ratio), group_box_height)
        self._grb_SectionPoint.Name = "grb_SectionPoint"
        self._grb_SectionPoint.TabIndex = 6
        self._grb_SectionPoint.TabStop = False

        # btt_PickPoints
        self._btt_PickPoints.AutoSize = True
        self._btt_PickPoints.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_PickPoints.Cursor = System.Windows.Forms.Cursors.AppStarting
        self._btt_PickPoints.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_PickPoints.ForeColor = System.Drawing.Color.Red
        self._btt_PickPoints.Location = System.Drawing.Point(int(6 * width_ratio), int(16 * height_ratio))
        self._btt_PickPoints.Size = System.Drawing.Size(int(138 * width_ratio), int(29 * height_ratio))
        self._btt_PickPoints.Name = "btt_PickPoints"
        self._btt_PickPoints.TabIndex = 0
        self._btt_PickPoints.Text = "Pick Points"
        self._btt_PickPoints.UseVisualStyleBackColor = False
        self._btt_PickPoints.Click += self.Btt_PickPointsClick

        # clb_SectionPoints
        self._clb_SectionPoints.AllowDrop = True
        self._clb_SectionPoints.BackColor = System.Drawing.SystemColors.MenuBar
        self._clb_SectionPoints.CheckOnClick = True
        self._clb_SectionPoints.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._clb_SectionPoints.FormattingEnabled = True
        self._clb_SectionPoints.HorizontalScrollbar = True
        self._clb_SectionPoints.Location = System.Drawing.Point(int(6 * width_ratio), int(86 * height_ratio))
        self._clb_SectionPoints.Size = System.Drawing.Size(int(306 * width_ratio), clb_height)
        self._clb_SectionPoints.Name = "clb_SectionPoints"
        self._clb_SectionPoints.TabIndex = 1
        self._clb_SectionPoints.SelectedIndexChanged += self.Clb_SectionPointsSelectedIndexChanged

        # ckb_AllNonePoints
        self._ckb_AllNonePoints.AutoSize = True
        self._ckb_AllNonePoints.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._ckb_AllNonePoints.Location = System.Drawing.Point(int(10 * width_ratio), int(58 * height_ratio))
        self._ckb_AllNonePoints.Size = System.Drawing.Size(int(89 * width_ratio), int(22 * height_ratio))
        self._ckb_AllNonePoints.Name = "ckb_AllNonePoints"
        self._ckb_AllNonePoints.TabIndex = 2
        self._ckb_AllNonePoints.Text = "All/None"
        self._ckb_AllNonePoints.UseVisualStyleBackColor = True
        self._ckb_AllNonePoints.CheckedChanged += self.Ckb_AllNonePointsCheckedChanged

        # txb_ToTalPoints
        self._txb_ToTalPoints.BackColor = System.Drawing.SystemColors.Menu
        self._txb_ToTalPoints.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._txb_ToTalPoints.Location = System.Drawing.Point(int(256 * width_ratio), int(56 * height_ratio))
        self._txb_ToTalPoints.Size = System.Drawing.Size(int(56 * width_ratio), int(24 * height_ratio))
        self._txb_ToTalPoints.Name = "txb_ToTalPoints"
        self._txb_ToTalPoints.TabIndex = 3
        self._txb_ToTalPoints.TextChanged += self.Txb_ToTalPointsTextChanged

        # lbb_TotalPoints
        self._lbb_TotalPoints.AutoSize = True
        self._lbb_TotalPoints.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._lbb_TotalPoints.Location = System.Drawing.Point(int(162 * width_ratio), int(58 * height_ratio))
        self._lbb_TotalPoints.Size = System.Drawing.Size(int(94 * width_ratio), int(19 * height_ratio))
        self._lbb_TotalPoints.Name = "lbb_TotalPoints"
        self._lbb_TotalPoints.TabIndex = 4
        self._lbb_TotalPoints.Text = "Total Points"

        # btt_ResetPickPoints
        self._btt_ResetPickPoints.AutoSize = True
        self._btt_ResetPickPoints.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_ResetPickPoints.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_ResetPickPoints.ForeColor = System.Drawing.Color.Red
        self._btt_ResetPickPoints.Location = System.Drawing.Point(int(225 * width_ratio), int(16 * height_ratio))
        self._btt_ResetPickPoints.Size = System.Drawing.Size(int(93 * width_ratio), int(29 * height_ratio))
        self._btt_ResetPickPoints.Name = "btt_ResetPickPoints"
        self._btt_ResetPickPoints.TabIndex = 5
        self._btt_ResetPickPoints.Text = "Reset"
        self._btt_ResetPickPoints.UseVisualStyleBackColor = False
        self._btt_ResetPickPoints.Click += self.Btt_ResetPickPointsClick

        # grb_LinkCadPath
        self._grb_LinkCadPath.Controls.Add(self._btt_Reset)
        self._grb_LinkCadPath.Controls.Add(self._lbl_TotalCad)
        self._grb_LinkCadPath.Controls.Add(self._txb_TotalCad)
        self._grb_LinkCadPath.Controls.Add(self._ckb_AllNoneCad)
        self._grb_LinkCadPath.Controls.Add(self._clb_LinkCad)
        self._grb_LinkCadPath.Controls.Add(self._btt_LinkCad)
        self._grb_LinkCadPath.Location = System.Drawing.Point(int(636 * width_ratio), int(12 * height_ratio))
        self._grb_LinkCadPath.Size = System.Drawing.Size(int(318 * width_ratio), group_box_height)
        self._grb_LinkCadPath.Name = "grb_LinkCadPath"
        self._grb_LinkCadPath.TabIndex = 0
        self._grb_LinkCadPath.TabStop = False

        # btt_LinkCad
        self._btt_LinkCad.AutoSize = True
        self._btt_LinkCad.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_LinkCad.Cursor = System.Windows.Forms.Cursors.AppStarting
        self._btt_LinkCad.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_LinkCad.ForeColor = System.Drawing.Color.Red
        self._btt_LinkCad.Location = System.Drawing.Point(int(6 * width_ratio), int(16 * height_ratio))
        self._btt_LinkCad.Size = System.Drawing.Size(int(138 * width_ratio), int(29 * height_ratio))
        self._btt_LinkCad.Name = "btt_LinkCad"
        self._btt_LinkCad.TabIndex = 0
        self._btt_LinkCad.Text = "Select Link Cad"
        self._btt_LinkCad.UseVisualStyleBackColor = False
        self._btt_LinkCad.Click += self.Btt_LinkCadClick

        # clb_LinkCad
        self._clb_LinkCad.DisplayMember = 'Name'
        self._clb_LinkCad.AllowDrop = True
        self._clb_LinkCad.BackColor = System.Drawing.SystemColors.MenuBar
        self._clb_LinkCad.CheckOnClick = True
        self._clb_LinkCad.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._clb_LinkCad.FormattingEnabled = True
        self._clb_LinkCad.HorizontalScrollbar = True
        self._clb_LinkCad.Location = System.Drawing.Point(int(6 * width_ratio), int(86 * height_ratio))
        self._clb_LinkCad.Size = System.Drawing.Size(int(306 * width_ratio), clb_height)
        self._clb_LinkCad.Name = "clb_LinkCad"
        self._clb_LinkCad.TabIndex = 1
        self._clb_LinkCad.SelectedIndexChanged += self.Clb_LinkCadSelectedIndexChanged

        # ckb_AllNoneCad
        self._ckb_AllNoneCad.AutoSize = True
        self._ckb_AllNoneCad.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._ckb_AllNoneCad.Location = System.Drawing.Point(int(10 * width_ratio), int(58 * height_ratio))
        self._ckb_AllNoneCad.Size = System.Drawing.Size(int(89 * width_ratio), int(22 * height_ratio))
        self._ckb_AllNoneCad.Name = "ckb_AllNoneCad"
        self._ckb_AllNoneCad.TabIndex = 2
        self._ckb_AllNoneCad.Text = "All/None"
        self._ckb_AllNoneCad.UseVisualStyleBackColor = True
        self._ckb_AllNoneCad.CheckedChanged += self.Ckb_AllNoneCadCheckedChanged

        # txb_TotalCad
        self._txb_TotalCad.BackColor = System.Drawing.SystemColors.Menu
        self._txb_TotalCad.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._txb_TotalCad.Location = System.Drawing.Point(int(256 * width_ratio), int(56 * height_ratio))
        self._txb_TotalCad.Size = System.Drawing.Size(int(56 * width_ratio), int(24 * height_ratio))
        self._txb_TotalCad.Name = "txb_TotalCad"
        self._txb_TotalCad.TabIndex = 3
        self._txb_TotalCad.TextChanged += self.Txb_TotalCadTextChanged

        # lbl_TotalCad
        self._lbl_TotalCad.AutoSize = True
        self._lbl_TotalCad.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
        self._lbl_TotalCad.Location = System.Drawing.Point(int(178 * width_ratio), int(58 * height_ratio))
        self._lbl_TotalCad.Size = System.Drawing.Size(int(78 * width_ratio), int(19 * height_ratio))
        self._lbl_TotalCad.Name = "lbl_TotalCad"
        self._lbl_TotalCad.TabIndex = 4
        self._lbl_TotalCad.Text = "Total Cad"

        # btt_Reset
        self._btt_Reset.AutoSize = True
        self._btt_Reset.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_Reset.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_Reset.ForeColor = System.Drawing.Color.Red
        self._btt_Reset.Location = System.Drawing.Point(int(225 * width_ratio), int(16 * height_ratio))
        self._btt_Reset.Size = System.Drawing.Size(int(93 * width_ratio), int(29 * height_ratio))
        self._btt_Reset.Name = "btt_Reset"
        self._btt_Reset.TabIndex = 5
        self._btt_Reset.Text = "Reset"
        self._btt_Reset.UseVisualStyleBackColor = False
        self._btt_Reset.Click += self.Btt_ResetClick

        # btt_Run
        button_width = int(85 * width_ratio)
        button_height = int(29 * height_ratio)
        cancle_x = int(869 * width_ratio)
        self._btt_Run.AutoSize = True
        self._btt_Run.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_Run.Cursor = System.Windows.Forms.Cursors.AppStarting
        self._btt_Run.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_Run.ForeColor = System.Drawing.Color.Red
        self._btt_Run.Location = System.Drawing.Point(int(778 * width_ratio), int(297 * height_ratio))
        self._btt_Run.Size = System.Drawing.Size(button_width, button_height)
        self._btt_Run.Name = "btt_Run"
        self._btt_Run.TabIndex = 6
        self._btt_Run.Text = "RUN"
        self._btt_Run.UseVisualStyleBackColor = False
        self._btt_Run.Click += self.Btt_RunClick

        # btt_Cancle
        self._btt_Cancle.AutoSize = True
        self._btt_Cancle.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_Cancle.Cursor = System.Windows.Forms.Cursors.AppStarting
        self._btt_Cancle.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_Cancle.ForeColor = System.Drawing.Color.Red
        self._btt_Cancle.Location = System.Drawing.Point(cancle_x, int(297 * height_ratio))
        self._btt_Cancle.Size = System.Drawing.Size(button_width, button_height)
        self._btt_Cancle.Name = "btt_Cancle"
        self._btt_Cancle.TabIndex = 7
        self._btt_Cancle.Text = "CANCLE"
        self._btt_Cancle.UseVisualStyleBackColor = False
        self._btt_Cancle.Click += self.Btt_CancleClick

        # btt_Up
        up_width = int(70 * width_ratio)
        run_x = int(778 * width_ratio)
        up_x = run_x - 12 - up_width
        self._btt_Up.AutoSize = True
        self._btt_Up.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_Up.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_Up.ForeColor = System.Drawing.Color.Red
        self._btt_Up.Location = System.Drawing.Point(up_x, int(297 * height_ratio))
        self._btt_Up.Size = System.Drawing.Size(up_width, button_height)
        self._btt_Up.Name = "btt_Up"
        self._btt_Up.TabIndex = 8
        self._btt_Up.Text = "Up"
        self._btt_Up.UseVisualStyleBackColor = False
        self._btt_Up.Click += self.Btt_UpClick

        # btt_Down
        down_x = up_x - 12 - up_width
        self._btt_Down.AutoSize = True
        self._btt_Down.BackColor = System.Drawing.SystemColors.ButtonHighlight
        self._btt_Down.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._btt_Down.ForeColor = System.Drawing.Color.Red
        self._btt_Down.Location = System.Drawing.Point(down_x, int(297 * height_ratio))
        self._btt_Down.Size = System.Drawing.Size(up_width, button_height)
        self._btt_Down.Name = "btt_Down"
        self._btt_Down.TabIndex = 9
        self._btt_Down.Text = "Down"
        self._btt_Down.UseVisualStyleBackColor = False
        self._btt_Down.Click += self.Btt_DownClick

        # lbl_vitaminD
        self._lbl_vitaminD.AutoSize = True
        self._lbl_vitaminD.Font = System.Drawing.Font("Meiryo UI", 6, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
        self._lbl_vitaminD.Location = System.Drawing.Point(int(15 * width_ratio), int(315 * height_ratio))
        self._lbl_vitaminD.Size = System.Drawing.Size(int(25 * width_ratio), int(13 * height_ratio))
        self._lbl_vitaminD.Name = "lbl_vitaminD"
        self._lbl_vitaminD.TabIndex = 7
        self._lbl_vitaminD.Text = "@D"

        # MainForm
        self.BackColor = System.Drawing.SystemColors.Menu
        self.ClientSize = System.Drawing.Size(form_width, form_height)
        self.Controls.Add(self._grb_SectionPoint)
        self.Controls.Add(self._lbl_vitaminD)
        self.Controls.Add(self._btt_Down)
        self.Controls.Add(self._btt_Up)
        self.Controls.Add(self._btt_Cancle)
        self.Controls.Add(self._btt_Run)
        self.Controls.Add(self._grb_SectionView)
        self.Controls.Add(self._grb_LinkCadPath)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.Name = "MainForm"
        self.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        self.Text = "Link Cad To Section View_v2"
        self.TopMost = True
        self._grb_LinkCadPath.ResumeLayout(False)
        self._grb_LinkCadPath.PerformLayout()
        self._grb_SectionView.ResumeLayout(False)
        self._grb_SectionView.PerformLayout()
        self._grb_SectionPoint.ResumeLayout(False)
        self._grb_SectionPoint.PerformLayout()
        self.ResumeLayout(False)
        self.PerformLayout()

    def Clb_SectionViewSelectedIndexChanged(self, sender, e):
        varZ = self._clb_SectionView.CheckedItems.Count
        n = 0
        if varZ != 0:
            for i in range(varZ):
                n += 1
                self._txb_TotalView.Text = str(n)        
        else:
            self._txb_TotalView.Text = str(0)
        # Cập nhật last_selected_list và bỏ chọn các danh sách khác
        self.last_selected_list = self._clb_SectionView
        self._clb_SectionPoints.SelectedIndex = -1
        self._clb_LinkCad.SelectedIndex = -1
        pass        

    def Ckb_AllNoneSectionViewCheckedChanged(self, sender, e):
        var = self._clb_SectionView.Items.Count
        rangers = range(var)
        for i in rangers:
            if self._ckb_AllNoneSectionView.Checked == True:
                self._clb_SectionView.SetItemChecked(i, True)
                self._txb_TotalView.Text = str(var)
            else:
                self._clb_SectionView.SetItemChecked(i, False)
                self._txb_TotalView.Text = str(0)
        pass            

    def Txb_TotalViewTextChanged(self, sender, e):
        pass    

    def Btt_LinkCadClick(self, sender, e):    
        openDialog = OpenFileDialog()
        openDialog.Multiselect = True  # Cho phép chọn nhiều file
        openDialog.Filter = "DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
        if openDialog.ShowDialog() == DialogResult.OK:
            filePaths = openDialog.FileNames
            if len(filePaths) > 0:  # Kiểm tra xem có file nào được chọn
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

    def Clb_LinkCadSelectedIndexChanged(self, sender, e):
        varZ = self._clb_LinkCad.CheckedItems.Count
        n = 0
        if varZ != 0:
            for i in range(varZ):
                n += 1
                self._txb_TotalCad.Text = str(n)        
        else:
            self._txb_TotalCad.Text = str(0)
        # Cập nhật last_selected_list và bỏ chọn các danh sách khác
        self.last_selected_list = self._clb_LinkCad
        self._clb_SectionView.SelectedIndex = -1
        self._clb_SectionPoints.SelectedIndex = -1
        pass

    def Txb_TotalCadTextChanged(self, sender, e):
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

    def Btt_PickPointsClick(self, sender, e):
        TransactionManager.Instance.EnsureInTransaction(doc)
        activeView = doc.ActiveView
        iRefPlane = Plane.CreateByNormalAndOrigin(activeView.ViewDirection, activeView.Origin)
        sketchPlane = SketchPlane.Create(doc, iRefPlane)
        doc.ActiveView.SketchPlane = sketchPlane
        condition = True
        sectionPointsList = []
        n = 0
        msg = "Pick Points on Current Section plane, hit ESC when finished."
        TaskDialog.Show("^------^", msg)
        while condition:
            try:
                revitPoint = uidoc.Selection.PickPoint()
                n += 1
                sectionPointsList.append(revitPoint)                
            except:
                condition = False
        doc.Delete(sketchPlane.Id)    
        for p in sectionPointsList:
            dynPoint = p.ToPoint()
            self._clb_SectionPoints.Items.Add(dynPoint)
        self._ckb_AllNonePoints.Checked = True    
        TransactionManager.Instance.TransactionTaskDone()    
        pass

    def Btt_ResetPickPointsClick(self, sender, e):
        self._clb_SectionPoints.Items.Clear()
        pass

    def Ckb_AllNonePointsCheckedChanged(self, sender, e):
        var = self._clb_SectionPoints.Items.Count
        rangers = range(var)
        for i in rangers:
            if self._ckb_AllNonePoints.Checked == True:
                self._clb_SectionPoints.SetItemChecked(i, True)
                self._txb_ToTalPoints.Text = str(var)
            else:
                self._clb_SectionPoints.SetItemChecked(i, False)
                self._txb_ToTalPoints.Text = str(0)        
        pass

    def Txb_ToTalPointsTextChanged(self, sender, e):
        pass

    def Clb_SectionPointsSelectedIndexChanged(self, sender, e):
        varZ = self._clb_SectionPoints.CheckedItems.Count
        n = 0
        if varZ != 0:
            for i in range(varZ):
                n += 1
                self._txb_ToTalPoints.Text = str(n)        
        else:
            self._txb_ToTalPoints.Text = str(0)
        # Cập nhật last_selected_list và bỏ chọn các danh sách khác
        self.last_selected_list = self._clb_SectionPoints
        self._clb_SectionView.SelectedIndex = -1
        self._clb_LinkCad.SelectedIndex = -1
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
        if self._clb_SectionView.Focused and self._clb_SectionView.SelectedIndex >= 0:
            self.move_item_up(self._clb_SectionView, self._clb_SectionView.SelectedIndex, self.Clb_SectionViewSelectedIndexChanged)
        elif self._clb_SectionPoints.Focused and self._clb_SectionPoints.SelectedIndex >= 0:
            self.move_item_up(self._clb_SectionPoints, self._clb_SectionPoints.SelectedIndex, self.Clb_SectionPointsSelectedIndexChanged)
        elif self._clb_LinkCad.Focused and self._clb_LinkCad.SelectedIndex >= 0:
            self.move_item_up(self._clb_LinkCad, self._clb_LinkCad.SelectedIndex, self.Clb_LinkCadSelectedIndexChanged)
        # Nếu không danh sách nào được focus, dùng last_selected_list
        elif self.last_selected_list == self._clb_SectionView and self._clb_SectionView.SelectedIndex >= 0:
            self.move_item_up(self._clb_SectionView, self._clb_SectionView.SelectedIndex, self.Clb_SectionViewSelectedIndexChanged)
        elif self.last_selected_list == self._clb_SectionPoints and self._clb_SectionPoints.SelectedIndex >= 0:
            self.move_item_up(self._clb_SectionPoints, self._clb_SectionPoints.SelectedIndex, self.Clb_SectionPointsSelectedIndexChanged)
        elif self.last_selected_list == self._clb_LinkCad and self._clb_LinkCad.SelectedIndex >= 0:
            self.move_item_up(self._clb_LinkCad, self._clb_LinkCad.SelectedIndex, self.Clb_LinkCadSelectedIndexChanged)
        else:
            System.Windows.Forms.MessageBox.Show("Vui lòng chọn một mục trong danh sách Section View, Section Points hoặc Link CAD.")
        pass

    def Btt_DownClick(self, sender, e):
        # Kiểm tra danh sách nào đang được focus
        if self._clb_SectionView.Focused and self._clb_SectionView.SelectedIndex >= 0:
            self.move_item_down(self._clb_SectionView, self._clb_SectionView.SelectedIndex, self.Clb_SectionViewSelectedIndexChanged)
        elif self._clb_SectionPoints.Focused and self._clb_SectionPoints.SelectedIndex >= 0:
            self.move_item_down(self._clb_SectionPoints, self._clb_SectionPoints.SelectedIndex, self.Clb_SectionPointsSelectedIndexChanged)
        elif self._clb_LinkCad.Focused and self._clb_LinkCad.SelectedIndex >= 0:
            self.move_item_down(self._clb_LinkCad, self._clb_LinkCad.SelectedIndex, self.Clb_LinkCadSelectedIndexChanged)
        # Nếu không danh sách nào được focus, dùng last_selected_list
        elif self.last_selected_list == self._clb_SectionView and self._clb_SectionView.SelectedIndex >= 0:
            self.move_item_down(self._clb_SectionView, self._clb_SectionView.SelectedIndex, self.Clb_SectionViewSelectedIndexChanged)
        elif self.last_selected_list == self._clb_SectionPoints and self._clb_SectionPoints.SelectedIndex >= 0:
            self.move_item_down(self._clb_SectionPoints, self._clb_SectionPoints.SelectedIndex, self.Clb_SectionPointsSelectedIndexChanged)
        elif self.last_selected_list == self._clb_LinkCad and self._clb_LinkCad.SelectedIndex >= 0:
            self.move_item_down(self._clb_LinkCad, self._clb_LinkCad.SelectedIndex, self.Clb_LinkCadSelectedIndexChanged)
        else:
            System.Windows.Forms.MessageBox.Show("Vui lòng chọn một mục trong danh sách Section View, Section Points hoặc Link CAD.")
        pass

    def Btt_RunClick(self, sender, e):
        sectionViews_count = self._clb_SectionView.CheckedItems.Count
        sectionViewsPoints_count = self._clb_SectionPoints.CheckedItems.Count
        linkCad_count = self._clb_LinkCad.CheckedItems.Count
        if sectionViews_count == sectionViewsPoints_count == linkCad_count:
            sectionViews = []
            sectionPoints = []
            linkCads = []
            for _sectionView in self._clb_SectionView.CheckedItems:
                sectionViews.append(_sectionView)
            for _sectionPoint in self._clb_SectionPoints.CheckedItems:
                sectionPoints.append(_sectionPoint)
            for _linkCadName in self._clb_LinkCad.CheckedItems:
                if _linkCadName in self.filePathMap:
                    linkCads.append(self.filePathMap[_linkCadName])                
            for sectionView, sectionPoint, linkCad in zip(sectionViews, sectionPoints, linkCads):
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
                doc.Link(linkCad, options, sectionView, linkedElem)
                TransactionManager.Instance.TransactionTaskDone()
                importInst = doc.GetElement(linkedElem.Value)
                TransactionManager.Instance.EnsureInTransaction(doc)
                importInst.Pinned = False
                TransactionManager.Instance.TransactionTaskDone()
                CADLink = doc.GetElement(importInst.GetTypeId())
                pointOrigin = XYZ(0,0,0)
                moveVector = sectionPoint.ToRevitType() - pointOrigin     
                TransactionManager.Instance.EnsureInTransaction(doc)
                elementTransform = Transform.CreateTranslation(moveVector)
                importInst.Location.Move(moveVector)
                TransactionManager.Instance.TransactionTaskDone()    
        pass

    def Btt_CancleClick(self, sender, e):
        self.Close()
        pass
#endregion

f = MainForm()
Application.Run(f)
