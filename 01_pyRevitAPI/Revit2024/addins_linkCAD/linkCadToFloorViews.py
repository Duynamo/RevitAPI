"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import os

# Revit API & Services
clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import * 
clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ISelectionFilter

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Windows Forms for UI
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *
#endregion

#region ___Revit Context
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = uiapp.ActiveUIDocument
#endregion

#region ___Helper Functions
def get_all_floor_plans(doc):
    """Lấy tất cả các ViewPlan không phải là template và sắp xếp theo tên."""
    collector = FilteredElementCollector(doc).OfClass(ViewPlan)
    # Lọc ra các floor plan thực sự, không phải template
    floor_plans = [v for v in collector if not v.IsTemplate and v.ViewType == ViewType.FloorPlan]
    # Sắp xếp theo tên để dễ tìm kiếm trong UI
    floor_plans.sort(key=lambda x: x.Name)
    return floor_plans
#endregion

#region ___Main UI Form
class LinkCadToFloorViewsForm(Form):
    def __init__(self):
        self.filePathMap = {}
        self.last_selected_list = None
        self.InitializeComponent()

    def InitializeComponent(self):
        # --- UI Dimensions ---
        self.Text = "Link CAD to Floor Views"
        self.ClientSize = Size(650, 450)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.Font = Font("Segoe UI", 9)
        self.BackColor = Color.FromArgb(245, 246, 248)

        # --- GroupBox: Floor Views ---
        self._grb_FloorView = GroupBox()
        self._grb_FloorView.Text = "1. Select Floor Views"
        self._grb_FloorView.Location = Point(15, 15)
        self._grb_FloorView.Size = Size(300, 350)
        self.Controls.Add(self._grb_FloorView)

        # --- GroupBox: CAD Files ---
        self._grb_LinkCadPath = GroupBox()
        self._grb_LinkCadPath.Text = "2. Select CAD Files"
        self._grb_LinkCadPath.Location = Point(330, 15)
        self._grb_LinkCadPath.Size = Size(300, 350)
        self.Controls.Add(self._grb_LinkCadPath)

        # --- Controls for Floor Views ---
        self._clb_FloorView = CheckedListBox()
        self._clb_FloorView.DisplayMember = 'Name'
        self._clb_FloorView.Location = Point(15, 60)
        self._clb_FloorView.Size = Size(270, 275)
        self._clb_FloorView.CheckOnClick = True
        self._clb_FloorView.Items.AddRange(System.Array[System.Object](get_all_floor_plans(doc)))
        self._clb_FloorView.SelectedIndexChanged += self.OnListSelectionChanged
        self._grb_FloorView.Controls.Add(self._clb_FloorView)

        self._ckb_AllNoneFloorView = CheckBox()
        self._ckb_AllNoneFloorView.Text = "All / None"
        self._ckb_AllNoneFloorView.Location = Point(15, 30)
        self._ckb_AllNoneFloorView.AutoSize = True
        self._ckb_AllNoneFloorView.CheckedChanged += self.Ckb_AllNoneFloorViewCheckedChanged
        self._grb_FloorView.Controls.Add(self._ckb_AllNoneFloorView)

        # --- Controls for CAD Files ---
        self._btt_LinkCad = Button()
        self._btt_LinkCad.Text = "Browse Files..."
        self._btt_LinkCad.Location = Point(15, 25)
        self._btt_LinkCad.Size = Size(120, 30)
        self._btt_LinkCad.Click += self.Btt_LinkCadClick
        self._grb_LinkCadPath.Controls.Add(self._btt_LinkCad)

        self._btt_Reset = Button()
        self._btt_Reset.Text = "Reset List"
        self._btt_Reset.Location = Point(145, 25)
        self._btt_Reset.Size = Size(120, 30)
        self._btt_Reset.Click += self.Btt_ResetClick
        self._grb_LinkCadPath.Controls.Add(self._btt_Reset)

        self._clb_LinkCad = CheckedListBox()
        self._clb_LinkCad.Location = Point(15, 95)
        self._clb_LinkCad.Size = Size(270, 240)
        self._clb_LinkCad.CheckOnClick = True
        self._clb_LinkCad.SelectedIndexChanged += self.OnListSelectionChanged
        self._grb_LinkCadPath.Controls.Add(self._clb_LinkCad)

        self._ckb_AllNoneCad = CheckBox()
        self._ckb_AllNoneCad.Text = "All / None"
        self._ckb_AllNoneCad.Location = Point(15, 65)
        self._ckb_AllNoneCad.AutoSize = True
        self._ckb_AllNoneCad.CheckedChanged += self.Ckb_AllNoneCadCheckedChanged
        self._grb_LinkCadPath.Controls.Add(self._ckb_AllNoneCad)

        # --- Action Buttons ---
        self._btt_Up = Button()
        self._btt_Up.Text = "Up"
        self._btt_Up.Location = Point(15, 380)
        self._btt_Up.Size = Size(75, 30)
        self._btt_Up.Click += self.Btt_UpClick
        self.Controls.Add(self._btt_Up)

        self._btt_Down = Button()
        self._btt_Down.Text = "Down"
        self._btt_Down.Location = Point(100, 380)
        self._btt_Down.Size = Size(75, 30)
        self._btt_Down.Click += self.Btt_DownClick
        self.Controls.Add(self._btt_Down)

        self._btt_Run = Button()
        self._btt_Run.Text = "RUN"
        self._btt_Run.Font = Font("Segoe UI", 10, FontStyle.Bold)
        self._btt_Run.BackColor = Color.FromArgb(76, 175, 80)
        self._btt_Run.ForeColor = Color.White
        self._btt_Run.Location = Point(420, 380)
        self._btt_Run.Size = Size(100, 35)
        self._btt_Run.Click += self.Btt_RunClick
        self.Controls.Add(self._btt_Run)

        self._btt_Cancle = Button()
        self._btt_Cancle.Text = "CANCEL"
        self._btt_Cancle.Location = Point(530, 380)
        self._btt_Cancle.Size = Size(100, 35)
        self._btt_Cancle.Click += self.Btt_CancleClick
        self.Controls.Add(self._btt_Cancle)

        # --- Copyright Label ---
        self._lbl_Copyright = Label()
        self._lbl_Copyright.Text = "@V2D"
        self._lbl_Copyright.Font = Font("Segoe UI", 8, FontStyle.Bold)
        self._lbl_Copyright.ForeColor = Color.Gray
        self._lbl_Copyright.AutoSize = True
        self._lbl_Copyright.Location = Point(15, 420)
        self.Controls.Add(self._lbl_Copyright)

        self.TopMost = True

    def OnListSelectionChanged(self, sender, e):
        if sender.Focused:
            self.last_selected_list = sender
            if sender == self._clb_FloorView:
                self._clb_LinkCad.SelectedIndex = -1
            else:
                self._clb_FloorView.SelectedIndex = -1

    def Ckb_AllNoneFloorViewCheckedChanged(self, sender, e):
        self.ToggleAllItems(self._clb_FloorView, self._ckb_AllNoneFloorView.Checked)

    def Ckb_AllNoneCadCheckedChanged(self, sender, e):
        self.ToggleAllItems(self._clb_LinkCad, self._ckb_AllNoneCad.Checked)

    def ToggleAllItems(self, checklistbox, state):
        for i in range(checklistbox.Items.Count):
            checklistbox.SetItemChecked(i, state)

    def Btt_LinkCadClick(self, sender, e):
        openDialog = OpenFileDialog()
        openDialog.Multiselect = True
        openDialog.Filter = "DWG Files (*.dwg)|*.dwg|All Files (*.*)|*.*"
        if openDialog.ShowDialog() == DialogResult.OK:
            for filePath in openDialog.FileNames:
                fileName = os.path.basename(filePath)
                if fileName not in self.filePathMap:
                    self._clb_LinkCad.Items.Add(fileName)
                    self.filePathMap[fileName] = filePath
            self._ckb_AllNoneCad.Checked = True

    def Btt_ResetClick(self, sender, e):
        self._clb_LinkCad.Items.Clear()
        self.filePathMap.clear()

    def Btt_UpClick(self, sender, e):
        if self.last_selected_list:
            self.MoveItem(self.last_selected_list, -1)

    def Btt_DownClick(self, sender, e):
        if self.last_selected_list:
            self.MoveItem(self.last_selected_list, 1)

    def MoveItem(self, checklistbox, direction):
        if checklistbox.SelectedItem is None:
            MessageBox.Show("Please select an item to move.", "Information")
            return
        
        index = checklistbox.SelectedIndex
        if direction == -1 and index > 0:
            newIndex = index - 1
        elif direction == 1 and index < checklistbox.Items.Count - 1:
            newIndex = index + 1
        else:
            return

        item = checklistbox.SelectedItem
        isChecked = checklistbox.GetItemChecked(index)
        
        checklistbox.Items.RemoveAt(index)
        checklistbox.Items.Insert(newIndex, item)
        checklistbox.SetItemChecked(newIndex, isChecked)
        checklistbox.SelectedIndex = newIndex

    def Btt_RunClick(self, sender, e):
        floorViews = [item for item in self._clb_FloorView.CheckedItems]
        linkCadNames = [item for item in self._clb_LinkCad.CheckedItems]

        if len(floorViews) != len(linkCadNames):
            MessageBox.Show("The number of selected views must match the number of selected CAD files.", "Error")
            return
        
        if not floorViews:
            MessageBox.Show("Please select at least one view and one CAD file.", "Error")
            return

        linkCads = [self.filePathMap[name] for name in linkCadNames]

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
        try:
            for floorView, linkCad in zip(floorViews, linkCads):
                # Link the CAD file to the specified floor view
                doc.Link(linkCad, options, floorView, linkedElem)
                importInst = doc.GetElement(linkedElem.Value)
                # Unpin the linked instance so it can be moved if needed later
                if importInst:
                    importInst.Pinned = False
            TransactionManager.Instance.TransactionTaskDone()
            MessageBox.Show("Successfully linked {} CAD files.".format(len(linkCads)), "Success")
            self.Close()
        except Exception as ex:
            TransactionManager.Instance.ForceCloseTransaction()
            MessageBox.Show("An error occurred: " + str(ex), "Error")

    def Btt_CancleClick(self, sender, e):
        self.Close()

#endregion

# --- Main Execution ---
try:
    form = LinkCadToFloorViewsForm()
    Application.Run(form)
except Exception as e:
    MessageBox.Show("Failed to launch the tool: " + str(e), "Error")