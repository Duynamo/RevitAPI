import clr
import sys
import System
import os

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
clr.AddReference("RevitAPIUI")
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

"""_______________________________________________________________________________________"""
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app   = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

#region__function
def getAllSheets(doc):
    collector = (FilteredElementCollector(doc)
                 .OfClass(ViewSheet)
                 .ToElements())
    sheets = sorted(collector, key=lambda s: s.SheetNumber)
    return sheets
#endregion

#region__UI
class SheetEditorForm(Form):
    def __init__(self):
        self._sheets    = getAllSheets(doc)
        self._isRunning = False
        self.InitializeComponent()
        self._loadSheets()

    def InitializeComponent(self):
        self._dgv         = System.Windows.Forms.DataGridView()
        self._btt_RUN     = System.Windows.Forms.Button()
        self._btt_CLEAR   = System.Windows.Forms.Button()
        self._btt_CANCLE  = System.Windows.Forms.Button()
        self._lb_status   = System.Windows.Forms.Label()
        self._lb_title    = System.Windows.Forms.Label()
        self._lb_hint     = System.Windows.Forms.Label()
        self._lb_FVC      = System.Windows.Forms.Label()
        self._progressBar = System.Windows.Forms.ProgressBar()

        self.SuspendLayout()

        # ── lb_title ─────────────────────────────────────────
        self._lb_title.Text      = "Batch Edit Sheet Number & Name"
        self._lb_title.Font      = System.Drawing.Font("Meiryo UI", 11, System.Drawing.FontStyle.Bold)
        self._lb_title.ForeColor = System.Drawing.Color.Red
        self._lb_title.Location  = System.Drawing.Point(12, 10)
        self._lb_title.Size      = System.Drawing.Size(460, 24)

        # ── lb_hint ───────────────────────────────────────────
        self._lb_hint.Text      = "Leave 'New Number' or 'New Name' blank to keep existing value."
        self._lb_hint.Font      = System.Drawing.Font("Meiryo UI", 7.5, System.Drawing.FontStyle.Italic)
        self._lb_hint.ForeColor = System.Drawing.Color.Gray
        self._lb_hint.Location  = System.Drawing.Point(14, 36)
        self._lb_hint.Size      = System.Drawing.Size(560, 16)

        # ── DataGridView ──────────────────────────────────────
        self._dgv.Location              = System.Drawing.Point(12, 58)
        self._dgv.Size                  = System.Drawing.Size(760, 460)
        self._dgv.AllowUserToAddRows    = False
        self._dgv.AllowUserToDeleteRows = False
        self._dgv.RowHeadersVisible     = False
        self._dgv.SelectionMode         = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        self._dgv.MultiSelect           = True
        self._dgv.AutoSizeRowsMode      = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        self._dgv.BorderStyle           = System.Windows.Forms.BorderStyle.Fixed3D
        self._dgv.Font                  = System.Drawing.Font("Meiryo UI", 8)
        self._dgv.GridColor             = System.Drawing.Color.LightGray
        self._dgv.BackgroundColor       = System.Drawing.Color.WhiteSmoke
        self._dgv.CellBeginEdit        += self.Dgv_CellBeginEdit
        self._dgv.KeyDown              += self.Dgv_KeyDown

        # Cot 1: STT
        col_stt            = System.Windows.Forms.DataGridViewTextBoxColumn()
        col_stt.HeaderText = "#"
        col_stt.Name       = "col_stt"
        col_stt.Width      = 36
        col_stt.ReadOnly   = True
        col_stt.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        col_stt.DefaultCellStyle.ForeColor = System.Drawing.Color.Gray

        # Cot 2: Current Number (readonly)
        col_cur_num                              = System.Windows.Forms.DataGridViewTextBoxColumn()
        col_cur_num.HeaderText                   = "Current Number"
        col_cur_num.Name                         = "col_cur_num"
        col_cur_num.Width                        = 130
        col_cur_num.ReadOnly                     = True
        col_cur_num.DefaultCellStyle.BackColor   = System.Drawing.Color.FromArgb(235, 240, 255)
        col_cur_num.DefaultCellStyle.ForeColor   = System.Drawing.Color.DarkBlue
        col_cur_num.DefaultCellStyle.Font        = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)

        # Cot 3: Current Name (readonly)
        col_cur_name                             = System.Windows.Forms.DataGridViewTextBoxColumn()
        col_cur_name.HeaderText                  = "Current Name"
        col_cur_name.Name                        = "col_cur_name"
        col_cur_name.Width                       = 210
        col_cur_name.ReadOnly                    = True
        col_cur_name.DefaultCellStyle.BackColor  = System.Drawing.Color.FromArgb(235, 240, 255)
        col_cur_name.DefaultCellStyle.ForeColor  = System.Drawing.Color.DarkBlue
        col_cur_name.DefaultCellStyle.Font       = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)

        # Cot 4: Separator
        col_sep            = System.Windows.Forms.DataGridViewTextBoxColumn()
        col_sep.HeaderText = ">>"
        col_sep.Name       = "col_sep"
        col_sep.Width      = 28
        col_sep.ReadOnly   = True
        col_sep.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        col_sep.DefaultCellStyle.ForeColor = System.Drawing.Color.Gray
        col_sep.DefaultCellStyle.BackColor = System.Drawing.Color.WhiteSmoke

        # Cot 5: New Number (editable)
        col_new_num                              = System.Windows.Forms.DataGridViewTextBoxColumn()
        col_new_num.HeaderText                   = "New Number"
        col_new_num.Name                         = "col_new_num"
        col_new_num.Width                        = 130
        col_new_num.ReadOnly                     = False
        col_new_num.DefaultCellStyle.BackColor   = System.Drawing.Color.FromArgb(255, 255, 220)
        col_new_num.DefaultCellStyle.ForeColor   = System.Drawing.Color.DarkRed
        col_new_num.DefaultCellStyle.Font        = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)

        # Cot 6: New Name (editable)
        col_new_name                             = System.Windows.Forms.DataGridViewTextBoxColumn()
        col_new_name.HeaderText                  = "New Name"
        col_new_name.Name                        = "col_new_name"
        col_new_name.AutoSizeMode                = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        col_new_name.ReadOnly                    = False
        col_new_name.DefaultCellStyle.BackColor  = System.Drawing.Color.FromArgb(255, 255, 220)
        col_new_name.DefaultCellStyle.ForeColor  = System.Drawing.Color.DarkRed
        col_new_name.DefaultCellStyle.Font       = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)

        # Header style
        self._dgv.ColumnHeadersDefaultCellStyle.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._dgv.ColumnHeadersDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(60, 60, 80)
        self._dgv.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White
        self._dgv.ColumnHeadersDefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        self._dgv.ColumnHeadersHeightSizeMode             = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        self._dgv.ColumnHeadersHeight                     = 28
        self._dgv.EnableHeadersVisualStyles               = False

        self._dgv.Columns.Add(col_stt)
        self._dgv.Columns.Add(col_cur_num)
        self._dgv.Columns.Add(col_cur_name)
        self._dgv.Columns.Add(col_sep)
        self._dgv.Columns.Add(col_new_num)
        self._dgv.Columns.Add(col_new_name)

        # ── progressBar ───────────────────────────────────────
        self._progressBar.Location = System.Drawing.Point(12, 526)
        self._progressBar.Size     = System.Drawing.Size(760, 14)
        self._progressBar.Minimum  = 0
        self._progressBar.Maximum  = 100
        self._progressBar.Value    = 0
        self._progressBar.Style    = System.Windows.Forms.ProgressBarStyle.Continuous

        # ── lb_status ─────────────────────────────────────────
        self._lb_status.Text      = "Ready | {} sheets loaded.".format(len(self._sheets))
        self._lb_status.Font      = System.Drawing.Font("Meiryo UI", 7.5, System.Drawing.FontStyle.Bold)
        self._lb_status.ForeColor = System.Drawing.Color.Gray
        self._lb_status.Location  = System.Drawing.Point(12, 546)
        self._lb_status.Size      = System.Drawing.Size(560, 18)

        # ── btt_RUN ───────────────────────────────────────────
        self._btt_RUN.Text      = "RUN"
        self._btt_RUN.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._btt_RUN.ForeColor = System.Drawing.Color.Red
        self._btt_RUN.Location  = System.Drawing.Point(492, 570)
        self._btt_RUN.Size      = System.Drawing.Size(85, 36)
        self._btt_RUN.UseVisualStyleBackColor = True
        self._btt_RUN.Click    += self.Btt_RUNClick

        # ── btt_CLEAR ─────────────────────────────────────────
        self._btt_CLEAR.Text      = "CLEAR"
        self._btt_CLEAR.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._btt_CLEAR.ForeColor = System.Drawing.Color.DarkOrange
        self._btt_CLEAR.Location  = System.Drawing.Point(590, 570)
        self._btt_CLEAR.Size      = System.Drawing.Size(85, 36)
        self._btt_CLEAR.UseVisualStyleBackColor = True
        self._btt_CLEAR.Click    += self.Btt_CLEARClick

        # ── btt_CANCLE ────────────────────────────────────────
        self._btt_CANCLE.Text      = "CLOSE"
        self._btt_CANCLE.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
        self._btt_CANCLE.Location  = System.Drawing.Point(687, 570)
        self._btt_CANCLE.Size      = System.Drawing.Size(85, 36)
        self._btt_CANCLE.UseVisualStyleBackColor = True
        self._btt_CANCLE.Click    += self.Btt_CANCLEClick

        # ── lb_FVC ────────────────────────────────────────────
        self._lb_FVC.Text      = "@FVC"
        self._lb_FVC.Font      = System.Drawing.Font("Meiryo UI", 4.8, System.Drawing.FontStyle.Bold)
        self._lb_FVC.ForeColor = System.Drawing.Color.LightGray
        self._lb_FVC.Location  = System.Drawing.Point(14, 580)
        self._lb_FVC.Size      = System.Drawing.Size(48, 16)

        # ── MainForm ──────────────────────────────────────────
        self.Text            = "Sheet Batch Editor"
        self.ClientSize      = System.Drawing.Size(784, 616)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.TopMost         = True
        self.Controls.Add(self._lb_title)
        self.Controls.Add(self._lb_hint)
        self.Controls.Add(self._dgv)
        self.Controls.Add(self._progressBar)
        self.Controls.Add(self._lb_status)
        self.Controls.Add(self._btt_RUN)
        self.Controls.Add(self._btt_CLEAR)
        self.Controls.Add(self._btt_CANCLE)
        self.Controls.Add(self._lb_FVC)
        self.ResumeLayout(False)

    # ── Load du lieu vao grid ─────────────────────────────────
    def _loadSheets(self):
        self._dgv.Rows.Clear()
        for i, sheet in enumerate(self._sheets):
            row_idx = self._dgv.Rows.Add()
            row     = self._dgv.Rows[row_idx]
            row.Cells["col_stt"].Value      = i + 1
            row.Cells["col_cur_num"].Value  = sheet.SheetNumber
            row.Cells["col_cur_name"].Value = sheet.Name
            row.Cells["col_sep"].Value      = ">>"
            row.Cells["col_new_num"].Value  = ""
            row.Cells["col_new_name"].Value = ""
            if i % 2 == 0:
                row.DefaultCellStyle.BackColor = System.Drawing.Color.White
            else:
                row.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(248, 248, 252)

    # ── Highlight o dang edit ─────────────────────────────────
    def Dgv_CellBeginEdit(self, sender, e):
        col = e.ColumnIndex
        if col == self._dgv.Columns["col_new_num"].Index or \
           col == self._dgv.Columns["col_new_name"].Index:
            self._dgv.Rows[e.RowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 250, 200)

    # ── Enter de xuong dong ───────────────────────────────────
    def Dgv_KeyDown(self, sender, e):
        if e.KeyCode == System.Windows.Forms.Keys.Enter:
            cur_row = self._dgv.CurrentCell.RowIndex
            cur_col = self._dgv.CurrentCell.ColumnIndex
            self._dgv.EndEdit()
            next_row = cur_row + 1
            if next_row < self._dgv.Rows.Count:
                self._dgv.CurrentCell = self._dgv.Rows[next_row].Cells[cur_col]
            e.Handled = True

    # ── RUN ───────────────────────────────────────────────────
    def Btt_RUNClick(self, sender, e):
        if self._isRunning:
            return

        updates = []
        for i, sheet in enumerate(self._sheets):
            row      = self._dgv.Rows[i]
            new_num  = str(row.Cells["col_new_num"].Value  or "").strip()
            new_name = str(row.Cells["col_new_name"].Value or "").strip()
            if new_num or new_name:
                updates.append((sheet, new_num, new_name))

        if not updates:
            MessageBox.Show(
                "No changes detected.\nPlease enter New Number or New Name for at least one sheet.",
                "Nothing to Update",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            return

        # Kiem tra trung Number
        all_existing_numbers = set(s.SheetNumber for s in self._sheets)
        new_numbers_entered  = {}
        duplicates           = []
        for sheet, new_num, new_name in updates:
            if not new_num:
                continue
            if new_num in all_existing_numbers and new_num != sheet.SheetNumber:
                duplicates.append("'{}' -> '{}' (already exists)".format(sheet.SheetNumber, new_num))
            if new_num in new_numbers_entered:
                duplicates.append("'{}' duplicate in new entries".format(new_num))
            new_numbers_entered[new_num] = sheet

        if duplicates:
            msg = "Duplicate sheet numbers detected:\n\n" + "\n".join(duplicates) + \
                  "\n\nPlease fix before running."
            MessageBox.Show(msg, "Duplicate Numbers",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return

        confirm_msg = "About to update {} sheet(s).\n\nContinue?".format(len(updates))
        if MessageBox.Show(confirm_msg, "Confirm Update",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) \
                != System.Windows.Forms.DialogResult.Yes:
            return

        # Transaction
        self._isRunning          = True
        self._btt_RUN.Enabled    = False
        self._btt_CLEAR.Enabled  = False
        self._btt_CANCLE.Enabled = False
        self._progressBar.Maximum = len(updates)
        self._progressBar.Value   = 0

        success = []
        failed  = []

        TransactionManager.Instance.EnsureInTransaction(doc)
        try:
            for i, (sheet, new_num, new_name) in enumerate(updates):
                old_num  = sheet.SheetNumber
                old_name = sheet.Name
                try:
                    if new_num:
                        sheet.SheetNumber = new_num
                    if new_name:
                        sheet.Name = new_name
                    success.append("[{}] '{}' -> '{}' | '{}' -> '{}'".format(
                        i + 1,
                        old_num,  new_num  if new_num  else old_num,
                        old_name, new_name if new_name else old_name))
                except Exception as ex:
                    try:
                        sheet.SheetNumber = old_num
                        sheet.Name        = old_name
                    except:
                        pass
                    failed.append("  x [{}] '{}' | {}".format(old_num, old_name, str(ex)))

                self._progressBar.Value   = i + 1
                self._lb_status.ForeColor = System.Drawing.Color.DarkOrange
                self._lb_status.Text      = "Updating [{}/{}]: {}".format(
                    i + 1, len(updates), sheet.SheetNumber)
                System.Windows.Forms.Application.DoEvents()

            TransactionManager.Instance.TransactionTaskDone()
        except Exception as ex:
            TransactionManager.Instance.ForceCloseTransaction()
            failed.append("Transaction error: " + str(ex))

        # Reload grid
        self._sheets = getAllSheets(doc)
        self._loadSheets()

        if failed:
            self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
            self._lb_status.Text = "Done | {} updated, {} failed.".format(
                len(success), len(failed))
            detail = "Updated:\n" + "\n".join(success) + \
                     "\n\nFailed:\n" + "\n".join(failed)
            MessageBox.Show(detail, "Result",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
        else:
            self._lb_status.ForeColor = System.Drawing.Color.Green
            self._lb_status.Text = "Done | {} sheet(s) updated successfully!".format(len(success))

        self._isRunning          = False
        self._btt_RUN.Enabled    = True
        self._btt_CLEAR.Enabled  = True
        self._btt_CANCLE.Enabled = True

    # ── CLEAR ─────────────────────────────────────────────────
    def Btt_CLEARClick(self, sender, e):
        for row in self._dgv.Rows:
            row.Cells["col_new_num"].Value  = ""
            row.Cells["col_new_name"].Value = ""
        self._lb_status.ForeColor = System.Drawing.Color.Gray
        self._lb_status.Text      = "Cleared | {} sheets loaded.".format(len(self._sheets))

    # ── CLOSE ─────────────────────────────────────────────────
    def Btt_CANCLEClick(self, sender, e):
        self.Close()
#endregion

f = SheetEditorForm()
Application.Run(f)