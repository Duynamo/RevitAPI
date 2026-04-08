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

# ── Japanese parameter name constants ──────────────────────────────────────
CONST_ROOM_NAME   = u"01_\u5ba4\u540d"                    # 01_室名
CONST_FLOOR_LEVEL = u"02_\u30d5\u30ed\u30a2\u30ec\u30d9\u30eb"    # 02_フロアレベル
CONST_CEIL_LEVEL  = u"03_\u5929\u4e95\u30ec\u30d9\u30eb"          # 03_天井レベル
CONST_RECEIPT_DATE= u"04_\u53c2\u8003\u56f3\u53d7\u9818\u65e5"    # 04_参考図受領日
CONST_REF_PAGE    = u"05_\u30d5\u30ed\u30a2\u30ec\u30d9\u30eb\u53c2\u8003\u30da\u30fc\u30b8"  # 05_フロアレベル参考ページ

# ── Revit document handles ─────────────────────────────────────────────────
doc    = DocumentManager.Instance.CurrentDBDocument
uiapp  = DocumentManager.Instance.CurrentUIApplication
app    = uiapp.Application
uidoc  = uiapp.ActiveUIDocument

# ══════════════════════════════════════════════════════════════════════════
#  EXCEL HELPERS
# ══════════════════════════════════════════════════════════════════════════

def get_hidden_excel(file_path):
    excel_type = System.Type.GetTypeFromProgID("Excel.Application")
    if excel_type is None:
        raise Exception("Excel is not installed or COM registry is missing.")
    excel = System.Activator.CreateInstance(excel_type)
    excel.Visible        = False
    excel.ScreenUpdating = False
    excel.DisplayAlerts  = False
    try:
        workbook = excel.Workbooks.Open(file_path)
        return excel, workbook
    except Exception as e:
        try: excel.Quit()
        except: pass
        raise e


def read_excel_data(file_path, sheet_name):
    """Return (headers_list, rows_list_of_dicts) or (None, error_msg)."""
    excel = workbook = None
    try:
        excel, workbook = get_hidden_excel(file_path)

        sheet = None
        for i in range(1, workbook.Worksheets.Count + 1):
            if unicode(workbook.Worksheets(i).Name) == unicode(sheet_name):
                sheet = workbook.Worksheets(i)
                break

        if sheet is None:
            return None, u"Sheet '%s' does not exist." % sheet_name

        used  = sheet.UsedRange
        n_row = used.Rows.Count
        n_col = used.Columns.Count

        if n_row < 1:
            return None, u"Sheet '%s' is empty." % sheet_name

        # Row 1 = headers
        headers = []
        for c in range(1, n_col + 1):
            v = sheet.Cells(1, c).Value2
            headers.append(unicode(v) if v is not None else u"Col_%d" % c)

        if CONST_ROOM_NAME not in headers:
            return None, u"Missing column '%s' in sheet '%s'." % (CONST_ROOM_NAME, sheet_name)

        # Rows 2..n = data
        data = []
        for r in range(2, n_row + 1):
            row = {}
            for idx, h in enumerate(headers):
                v = sheet.Cells(r, idx + 1).Value2
                row[h] = v  # keep raw (float / str / None)
            key = row.get(CONST_ROOM_NAME)
            if key is not None and unicode(key).strip():
                data.append(row)

        return headers, data

    except Exception as e:
        return None, u"Error reading Excel: %s" % unicode(e)
    finally:
        if workbook:
            try: workbook.Close(False)
            except: pass
        if excel:
            try: excel.Quit()
            except: pass


# ══════════════════════════════════════════════════════════════════════════
#  PARAMETER WRITE HELPERS
# ══════════════════════════════════════════════════════════════════════════

def _write_param(param, value):
    """
    Low-level: clear then write to a single Revit Parameter object.
    Returns (True, info) or (False, reason).
    """
    if param is None:
        return False, "param is None"
    if param.IsReadOnly:
        return False, "read-only"

    st = param.StorageType

    if st == StorageType.String:
        # clear
        param.Set(u"")
        if value is None:
            return True, "cleared"
        # convert float-integers (e.g. 3500.0) to "3500"
        if isinstance(value, float) and value == math.floor(value):
            s = unicode(int(value))
        else:
            s = unicode(value)
        param.Set(s)
        return True, u"set string -> '%s'" % s

    elif st == StorageType.Double:
        param.Set(0.0)
        if value is None:
            return True, "cleared"
        param.Set(float(value))
        return True, u"set double -> %s" % value

    elif st == StorageType.Integer:
        param.Set(0)
        if value is None:
            return True, "cleared"
        param.Set(int(float(value)))
        return True, u"set int -> %s" % value

    else:
        return False, u"StorageType.ElementId or unknown — cannot set from Excel value"


def set_shared_param(elem, param_name, value):
    """
    Find shared/instance parameter by name on element (not type),
    clear then write value.  Returns (True/False, message).
    """
    param = elem.LookupParameter(param_name)
    if param is None:
        return False, u"param '%s' not found on element" % param_name
    return _write_param(param, value)


def set_level_builtin(elem, level_obj):
    """Try to set the Level BuiltInParameter to level_obj.Id.
    Works for Ceiling, Floor, Wall, Family Instance, etc. Returns bool.
    """
    bips = [
        BuiltInParameter.LEVEL_PARAM,                      # Ceiling / Floor / generic
        BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,   # Family instances
        BuiltInParameter.FAMILY_LEVEL_PARAM,
        BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
        BuiltInParameter.WALL_BASE_CONSTRAINT,              # Walls
    ]
    for bip in bips:
        p = elem.get_Parameter(bip)
        if p is not None and not p.IsReadOnly:
            try:
                p.Set(level_obj.Id)
                return True
            except:
                pass
    return False


def set_offset_builtin(elem, offset_mm):
    """Set height-offset BuiltInParameter (mm → feet).
    Priority: Ceiling > Floor > Wall > Family > Roof. Returns bool.
    """
    offset_ft = float(offset_mm) / 304.8
    bips = [
        BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM,   # Ceiling  ← 1st priority
        BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM,     # Floor
        BuiltInParameter.WALL_BASE_OFFSET,                 # Wall
        BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,  # Family hosted
        BuiltInParameter.INSTANCE_ELEVATION_PARAM,         # Family unhosted
        BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM,
        BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM,          # Roof
    ]
    for bip in bips:
        p = elem.get_Parameter(bip)
        if p is not None and not p.IsReadOnly:
            try:
                p.Set(offset_ft)
                return True
            except:
                pass
    return False


# ══════════════════════════════════════════════════════════════════════════
#  MAIN PROCESS  (uses a native Revit Transaction — NOT TransactionManager)
# ══════════════════════════════════════════════════════════════════════════

def process_excel_data(headers, data):
    """
    Map Excel rows  →  Revit elements via 01_室名.
    For every matched element:
      1. Write ALL shared-parameter columns (clear old, write new).
      2. If 02_フロアレベル  → also set Level   BuiltInParameter.
      3. If 03_天井レベル   → also set Offset  BuiltInParameter (mm→ft).
    Returns (success_count, error_count, log_lines_list).
    """
    success_count = 0
    error_count   = 0
    log           = []

    # ── Build Excel lookup map ────────────────────────────────────────────
    excel_map = {}
    for row in data:
        key = row.get(CONST_ROOM_NAME)
        if key is not None:
            k = unicode(key).strip()
            if k:
                excel_map[k] = row

    if not excel_map:
        return 0, 0, [u"Excel map is empty — no valid rows found."]

    # ── Build Level name dict ─────────────────────────────────────────────
    level_dict = {}
    for lvl in FilteredElementCollector(doc).OfClass(Level).ToElements():
        level_dict[lvl.Name] = lvl

    # ── Collect candidate elements ────────────────────────────────────────
    sel_ids = list(uidoc.Selection.GetElementIds())
    if sel_ids:
        candidates = [doc.GetElement(eid) for eid in sel_ids]
        candidates = [e for e in candidates if e is not None]
        log.append(u"Source: %d selected elements." % len(candidates))
    else:
        candidates = list(
            FilteredElementCollector(doc, doc.ActiveView.Id)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        log.append(u"Source: %d elements from active view." % len(candidates))

    # Keep only elements that carry 01_室名
    targets = [e for e in candidates if e.LookupParameter(CONST_ROOM_NAME) is not None]
    log.append(u"Elements with '%s': %d" % (CONST_ROOM_NAME, len(targets)))

    if not targets:
        return 0, 0, log + [u"No elements have the '%s' parameter." % CONST_ROOM_NAME]

    # ── Open a single native Revit transaction ────────────────────────────
    t = Transaction(doc, "Import Shared Parameters from Excel")
    t.Start()

    try:
        matched = 0
        for elem in targets:
            p_room = elem.LookupParameter(CONST_ROOM_NAME)
            if p_room is None:
                continue
            room_val = p_room.AsString()
            if not room_val:
                continue

            row_data = excel_map.get(unicode(room_val).strip())
            if row_data is None:
                continue

            matched += 1
            elem_ok  = False
            elem_log = []

            # ── Step 1: Write every Excel column to the matching shared param ──
            for col in headers:
                if col == CONST_ROOM_NAME:
                    continue          # skip key column
                val = row_data.get(col)

                ok, info = set_shared_param(elem, col, val)
                elem_log.append(u"  param '%s': %s" % (col, info))
                if ok:
                    elem_ok = True

            # ── Step 2: Set Level BuiltInParameter (02_フロアレベル) ───────────
            level_val = row_data.get(CONST_FLOOR_LEVEL)
            if level_val is not None:
                lvl_name = unicode(level_val).strip()
                if lvl_name in level_dict:
                    ok = set_level_builtin(elem, level_dict[lvl_name])
                    elem_log.append(u"  Level BIP '%s': %s" % (lvl_name, "OK" if ok else "FAIL"))
                    if ok: elem_ok = True
                else:
                    elem_log.append(u"  Level '%s' not found in doc." % lvl_name)

            # ── Step 3: Set Offset BuiltInParameter (03_天井レベル, mm→ft) ─────
            offset_val = row_data.get(CONST_CEIL_LEVEL)
            if offset_val is not None and unicode(offset_val).strip() != u"":
                try:
                    ok = set_offset_builtin(elem, float(offset_val))
                    elem_log.append(u"  Offset BIP (%s mm): %s" % (offset_val, "OK" if ok else "FAIL"))
                    if ok: elem_ok = True
                except Exception as ex:
                    elem_log.append(u"  Offset BIP error: %s" % unicode(ex))

            # ── Tally ──────────────────────────────────────────────────────────
            if elem_ok:
                success_count += 1
                log.append(u"[OK] %s (id=%s)" % (room_val, elem.Id))
            else:
                error_count += 1
                log.append(u"[FAIL] %s (id=%s)" % (room_val, elem.Id))
                log.extend(elem_log)

        if matched == 0:
            t.RollBack()
            return 0, 0, log + [u"No elements matched any room name from Excel."]

        t.Commit()

    except Exception as ex:
        try: t.RollBack()
        except: pass
        return 0, 0, log + [u"Transaction error: %s" % unicode(ex)]

    return success_count, error_count, log


# ══════════════════════════════════════════════════════════════════════════
#  WINFORMS UI
# ══════════════════════════════════════════════════════════════════════════

class MainForm(Form):
    def __init__(self):
        self.InitializeComponent()

    def InitializeComponent(self):
        c_primary    = Color.FromArgb(98,  0, 238)
        c_success    = Color.FromArgb(76, 175, 80)
        c_bg_main    = Color.FromArgb(245, 245, 246)
        c_surface    = Color.White
        c_text_dark  = Color.FromArgb(33,  33,  33)
        c_btn_cancel = Color.FromArgb(224, 224, 224)

        self.Text            = "Import Shared Parameters (KSC01)"
        self.ClientSize      = Size(500, 310)
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False
        self.MinimizeBox     = False
        self.BackColor       = c_bg_main
        self.Font            = Font("Segoe UI", 10)

        # Panel
        self.panel = Panel()
        self.panel.BackColor = c_surface
        self.panel.Location  = Point(20, 20)
        self.panel.Size      = Size(460, 200)

        self.lbl_title = Label()
        self.lbl_title.Text      = "Data Source"
        self.lbl_title.Font      = Font("Segoe UI", 12, FontStyle.Bold)
        self.lbl_title.ForeColor = c_primary
        self.lbl_title.Location  = Point(20, 15)
        self.lbl_title.AutoSize  = True

        x_left = 20; width_left = 160; x_right = 190; width_right = 255
        y_row1 = 60; y_row2 = 115;  h_btn = 35; h_input = 28

        self.btt_linkExcel = Button()
        self.btt_linkExcel.Text      = "SELECT EXCEL"
        self.btt_linkExcel.Font      = Font("Segoe UI", 10, FontStyle.Bold)
        self.btt_linkExcel.BackColor = c_primary
        self.btt_linkExcel.ForeColor = Color.White
        self.btt_linkExcel.FlatStyle = FlatStyle.Flat
        self.btt_linkExcel.FlatAppearance.BorderSize = 0
        self.btt_linkExcel.Location  = Point(x_left, y_row1)
        self.btt_linkExcel.Size      = Size(width_left, h_btn)
        self.btt_linkExcel.Click    += self.Btt_linkExcelClick

        self.txb_linkExcel = TextBox()
        self.txb_linkExcel.Location    = Point(x_right, y_row1 + 4)
        self.txb_linkExcel.Size        = Size(width_right, h_input)
        self.txb_linkExcel.ReadOnly    = True
        self.txb_linkExcel.BackColor   = Color.FromArgb(250, 250, 250)
        self.txb_linkExcel.BorderStyle = BorderStyle.FixedSingle

        self.lbl_ws = Label()
        self.lbl_ws.Text        = "WORKSHEET"
        self.lbl_ws.Font        = Font("Segoe UI", 10, FontStyle.Bold)
        self.lbl_ws.BackColor   = c_primary
        self.lbl_ws.ForeColor   = Color.White
        self.lbl_ws.Location    = Point(x_left, y_row2)
        self.lbl_ws.Size        = Size(width_left, h_btn)
        self.lbl_ws.TextAlign   = ContentAlignment.MiddleCenter
        self.lbl_ws.AutoSize    = False

        self.comboBox1 = ComboBox()
        self.comboBox1.DropDownStyle = ComboBoxStyle.DropDownList
        self.comboBox1.Location      = Point(x_right, y_row2 + 4)
        self.comboBox1.Size          = Size(width_right, h_input)
        self.comboBox1.BackColor     = Color.White

        self.lbl_D = Label()
        self.lbl_D.Font      = Font("Segoe UI", 7, FontStyle.Bold)
        self.lbl_D.ForeColor = Color.DarkGray
        self.lbl_D.Location  = Point(20, 165)
        self.lbl_D.Text      = "@D"
        self.lbl_D.AutoSize  = True

        self.panel.Controls.Add(self.lbl_title)
        self.panel.Controls.Add(self.btt_linkExcel)
        self.panel.Controls.Add(self.txb_linkExcel)
        self.panel.Controls.Add(self.lbl_ws)
        self.panel.Controls.Add(self.comboBox1)
        self.panel.Controls.Add(self.lbl_D)

        self.btt_Run = Button()
        self.btt_Run.Text      = "RUN IMPORT"
        self.btt_Run.Font      = Font("Segoe UI", 10, FontStyle.Bold)
        self.btt_Run.BackColor = c_success
        self.btt_Run.ForeColor = Color.White
        self.btt_Run.FlatStyle = FlatStyle.Flat
        self.btt_Run.FlatAppearance.BorderSize = 0
        self.btt_Run.Location  = Point(215, 240)
        self.btt_Run.Size      = Size(140, 42)
        self.btt_Run.Click    += self.Btt_RunClick

        self.btt_Cancle = Button()
        self.btt_Cancle.Text      = "CANCEL"
        self.btt_Cancle.Font      = Font("Segoe UI", 10, FontStyle.Bold)
        self.btt_Cancle.BackColor = c_btn_cancel
        self.btt_Cancle.ForeColor = c_text_dark
        self.btt_Cancle.FlatStyle = FlatStyle.Flat
        self.btt_Cancle.FlatAppearance.BorderSize = 0
        self.btt_Cancle.Location  = Point(370, 240)
        self.btt_Cancle.Size      = Size(110, 42)
        self.btt_Cancle.Click    += self.Btt_CancleClick

        self.Controls.Add(self.panel)
        self.Controls.Add(self.btt_Run)
        self.Controls.Add(self.btt_Cancle)
        self.TopMost = True

    # ── Open file dialog & populate worksheet combo ────────────────────────
    def Btt_linkExcelClick(self, sender, e):
        dlg = OpenFileDialog()
        dlg.Multiselect = False
        dlg.Filter      = "Excel Files (*.xlsx)|*.xlsx"
        dlg.Title       = "Select Excel file"
        if dlg.ShowDialog() != DialogResult.OK:
            return

        path = dlg.FileName
        self.txb_linkExcel.Text = path
        excel = workbook = None
        try:
            excel, workbook = get_hidden_excel(path)
            self.comboBox1.Items.Clear()
            for i in range(1, workbook.Worksheets.Count + 1):
                self.comboBox1.Items.Add(unicode(workbook.Worksheets(i).Name))
            if self.comboBox1.Items.Count > 0:
                self.comboBox1.SelectedIndex = 0
        except Exception as ex:
            MessageBox.Show("Error opening Excel: " + str(ex), "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        finally:
            if workbook:
                try: workbook.Close(False)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass

    # ── RUN button ─────────────────────────────────────────────────────────
    def Btt_RunClick(self, sender, e):
        sheet_name = self.comboBox1.SelectedItem
        file_path  = self.txb_linkExcel.Text

        if not file_path or not sheet_name:
            MessageBox.Show("Please select a valid file and Worksheet.",
                            "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return

        # 1. Read Excel (outside transaction — pure read)
        headers, data = read_excel_data(file_path, unicode(sheet_name))
        if headers is None:
            # data holds the error message when headers is None
            MessageBox.Show(unicode(data), "Excel Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            self.Close()
            return

        # 2. Process and write to Revit (uses native Transaction inside)
        try:
            success, errors, log = process_excel_data(headers, data)

            if success == 0 and errors == 0:
                detail = u"\n".join(log)
                MessageBox.Show(detail, "Import Finished — No Updates",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            elif errors > 0:
                detail = u"Updated: %d   Failed: %d\n\n" % (success, errors)
                detail += u"\n".join(log[-30:])   # show last 30 log lines
                MessageBox.Show(detail, "Import Finished (with issues)",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            else:
                detail = u"Successfully updated %d elements." % success
                MessageBox.Show(detail, "Import Finished",
                                MessageBoxButtons.OK, MessageBoxIcon.Information)

        except Exception as ex:
            MessageBox.Show("Unexpected error: " + str(ex), "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

        self.Close()

    def Btt_CancleClick(self, sender, e):
        self.Close()


# ── Entry point ────────────────────────────────────────────────────────────
f = MainForm()
Application.Run(f)
