"""
Create Piping System Tool - V2D Tools
Chuc nang: Doc thong tin Piping System tu file Excel va tao cac system trong Revit.
Cac thong so doc tu sheet "Piping System":
  - STT         : So thu tu
  - TEN SYSTEM  : Ten cua piping system
  - ABBREVIATION: Viet tat
  - DO DAY      : Do day net (line weight) - luu vao LineWeight
  - KIEU NET    : Kieu duong ket (Solid, Dash, Dot, Hidden, ...)
  - MA MAU RGB  : Mau sac dang "R-G-B" (vi du: 128-100-162)
  - HIEN THI    : Co hien thi hay khong

Luu y:
  - Dung WinForms (khong dung WPF) vi WinForms hop tac tot voi Revit/Dynamo.
  - Tuong thich moi phien ban Revit (2019 tro len).
  - Comment giai thich bang tieng Viet, UI bang tieng Anh.
"""

# region --- Imports chung ---
import clr
import System
import os

# Them tham chieu WinForms va Drawing de xay dung UI
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import System.Windows.Forms as WinForms
import System.Drawing as Drawing
from System.Windows.Forms import (
    Form, Label, Button, TextBox, Panel, Application,
    FormBorderStyle, FormStartPosition, FlatStyle, BorderStyle,
    AutoScaleMode, OpenFileDialog, DialogResult, MessageBox,
    MessageBoxButtons, MessageBoxIcon, ComboBox, ComboBoxStyle,
    DataGridView, DataGridViewAutoSizeColumnsMode, ScrollBars,
    ProgressBar, CheckBox, ToolTip, FlowLayoutPanel, AnchorStyles
)
from System.Drawing import (
    Point, Size, Font, FontStyle, GraphicsUnit,
    Color, ContentAlignment, SolidBrush, Pen, Rectangle
)

# Them tham chieu Revit API
clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import (
    Transaction, Color as RevitColor, ElementId,
    FilteredElementCollector, BuiltInCategory, LinePatternElement,
    BuiltInParameter, GraphicsStyle, GraphicsStyleType
)


# Import Piping namespace - tuong thich moi phien ban Revit
try:
    # Revit 2019+: PipingSystemType nam trong DB.Plumbing
    from Autodesk.Revit.DB.Plumbing import PipingSystemType
    HAS_PLUMBING = True
except ImportError:
    HAS_PLUMBING = False

# endregion

# region --- Revit context ---
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
uidoc = uiapp.ActiveUIDocument
# endregion

# region --- Global output ---
# OUT la ket qua tra ve cho Dynamo
log_messages = []
# endregion

# region --- Xu ly Excel ---

def read_excel_data(filepath, sheet_name=None):
    """
    Doc file Excel bang phuong phap Late Binding (dieu khien Excel an).
    Yeu cau Microsoft Excel phai duoc cai dat tren may.
    Khong yeu cau driver phu (nhu OLEDB) hay thu vien ngoai (nhu openpyxl).

    Neu sheet_name la None, chi tra ve danh sach sheet.
    Neu sheet_name duoc cung cap, tra ve du lieu cua sheet do.

    Tra ve: (data, sheet_names, error_message)
    """
    excel = None
    workbook = None

    if not os.path.exists(filepath):
        return None, [], "File not found: {}".format(filepath)

    try:
        # Khoi tao instance Excel an
        excel_type = System.Type.GetTypeFromProgID("Excel.Application")
        excel = System.Activator.CreateInstance(excel_type)
        excel.Visible = False
        excel.DisplayAlerts = False

        # Mo workbook o che do chi doc
        workbook = excel.Workbooks.Open(filepath, False, True) # UpdateLinks, ReadOnly

        # Lay danh sach tat ca sheet
        sheet_names = [sh.Name for sh in workbook.Worksheets]

        if sheet_name is None:
            return None, sheet_names, None

        if sheet_name not in sheet_names:
            return None, sheet_names, "Sheet '{}' not found.".format(sheet_name)

        ws = workbook.Worksheets(sheet_name)
        used_range = ws.UsedRange

        if used_range is None or used_range.Rows.Count < 2:
            return [], sheet_names, None

        # Doc toan bo du lieu vao mang de toi uu toc do
        all_values = used_range.Value2
        
        if not isinstance(all_values, System.Array): # Neu chi co 1 cell
            return [], sheet_names, None
        
        num_rows = all_values.GetLength(0)
        num_cols = all_values.GetLength(1)

        # Dong dau tien la header
        headers = [str(all_values[0, j]).strip().upper() if all_values[0, j] is not None else "" for j in range(num_cols)]

        # Doc cac dong du lieu
        data = []
        for i in range(1, num_rows): # Bat dau tu dong thu 2 (index 1)
            row_dict = {}
            is_empty_row = True
            for j in range(num_cols):
                header = headers[j]
                if not header: continue
                
                val = all_values[i, j]
                row_dict[header] = val
                if val is not None and str(val).strip() != "":
                    is_empty_row = False
            
            if not is_empty_row:
                data.append(row_dict)

        return data, sheet_names, None

    except Exception as ex:
        msg = str(ex)
        if "Excel is not installed" in msg or "ProgID" in msg or "80040154" in msg:
            msg = "Microsoft Excel is not installed or not registered correctly."
        return None, [], "Excel Read Error: " + msg
    finally:
        # Luon luon dam bao dong va giai phong tien trinh Excel
        if workbook is not None:
            try: workbook.Close(False)
            except: pass
        if excel is not None:
            try: excel.Quit()
            except: pass


def parse_rgb(rgb_str):
    """
    Chuyen chuoi 'R-G-B' (vi du: '128-100-162') thanh tuple (r, g, b).
    Tra ve (128, 100, 162) neu hop le, nguoc lai tra ve (128, 128, 128).
    """
    if not rgb_str:
        return (128, 128, 128)
    try:
        parts = str(rgb_str).strip().split('-')
        if len(parts) == 3:
            r = max(0, min(255, int(parts[0])))
            g = max(0, min(255, int(parts[1])))
            b = max(0, min(255, int(parts[2])))
            return (r, g, b)
    except Exception:
        pass
    return (128, 128, 128)


def parse_line_pattern(pattern_str):
    """
    Chuyen chuoi kieu net (line pattern) thanh ten chuan hoa.
    Ten nay se duoc dung de tim LinePatternElement trong Revit.
    Tra ve ten chuan nhat co the dung.
    """
    if not pattern_str:
        return "Solid"
    mapping = {
        "solid"       : "Solid",
        "dash"        : "Dash",
        "dot"         : "Dot",
        "hidden"      : "Hidden",
        "long dash"   : "Long Dash",
        "triple dash" : "Triple Dash",
        "hidden 1 5"  : "Hidden 1.5",
        "hidden1.5"   : "Hidden 1.5",
    }
    key = str(pattern_str).strip().lower()
    return mapping.get(key, str(pattern_str).strip())

# endregion

# region --- Revit logic: tao Piping System ---

def get_existing_pipe_system_types():
    """
    Lay tat ca PipingSystemType hien co trong du an.
    Tuong thich Revit 2019+.
    Tra ve dict {ten_upper: element}.
    """
    result = {}
    try:
        # Phuong phap 1: dung FilteredElementCollector voi PipingSystemType class
        if HAS_PLUMBING:
            collector = FilteredElementCollector(doc).OfClass(PipingSystemType)
            for elem in collector:
                try:
                    name_upper = elem.Name.strip().upper()
                    result[name_upper] = elem
                except Exception:
                    pass
    except Exception:
        pass

    # Phuong phap 2: fallback qua BuiltInCategory (tuong thich hon)
    if not result:
        try:
            collector = FilteredElementCollector(doc) \
                .OfCategory(BuiltInCategory.OST_PipingSystem) \
                .WhereElementIsElementType()
            for elem in collector:
                try:
                    name_upper = elem.Name.strip().upper()
                    result[name_upper] = elem
                except Exception:
                    pass
        except Exception:
            pass

    return result


def get_line_pattern_id(doc, pattern_name):
    """
    Tim ElementId cua mot LinePatternElement dua vao ten.
    Tra ve ElementId neu tim thay, nguoc lai tra ve ElementId.InvalidElementId.
    """
    # Ten "Solid" la mot truong hop dac biet, ID cua no la InvalidElementId
    if not pattern_name or pattern_name.lower() == "solid":
        return ElementId.InvalidElementId

    # Tim kiem trong cac LinePatternElement hien co
    collector = FilteredElementCollector(doc).OfClass(LinePatternElement)
    for pattern in collector:
        if pattern.Name == pattern_name:
            return pattern.Id
            
    # Neu khong tim thay, ghi log va tra ve Invalid de Revit dung mac dinh (Solid)
    log_messages.append("  LinePattern warning: Pattern '{}' not found in project. Defaulting to Solid.".format(pattern_name))
    return ElementId.InvalidElementId


def set_system_color(system_elem, r, g, b):
    """
    Thiet lap mau sac cho PipingSystemType mot cach an toan va triet de.
    Dung property 'LineColor' truc tiep.
    """
    try:
        # Ep kieu System.Byte tuong minh de IronPython khong bị nham overload
        revit_color = RevitColor(
            System.Byte(r),
            System.Byte(g),
            System.Byte(b)
        )
        if hasattr(system_elem, "LineColor"):
            system_elem.LineColor = revit_color
            return True
    except Exception as ex:
        log_messages.append("  Color set warning: " + str(ex))
    return False


def _set_abbreviation(elem, sys_name, abbrev):
    """
    Ghi Abbreviation vao parameter Identity Data cua PipingSystemType.
    Chi dung LookupParameter vi get_Parameter(BuiltInParameter) tren PipingSystemType
    co the trigger loi 'expected Reference, got PipingSystemType' trong Revit 2024.
    """
    # Thu 1: LookupParameter theo ten chinh xac
    try:
        p = elem.LookupParameter("Abbreviation")
        if p is not None and not p.IsReadOnly:
            p.Set(abbrev)
            return
    except Exception:
        pass

    # Thu 2: Tim trong Parameters collection
    try:
        for p in elem.Parameters:
            try:
                if p.Definition.Name == "Abbreviation" and not p.IsReadOnly:
                    p.Set(abbrev)
                    return
            except Exception:
                continue
    except Exception:
        pass

    log_messages.append("  Abbreviation set warning for '{}': parameter not found or read-only.".format(sys_name))


def _set_line_weight(elem, sys_name, lineweight):
    """
    Ghi Line Weight cho PipingSystemType bang property truc tiep.
    Fallback qua LookupParameter neu property khong ton tai.
    """
    # Thu 1: property LineWeight truc tiep (Revit 2019+)
    try:
        if hasattr(elem, "LineWeight"):
            elem.LineWeight = lineweight
            return
    except Exception:
        pass

    # Thu 2: LookupParameter
    try:
        p = elem.LookupParameter("Line Weight")
        if p and not p.IsReadOnly:
            p.Set(lineweight)
            return
    except Exception as ex:
        log_messages.append("  LineWeight set warning for '{}': {}".format(sys_name, str(ex)))


def _set_line_pattern(elem, sys_name, pattern_id):
    """
    Ghi Line Pattern cho PipingSystemType.
    Nhan pattern_id da duoc resolve TRUOC khi vao Transaction.

    NGUYEN TAC:
    - KHONG goi FilteredElementCollector ben trong Transaction.
    - KHONG dung parameter.Set(ElementId) voi RBS_GRAPHICS_LINE_PATTERN_ID_PARAM
      (gay loi 'expected Reference, got PipingSystemType').
    - Chi dung property LinePatternId truc tiep.
    """
    # --- Thu 1: Property LinePatternId truc tiep tren MEPSystemType ---
    try:
        elem.LinePatternId = pattern_id
        return
    except Exception:
        pass

    # --- Thu 2: LookupParameter theo ten "Line Pattern" ---
    try:
        p = elem.LookupParameter("Line Pattern")
        if p is not None and not p.IsReadOnly:
            if pattern_id != ElementId.InvalidElementId:
                p.Set(pattern_id)
            return
    except Exception:
        pass

    # --- Thu 3: Duyet Parameters de tim parameter co ten tuong tu ---
    try:
        for p in elem.Parameters:
            try:
                pname = p.Definition.Name.lower()
                if "pattern" in pname and not p.IsReadOnly:
                    if pattern_id != ElementId.InvalidElementId:
                        p.Set(pattern_id)
                    return
            except Exception:
                continue
    except Exception:
        pass

    # --- Khong the set: ghi log canh bao, KHONG crash ---
    log_messages.append(
        "  LinePattern WARNING: Cannot set pattern for '{}'. Set manually via Revit UI.".format(sys_name)
    )


def create_or_update_piping_systems(rows, progress_callback=None):
    """
    Tao moi hoac cap nhat PipingSystemType trong Revit tu danh sach rows.
    Moi row la 1 dict voi cac key:
      TEN SYSTEM, ABBREVIATION, LINE WEIGHT (hoac DO DAY), LINE TYPE (hoac KIEU NET), RGB (hoac MA MAU RGB).

    Quy trinh:
    1. Lay danh sach system hien co.
    2. Tim system mau (template) de duplicate.
    3. Voi moi row: neu da ton tai -> cap nhat, chua ton tai -> duplicate tu template.
    4. Set cac thong so (Abbreviation, Graphic Overrides).

    Tra ve (so_tao_moi, so_cap_nhat, danh_sach_loi).
    """
    existing = get_existing_pipe_system_types()
    created  = 0
    updated  = 0
    errors   = []

    # --- PRE-BUILD: Tat ca FilteredElementCollector phai chay TRUOC khi mo Transaction ---
    # Goi collector ben trong Transaction se gay freeze/deadlock nghiem trong.

    # Tim template element de duplicate
    template_elem = None
    if HAS_PLUMBING:
        template_elem = FilteredElementCollector(doc).OfClass(PipingSystemType).FirstElement()
    if template_elem is None:
        template_elem = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_PipingSystem) \
            .WhereElementIsElementType().FirstElement()
    if template_elem is None:
        return 0, 0, ["No existing PipingSystemType found. Please create at least one manually."]

    # Build map: ten pattern (lower) -> ElementId  (chay TRUOC transaction)
    pattern_map = {}
    try:
        collector = FilteredElementCollector(doc).OfClass(LinePatternElement)
        for pat in collector:
            try:
                pattern_map[pat.Name.lower()] = pat.Id
                pattern_map[pat.Name]         = pat.Id  # them ca key goc (case-sensitive)
            except Exception:
                pass
    except Exception:
        pass

    # --- Helper: lookup pattern tu cache (khong dung collector) ---
    def _get_pattern_id_cached(name):
        if not name or name.lower() == "solid":
            return ElementId.InvalidElementId
        pid = pattern_map.get(name) or pattern_map.get(name.lower())
        if pid is None:
            log_messages.append("  LinePattern: '{}' not found, using Solid.".format(name))
            return ElementId.InvalidElementId
        return pid

    # Mo Transaction bao toan bo qua trinh
    t = Transaction(doc, "V2D - Create Piping Systems from Excel")
    t.Start()
    try:
        for i, row in enumerate(rows):
            # Bao cao tien do
            if progress_callback:
                progress_callback(i + 1, len(rows))

            # Lay cac gia tri tu row.
            # Ho tro ca 2 ten cot: ten moi (tieng Anh tu Excel) va ten cu (tieng Viet fallback).
            sys_name         = str(row.get("TEN SYSTEM", "")).strip()
            abbrev           = str(row.get("ABBREVIATION", "")).strip()
            lineweight_raw   = row.get("LINE WEIGHT", row.get("DO DAY", 2))
            line_pattern_str = str(row.get("LINE TYPE", row.get("KIEU NET", "Solid"))).strip()
            rgb_str          = str(row.get("RGB", row.get("MA MAU RGB", "128-128-128"))).strip()

            # Bo qua dong trong
            if not sys_name:
                continue

            # Parse line weight (do day net)
            try:
                lineweight = int(float(str(lineweight_raw))) if lineweight_raw else 2
                lineweight = max(1, min(16, lineweight))  # Revit cho phep 1-16
            except Exception:
                lineweight = 2

            # Parse mau sac RGB
            r, g, b = parse_rgb(rgb_str)

            # Parse ten kieu net (line pattern)
            line_pattern_name = parse_line_pattern(line_pattern_str)

            # Kiem tra system da ton tai chua
            sys_name_upper = sys_name.upper()
            target_elem = existing.get(sys_name_upper, None)

            try:
                # === STEP A: Tao / lay system element ===
                if target_elem is None:
                    # CHUA TON TAI: duplicate tu template va doi ten.
                    # Revit 2024 / IronPython: Element.Duplicate(string) doi khi bi
                    # mapping sang overload nhan Reference, gay ra loi.
                    # Dung CopyElements lam fallback.
                    new_elem_id = None

                    # Phuong phap 1: Duplicate truc tiep
                    try:
                        raw_result  = template_elem.Duplicate(sys_name)
                        # raw_result co the la ElementId hoac element tuy Revit version
                        if hasattr(raw_result, 'IntegerValue'):
                            new_elem_id = raw_result
                        elif raw_result is not None:
                            new_elem_id = raw_result.Id
                    except Exception:
                        pass

                    # Phuong phap 2: CopyElements (an toan hon) neu Phuong phap 1 that bai
                    if new_elem_id is None:
                        try:
                            from Autodesk.Revit.DB import CopyPasteOptions, ElementTransformUtils
                            import System.Collections.Generic as SCG
                            id_list   = SCG.List[ElementId]([template_elem.Id])
                            copied    = ElementTransformUtils.CopyElements(
                                doc, id_list, doc, None, CopyPasteOptions()
                            )
                            if copied and copied.Count > 0:
                                new_elem_id = list(copied)[0]
                        except Exception as copy_ex:
                            raise Exception("Cannot create system type '{}': {}".format(sys_name, str(copy_ex)))

                    if new_elem_id is None:
                        raise Exception("Both Duplicate and CopyElements returned None for: " + sys_name)

                    target_elem = doc.GetElement(new_elem_id)
                    if target_elem is None:
                        raise Exception("GetElement returned None after duplicate for: " + sys_name)

                    # Doi ten neu can (CopyElements tao ra ten co "_Copy")
                    try:
                        if target_elem.Name != sys_name:
                            target_elem.Name = sys_name
                    except Exception:
                        pass

                    existing[sys_name_upper] = target_elem
                    created += 1
                    log_messages.append("  CREATED: " + sys_name)
                else:
                    updated += 1
                    log_messages.append("  UPDATED: " + sys_name)

                # === STEP B: ABBREVIATION ===
                # (da co try/except ben trong)
                if abbrev:
                    _set_abbreviation(target_elem, sys_name, abbrev)

                # === STEP C: COLOR ===
                # (da co try/except ben trong)
                set_system_color(target_elem, r, g, b)

                # === STEP D: LINE WEIGHT ===
                # (da co try/except ben trong)
                _set_line_weight(target_elem, sys_name, lineweight)

                # === STEP E: LINE PATTERN ===
                # Lay pattern_id tu CACHE da build truoc transaction (khong dung collector)
                resolved_pattern_id = _get_pattern_id_cached(line_pattern_name)
                _set_line_pattern(target_elem, sys_name, resolved_pattern_id)

                # Ghi log ket qua
                log_messages.append("    -> Abbrev: '{}' | RGB({},{},{}) | Weight:{} | Pattern:{}".format(
                    abbrev, r, g, b, lineweight, line_pattern_name
                ))

            except Exception as row_ex:
                err_msg = "  ERROR processing row for '{}': {}".format(sys_name, str(row_ex))
                errors.append(err_msg)
                log_messages.append(err_msg)

        t.Commit()
    except Exception as ex:
        t.RollBack()
        errors.append("Transaction failed: " + str(ex))
        return 0, 0, errors

    return created, updated, errors

# endregion

# region --- UI Constants ---
# Bảng mau tham khao tu createElbow.py (dark purple theme cua V2D)
C_PURPLE_DK = Color.FromArgb(255,  30,  10,  70)   # nen header / footer
C_PURPLE    = Color.FromArgb(255,  75,  35, 155)   # nut chinh
C_PURPLE_LT = Color.FromArgb(255, 160, 130, 220)   # chu mo / subtitle
C_PURPLE_MID= Color.FromArgb(255,  50,  22, 110)   # nen card
C_GREEN     = Color.FromArgb(255,  30, 185,  90)   # trang thai thanh cong / accent
C_ORANGE    = Color.FromArgb(255, 235, 145,  30)   # canh bao
C_RED       = Color.FromArgb(255, 220,  55,  55)   # loi
C_BG        = Color.FromArgb(255,  16,  10,  36)   # nen chinh form
C_CARD      = Color.FromArgb(255,  28,  18,  58)   # nen cac section card
C_TEXT      = Color.FromArgb(255, 225, 215, 255)   # chu chinh
C_TEXT_DIM  = Color.FromArgb(255, 140, 120, 185)   # chu phu / mo
C_SEP       = Color.FromArgb(255,  60,  40, 100)   # duong ke phan cach
C_BTN_TEXT  = Color.FromArgb(255, 255, 255, 255)   # chu tren nut

# Font chut
F_HDR    = Font("Segoe UI", 13, FontStyle.Bold,    GraphicsUnit.Point)
F_STEP   = Font("Segoe UI",  9, FontStyle.Bold,    GraphicsUnit.Point)
F_MAIN   = Font("Segoe UI",  9, FontStyle.Regular, GraphicsUnit.Point)
F_BTN    = Font("Segoe UI",  9, FontStyle.Bold,    GraphicsUnit.Point)
F_INFO   = Font("Segoe UI",  8, FontStyle.Regular, GraphicsUnit.Point)
F_SMALL  = Font("Segoe UI",  7, FontStyle.Regular, GraphicsUnit.Point)
F_LOGO   = Font("Segoe UI",  9, FontStyle.Bold,    GraphicsUnit.Point)
F_MONO   = Font("Consolas",  8, FontStyle.Regular, GraphicsUnit.Point)

# endregion

# region --- Helper tao Widget ---

def _btn(text, bg, fg, x, y, w, h, font=None):
    """Tao Button co style thong nhat."""
    b = Button()
    b.Text      = text
    b.Font      = font if font else F_BTN
    b.ForeColor = fg
    b.BackColor = bg
    b.FlatStyle = FlatStyle.Flat
    b.FlatAppearance.BorderSize  = 0
    b.FlatAppearance.MouseOverBackColor = Color.FromArgb(
        min(255, bg.R + 20), min(255, bg.G + 20), min(255, bg.B + 20)
    )
    b.Location  = Point(x, y)
    b.Size      = Size(w, h)
    b.Cursor    = WinForms.Cursors.Hand
    return b


def _lbl(text, font, fg, x, y, w, h, align=ContentAlignment.MiddleLeft, bg=None):
    """Tao Label co style thong nhat."""
    lb = Label()
    lb.Text      = text
    lb.Font      = font
    lb.ForeColor = fg
    lb.BackColor = bg if bg else Color.Transparent
    lb.Location  = Point(x, y)
    lb.Size      = Size(w, h)
    lb.TextAlign = align
    return lb


def _sep(x, y, w):
    """Tao duong ke ngang phan cach."""
    p = Panel()
    p.BackColor = C_SEP
    p.Location  = Point(x, y)
    p.Size      = Size(w, 1)
    return p

# endregion

# region --- Preview Panel (hien thi du lieu Excel) ---

class PreviewRow:
    """Chua thong tin 1 dong du lieu tu Excel de hien thi trong grid."""
    def __init__(self, stt, name, abbrev, weight, pattern, rgb_str, color_tuple):
        self.stt     = stt
        self.name    = name
        self.abbrev  = abbrev
        self.weight  = weight
        self.pattern = pattern
        self.rgb_str = rgb_str
        self.color   = color_tuple  # (r, g, b)

# endregion

# region --- Main Form ---

class CreatePipingSystemForm(Form):
    """
    Form chinh cho tool 'Create Piping System from Excel'.
    Cac buoc:
      1. User chon file Excel.
      2. Tool doc sheet 'Piping System' (hoac sheet do user chon).
      3. Hien thi preview du lieu.
      4. User bam 'Create in Revit' de ap dung.
    """

    # Hang so layout
    FW  = 520     # chieu rong form
    PAD = 14      # padding ngoai
    INN = 10      # padding trong card

    def __init__(self):
        self._excel_path   = None   # duong dan file Excel da chon
        self._sheet_name   = "Piping System"  # ten sheet mac dinh
        self._preview_rows = []     # danh sach PreviewRow tu Excel
        self._all_sheets   = []     # danh sach tat ca sheet trong workbook
        Form.__init__(self)
        self._build_ui()

    # -------------------------------------------------------------------------
    # Xay dung giao dien
    # -------------------------------------------------------------------------
    def _build_ui(self):
        FW  = self.FW
        PAD = self.PAD
        INN = self.INN

        self.SuspendLayout()

        # ===== HEADER =====
        HDR_H = 62
        hdr = Panel()
        hdr.BackColor = C_PURPLE_DK
        hdr.Location  = Point(0, 0)
        hdr.Size      = Size(FW, HDR_H)

        # Icon chu "P" trong o vuong mau xanh la
        icon = Label()
        icon.Text      = "P"
        icon.Font      = Font("Segoe UI", 16, FontStyle.Bold, GraphicsUnit.Point)
        icon.ForeColor = Color.White
        icon.BackColor = C_GREEN
        icon.Location  = Point(PAD, (HDR_H - 38) // 2)
        icon.Size      = Size(38, 38)
        icon.TextAlign = ContentAlignment.MiddleCenter
        hdr.Controls.Add(icon)

        # Ten cong cu
        hdr.Controls.Add(_lbl(
            "Create Piping System", F_HDR, Color.White,
            PAD + 48, 8, 340, 26
        ))
        # Mo ta ngan
        hdr.Controls.Add(_lbl(
            "Import piping systems from Excel  |  V2D Tools",
            F_SMALL, C_PURPLE_LT,
            PAD + 48, 36, 380, 16
        ))

        # ===== SECTION 1: Chon file Excel =====
        Y_S1  = HDR_H + 10
        H_S1  = 110
        s1 = self._make_card(PAD, Y_S1, FW - PAD * 2, H_S1)

        s1.Controls.Add(_lbl(
            "STEP 1  Select Excel File", F_STEP, C_GREEN,
            INN, 8, 300, 22
        ))

        # Thanh hien thi duong dan file
        self._txt_path = TextBox()
        self._txt_path.Text        = "No file selected..."
        self._txt_path.Font        = F_INFO
        self._txt_path.ForeColor   = C_TEXT_DIM
        self._txt_path.BackColor   = Color.FromArgb(255, 20, 12, 50)
        self._txt_path.BorderStyle = BorderStyle.FixedSingle
        self._txt_path.ReadOnly    = True
        self._txt_path.Location    = Point(INN, 36)
        self._txt_path.Size        = Size(FW - PAD * 2 - INN * 2 - 90, 26)
        s1.Controls.Add(self._txt_path)

        # Nut Browse
        btn_browse = _btn("Browse...", C_PURPLE, C_BTN_TEXT,
                          FW - PAD * 2 - INN - 82, 35, 82, 28)
        btn_browse.Click += self._on_browse
        s1.Controls.Add(btn_browse)

        # Combobox chon sheet
        s1.Controls.Add(_lbl("Worksheet:", F_MAIN, C_TEXT, INN, 70, 80, 24))

        self._cmb_sheet = ComboBox()
        self._cmb_sheet.Font        = F_INFO
        self._cmb_sheet.ForeColor   = C_TEXT
        self._cmb_sheet.BackColor   = Color.FromArgb(255, 20, 12, 50)
        self._cmb_sheet.FlatStyle   = FlatStyle.Flat
        self._cmb_sheet.DropDownStyle = ComboBoxStyle.DropDownList
        self._cmb_sheet.Location    = Point(INN + 86, 68)
        self._cmb_sheet.Size        = Size(FW - PAD * 2 - INN * 2 - 86 - 90, 26)
        self._cmb_sheet.SelectedIndexChanged += self._on_sheet_changed
        s1.Controls.Add(self._cmb_sheet)

        # Nut Load Sheet
        self._btn_load = _btn("Load", C_PURPLE, C_BTN_TEXT,
                              FW - PAD * 2 - INN - 82, 68, 82, 28)
        self._btn_load.Click   += self._on_load_sheet
        self._btn_load.Enabled  = False
        s1.Controls.Add(self._btn_load)

        # ===== SECTION 2: Preview du lieu =====
        Y_S2  = Y_S1 + H_S1 + 8
        H_S2  = 220
        s2 = self._make_card(PAD, Y_S2, FW - PAD * 2, H_S2)

        s2.Controls.Add(_lbl(
            "STEP 2  Preview Data", F_STEP, C_GREEN,
            INN, 8, 300, 22
        ))

        # Label dem so dong
        self._lbl_count = _lbl(
            "No data loaded.", F_INFO, C_TEXT_DIM,
            INN, 32, FW - PAD * 2 - INN * 2, 18
        )
        s2.Controls.Add(self._lbl_count)

        # Panel chua luoi hien thi mau sac preview
        self._pnl_preview = Panel()
        self._pnl_preview.BackColor  = Color.FromArgb(255, 20, 12, 50)
        self._pnl_preview.Location   = Point(INN, 54)
        self._pnl_preview.Size       = Size(FW - PAD * 2 - INN * 2, H_S2 - 64)
        self._pnl_preview.AutoScroll = True
        self._pnl_preview.BorderStyle= BorderStyle.None
        s2.Controls.Add(self._pnl_preview)

        # ===== SECTION 3: Thuc thi =====
        Y_S3  = Y_S2 + H_S2 + 8
        H_S3  = 90
        s3 = self._make_card(PAD, Y_S3, FW - PAD * 2, H_S3)

        s3.Controls.Add(_lbl(
            "STEP 3  Create in Revit", F_STEP, C_GREEN,
            INN, 8, 300, 22
        ))

        # Progress bar
        self._progress = ProgressBar()
        self._progress.Location = Point(INN, 36)
        self._progress.Size     = Size(FW - PAD * 2 - INN * 2, 12)
        self._progress.Minimum  = 0
        self._progress.Maximum  = 100
        self._progress.Value    = 0
        self._progress.Style    = WinForms.ProgressBarStyle.Continuous
        s3.Controls.Add(self._progress)

        # Nut tao system
        self._btn_create = _btn(
            "  Create Piping Systems in Revit",
            C_GREEN, C_BG,
            INN, 54, FW - PAD * 2 - INN * 2, 30
        )
        self._btn_create.Enabled = False
        self._btn_create.Click  += self._on_create
        s3.Controls.Add(self._btn_create)

        # ===== STATUS =====
        Y_STATUS = Y_S3 + H_S3 + 6
        H_STATUS = 24
        self._lbl_status = _lbl(
            "Ready. Select an Excel file to begin.",
            F_INFO, C_TEXT_DIM,
            PAD, Y_STATUS, FW - PAD * 2, H_STATUS
        )
        self._lbl_status.AutoSize = False

        # ===== FOOTER =====
        FTRH   = 34
        Y_FTR  = Y_STATUS + H_STATUS + 6
        FH     = Y_FTR + FTRH

        ftr = Panel()
        ftr.BackColor = C_PURPLE_DK
        ftr.Location  = Point(0, Y_FTR)
        ftr.Size      = Size(FW, FTRH)

        # Logo V2D (goc trai)
        lbl_logo = Label()
        lbl_logo.Text      = "V2D"
        lbl_logo.Font      = F_LOGO
        lbl_logo.ForeColor = C_GREEN
        lbl_logo.BackColor = Color.FromArgb(255, 20, 8, 50)
        lbl_logo.Location  = Point(PAD, (FTRH - 22) // 2)
        lbl_logo.Size      = Size(38, 22)
        lbl_logo.TextAlign = ContentAlignment.MiddleCenter
        ftr.Controls.Add(lbl_logo)

        # Mo ta footer
        ftr.Controls.Add(_lbl(
            "V2D Tools  |  Create Piping System v1.0",
            F_SMALL, C_TEXT_DIM,
            PAD + 46, (FTRH - 16) // 2, 280, 16
        ))

        # ===== Cai dat Form =====
        self.AutoScaleDimensions = Drawing.SizeF(96, 96)
        self.AutoScaleMode       = AutoScaleMode.Dpi
        self.ClientSize          = Size(FW, FH)
        self.FormBorderStyle     = FormBorderStyle.FixedSingle
        self.StartPosition       = FormStartPosition.CenterScreen
        self.BackColor           = C_BG
        self.Font                = F_MAIN
        self.Text                = "Create Piping System  |  V2D Tools"
        self.MaximizeBox         = False
        self.TopMost             = True

        # Them tat ca control vao form
        self.Controls.Add(hdr)
        self.Controls.Add(s1)
        self.Controls.Add(s2)
        self.Controls.Add(s3)
        self.Controls.Add(self._lbl_status)
        self.Controls.Add(ftr)

        self.ResumeLayout(False)
        self.PerformLayout()

    def _make_card(self, x, y, w, h):
        """Tao Panel nen card co mau nen C_CARD."""
        p = Panel()
        p.BackColor = C_CARD
        p.Location  = Point(x, y)
        p.Size      = Size(w, h)
        return p

    # -------------------------------------------------------------------------
    # Su kien: Browse file Excel
    # -------------------------------------------------------------------------
    def _on_browse(self, sender, e):
        """
        Mo hop thoai chon file Excel.
        Sau khi chon, doc danh sach sheet va dien vao combobox.
        """
        dlg = OpenFileDialog()
        dlg.Title  = "Select Excel File"
        dlg.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
        dlg.InitialDirectory = os.path.expanduser("~")

        if dlg.ShowDialog() != DialogResult.OK:
            return

        self._excel_path = dlg.FileName
        # Hien thi duong dan rut gon neu qua dai
        display_path = dlg.FileName
        if len(display_path) > 60:
            display_path = "..." + display_path[-57:]
        self._txt_path.Text      = display_path
        self._txt_path.ForeColor = C_TEXT

        # Doc danh sach sheet
        self._load_sheet_list()

    def _load_sheet_list(self):
        """Doc danh sach sheet tu workbook va dien vao combobox."""
        self._set_status("Reading workbook...", C_PURPLE_LT)
        self.Refresh()

        # Doc danh sach sheet bang OLEDB
        _, sheet_names, err = read_excel_data(self._excel_path, sheet_name=None)

        if err:
            self._set_status("ERROR: " + err, C_RED)
            MessageBox.Show(err, "Excel Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        if not sheet_names:
            self._set_status("No worksheets found or file cannot be read.", C_ORANGE)
            return

        self._all_sheets = sheet_names

        # Dien combobox
        self._cmb_sheet.Items.Clear()
        for s in sheet_names:
            self._cmb_sheet.Items.Add(s)

        # Chon "Piping System" mac dinh neu co, nguoc lai chon sheet dau tien
        default_idx = 0
        for i, s in enumerate(sheet_names):
            if s.strip().upper() == "PIPING SYSTEM":
                default_idx = i
                break

        self._cmb_sheet.SelectedIndex = default_idx
        self._sheet_name = sheet_names[default_idx]

        # Neu khong tim thay sheet mac dinh, nhac user
        if sheet_names[default_idx].strip().upper() != "PIPING SYSTEM":
            self._set_status(
                "Sheet 'Piping System' not found. Please select the correct sheet.",
                C_ORANGE
            )
        else:
            self._set_status("Workbook loaded. Click 'Load' to preview data.", C_TEXT_DIM)

        self._btn_load.Enabled = True

    # -------------------------------------------------------------------------
    # Su kien: Thay doi sheet trong combobox
    # -------------------------------------------------------------------------
    def _on_sheet_changed(self, sender, e):
        """Cap nhat ten sheet khi user thay doi combobox."""
        if self._cmb_sheet.SelectedItem is not None:
            self._sheet_name = str(self._cmb_sheet.SelectedItem)

    # -------------------------------------------------------------------------
    # Su kien: Load sheet va hien thi preview
    # -------------------------------------------------------------------------
    def _on_load_sheet(self, sender, e):
        """
        Doc du lieu tu sheet da chon va hien thi preview.
        """
        if not self._excel_path or not self._sheet_name:
            self._set_status("Please select a file and sheet first.", C_ORANGE)
            return

        self._set_status("Reading sheet '{}'...".format(self._sheet_name), C_PURPLE_LT)
        self.Refresh()

        rows, _, err = read_excel_data(self._excel_path, self._sheet_name)

        if err:
            self._set_status("ERROR: " + err, C_RED)
            MessageBox.Show(err, "Excel Read Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            self._preview_rows = []
            self._render_preview()
            self._lbl_count.Text = "Data loading failed."
            self._lbl_count.ForeColor = C_ORANGE
            self._btn_create.Enabled = False
            return

        if not rows:
            self._set_status("Sheet is empty or has no valid data rows.", C_ORANGE)
            self._btn_create.Enabled = False
            return

        # Chuyen doi rows thanh PreviewRow
        self._preview_rows = []
        for row in rows:
            # Linh hoat xu ly ten cot: uu tien ten moi (tieng Anh), sau do den ten cu (tieng Viet).
            # Cac key tu Excel da duoc chuyen thanh chu HOA boi ham read_excel_data.
            sys_name = str(row.get("SYSTEM NAME", row.get("TEN SYSTEM", ""))).strip()
            if not sys_name:
                continue

            # Abbreviation khong co trong format moi, chi lay neu co.
            abbrev   = str(row.get("ABBREVIATION", "")).strip()

            weight   = row.get("LINE WEIGHT", row.get("DO DAY", 2))
            pattern  = str(row.get("LINE TYPE", row.get("KIEU NET", "Solid"))).strip()
            rgb_str  = str(row.get("RGB", row.get("MA MAU RGB", "128-128-128"))).strip()

            # STT khong co trong format moi.
            stt_val  = row.get("STT", "")

            color_t  = parse_rgb(rgb_str)
            self._preview_rows.append(
                PreviewRow(stt_val, sys_name, abbrev, weight, pattern, rgb_str, color_t)
            )

        # Ve lai preview
        self._render_preview()

        count = len(self._preview_rows)
        self._lbl_count.Text = "{} piping system(s) found.".format(count)
        self._lbl_count.ForeColor = C_GREEN if count > 0 else C_ORANGE

        if count > 0:
            self._btn_create.Enabled = True
            self._set_status(
                "{} systems ready. Click 'Create Piping Systems in Revit'.".format(count),
                C_TEXT_DIM
            )
        else:
            self._btn_create.Enabled = False
            self._set_status("No valid rows found. Check column names in Excel.", C_ORANGE)

    def _render_preview(self):
        """
        Ve bang preview nho trong pnl_preview.
        Moi dong hien thi: [mau] Ten System | Viet tat | Kieu net | Do day
        """
        self._pnl_preview.Controls.Clear()

        ROW_H   = 22     # chieu cao moi dong
        SWATCH  = 18     # chieu rong o mau sac
        PAD_X   = 4
        y       = 2

        PW = self._pnl_preview.Width  # chieu rong panel

        for pr in self._preview_rows:
            # Panel chua 1 dong
            row_pnl = Panel()
            row_pnl.BackColor  = Color.Transparent
            row_pnl.Location   = Point(2, y)
            row_pnl.Size       = Size(PW - 4, ROW_H)

            # O mau sac
            swatch = Panel()
            swatch.BackColor = Color.FromArgb(255, pr.color[0], pr.color[1], pr.color[2])
            swatch.Location  = Point(PAD_X, (ROW_H - 14) // 2)
            swatch.Size      = Size(SWATCH, 14)
            row_pnl.Controls.Add(swatch)

            # Ten system
            x_name = PAD_X + SWATCH + 6
            lbl_name = Label()
            lbl_name.Text      = pr.name
            lbl_name.Font      = F_INFO
            lbl_name.ForeColor = C_TEXT
            lbl_name.BackColor = Color.Transparent
            lbl_name.Location  = Point(x_name, 2)
            lbl_name.Size      = Size(200, ROW_H - 4)
            lbl_name.TextAlign = ContentAlignment.MiddleLeft
            row_pnl.Controls.Add(lbl_name)

            # Abbreviation
            lbl_abbrev = Label()
            lbl_abbrev.Text      = pr.abbrev
            lbl_abbrev.Font      = F_SMALL
            lbl_abbrev.ForeColor = C_TEXT_DIM
            lbl_abbrev.BackColor = Color.Transparent
            lbl_abbrev.Location  = Point(x_name + 206, 2)
            lbl_abbrev.Size      = Size(54, ROW_H - 4)
            lbl_abbrev.TextAlign = ContentAlignment.MiddleLeft
            row_pnl.Controls.Add(lbl_abbrev)

            # RGB
            lbl_rgb = Label()
            lbl_rgb.Text      = pr.rgb_str
            lbl_rgb.Font      = F_SMALL
            lbl_rgb.ForeColor = C_TEXT_DIM
            lbl_rgb.BackColor = Color.Transparent
            lbl_rgb.Location  = Point(x_name + 266, 2)
            lbl_rgb.Size      = Size(100, ROW_H - 4)
            lbl_rgb.TextAlign = ContentAlignment.MiddleLeft
            row_pnl.Controls.Add(lbl_rgb)

            self._pnl_preview.Controls.Add(row_pnl)
            y += ROW_H

    # -------------------------------------------------------------------------
    # Su kien: Tao Piping System trong Revit
    # -------------------------------------------------------------------------
    def _on_create(self, sender, e):
        """
        Bat dau qua trinh tao PipingSystemType trong Revit.
        Sau khi xong hien ket qua (so tao moi, so cap nhat, loi neu co).
        """
        if not self._preview_rows:
            self._set_status("No data to process.", C_ORANGE)
            return

        self._btn_create.Enabled = False
        self._progress.Value     = 0
        self._set_status("Creating piping systems in Revit...", C_PURPLE_LT)
        self.Refresh()

        # Chuyen doi PreviewRow sang dict de goi ham backend
        rows_for_revit = []
        for pr in self._preview_rows:
            rows_for_revit.append({
                "TEN SYSTEM"  : pr.name,
                "ABBREVIATION": pr.abbrev,
                "LINE WEIGHT" : pr.weight,   # key moi khop voi backend
                "LINE TYPE"   : pr.pattern,  # key moi khop voi backend
                "RGB"         : pr.rgb_str,  # key moi khop voi backend
            })

        total = len(rows_for_revit)
        created_count  = [0]
        updated_count  = [0]
        errors_list    = []

        # Ham callback cap nhat progress bar
        # KHONG goi self.Refresh() o day vi se conflict voi Revit transaction thread
        def progress_cb(current, total_rows):
            pct = int(current * 100 / total_rows)
            self._progress.Value = min(pct, 100)

        try:
            n_created, n_updated, errors_list = create_or_update_piping_systems(
                rows_for_revit, progress_callback=progress_cb
            )
        except Exception as ex:
            errors_list = [str(ex)]
            n_created, n_updated = 0, 0

        self._progress.Value = 100

        # Hien ket qua
        if errors_list:
            summary = "Done with errors. Created: {}, Updated: {}, Errors: {}.".format(
                n_created, n_updated, len(errors_list)
            )
            self._set_status(summary, C_ORANGE)
            # Hien chi tiet loi
            err_msg = "\n".join(errors_list[:10])
            if len(errors_list) > 10:
                err_msg += "\n... and {} more errors.".format(len(errors_list) - 10)
            MessageBox.Show(
                err_msg,
                "Errors Occurred",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
        else:
            summary = "Success! Created: {}, Updated: {}.".format(n_created, n_updated)
            self._set_status(summary, C_GREEN)

        log_messages.insert(0, summary)
        self._btn_create.Enabled = True

    # -------------------------------------------------------------------------
    # Helper: cap nhat label trang thai
    # -------------------------------------------------------------------------
    def _set_status(self, msg, color):
        """Cap nhat noi dung va mau sac cua label trang thai."""
        self._lbl_status.Text      = msg
        self._lbl_status.ForeColor = color

# endregion

# region --- Chay chuong trinh ---
form = CreatePipingSystemForm()
form.ShowDialog()  # Dung ShowDialog thay Application.Run de tranh conflict voi Revit message loop
OUT = log_messages
# endregion