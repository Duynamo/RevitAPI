"""Pipe Accessory Placer Tool - @FVC"""
#region ___imports
import clr, System, math, re

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB.Plumbing import *

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

#region ___doc/app/ui
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
uidoc = uiapp.ActiveUIDocument
#endregion

#region ___selection filters
class PipeFilter(ISelectionFilter):
    """Filter chi danh cho Pipe element"""
    def AllowElement(self, e):
        return (e.Category is not None and
                e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeCurves))
    def AllowReference(self, r, pt): return False

class PipeFaceFilter(ISelectionFilter):
    """
    Filter cho phep pick diem tren be mat ong (PickObject PointOnElement).
    AllowElement: chi chap nhan Pipe.
    AllowReference: chi chap nhan face (Reference.ElementReferenceType = SURFACE).
    """
    def AllowElement(self, e):
        return (e.Category is not None and
                e.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeCurves))
    def AllowReference(self, r, pt):
        try:
            return r.ElementReferenceType == ElementReferenceType.REFERENCE_TYPE_SURFACE
        except:
            return True
#endregion

#region ___helpers

def getPipeDiameter_mm(pipe):
    p = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
    return round(p.AsDouble() * 304.8, 1) if p else 0.0

# ─────────────────────────────────────────────────────────────
# Lay DN tu instances thuc te trong project
# ─────────────────────────────────────────────────────────────
def _buildSymbolDnMap():
    """symbolId(int) -> set{dn_mm} tu instances dang co trong project"""
    sym_map = {}
    instances = (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_PipeAccessory)
                 .WhereElementIsNotElementType()
                 .ToElements())
    for inst in instances:
        try:
            sid = inst.GetTypeId().IntegerValue
            cm  = inst.MEPModel.ConnectorManager
            if cm is None: continue
            for c in cm.Connectors:
                try:
                    if c.ConnectorType == ConnectorType.End and c.Radius > 0:
                        dn = round(c.Radius * 2 * 304.8, 1)
                        if sid not in sym_map:
                            sym_map[sid] = set()
                        sym_map[sid].add(dn)
                except: pass
        except: pass
    return sym_map

def _dnFromSymbolParams(symbol):
    """Fallback: doc parameter DN tren FamilySymbol"""
    NAMES = ["Nominal Diameter", "DN", "Size", "Diameter",
             "Nominal Size", "NPS", "ND", "Pipe Diameter"]
    for name in NAMES:
        p = symbol.LookupParameter(name)
        if p is None: continue
        try:
            if p.StorageType == StorageType.Double and p.AsDouble() > 0:
                return round(p.AsDouble() * 304.8, 1)
            if p.StorageType == StorageType.String:
                nums = re.findall(r'\d+\.?\d*', p.AsString() or '')
                if nums: return float(nums[0])
        except: pass
    return None

def getSymbolDisplayName(symbol):
    try:
        fname = symbol.Family.Name
        tname = symbol.get_Parameter(
            BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        return "{0} : {1}".format(fname, tname)
    except:
        try: return symbol.Name
        except: return "ID_{0}".format(symbol.Id.IntegerValue)

class SymbolWrapper(object):
    def __init__(self, symbol, dn_set=None):
        self.Symbol   = symbol
        self.Name     = getSymbolDisplayName(symbol)
        self.DN       = dn_set or set()
        self.Resolved = bool(dn_set)
    def __str__(self): return self.Name

def loadAccessoriesByDN(dn_mm, tol=1.0):
    inst_map   = _buildSymbolDnMap()
    collector  = (FilteredElementCollector(doc)
                  .OfCategory(BuiltInCategory.OST_PipeAccessory)
                  .WhereElementIsElementType()
                  .ToElements())
    matched    = []
    unresolved = []
    for sym in collector:
        if not isinstance(sym, FamilySymbol): continue
        try:
            sid    = sym.Id.IntegerValue
            dn_set = inst_map.get(sid, set())
            if not dn_set:
                d = _dnFromSymbolParams(sym)
                if d: dn_set = {d}
            wrapper = SymbolWrapper(sym, dn_set if dn_set else None)
            if not dn_set:
                wrapper.Name = "[?] " + wrapper.Name
                unresolved.append(wrapper)
            elif any(abs(d - dn_mm) <= tol for d in dn_set):
                matched.append(wrapper)
        except:
            w      = SymbolWrapper(sym)
            w.Name = "[?] " + w.Name
            unresolved.append(w)
    matched.sort(key=lambda w: w.Name.lower())
    unresolved.sort(key=lambda w: w.Name.lower())
    return matched + unresolved

# ─────────────────────────────────────────────────────────────
# Chieu diem len duong tam ong
# ─────────────────────────────────────────────────────────────
def projectOntoCenterline(pipe, point):
    """
    Chieu XYZ point len Location.Curve (duong tam ong).
    Tra ve XYZ chinh xac tren truc ong.
    """
    curve = pipe.Location.Curve
    try:
        result = curve.Project(point)
        if result is not None:
            return result.XYZPoint
    except: pass
    # fallback: tim diem gan nhat tren doan thang
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    v  = p1 - p0
    w  = point - p0
    t  = w.DotProduct(v) / v.DotProduct(v)
    t  = max(0.0, min(1.0, t))
    return p0 + v.Multiply(t)

# ─────────────────────────────────────────────────────────────
# Dat accessory + xoay dung huong pipe
# ─────────────────────────────────────────────────────────────
def placeAccessory(pipe, insert_pt, symbol):
    """
    insert_pt: XYZ da chieu len duong tam ong
    Dat FamilyInstance tai insert_pt, xoay can huong pipe.
    """
    TransactionManager.Instance.EnsureInTransaction(doc)
    try:
        if not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()

        pipe_dir = pipe.Location.Curve.Direction.Normalize()

        # Level tu pipe
        lvl_id = pipe.get_Parameter(
            BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId()
        level  = doc.GetElement(lvl_id)

        # Dat instance
        instance = doc.Create.NewFamilyInstance(
            insert_pt, symbol, level, StructuralType.NonStructural)
        doc.Regenerate()

        # ── Xoay can huong pipe ──────────────────────────────
        # Lay huong connector mac dinh cua instance
        default_dir = None
        try:
            for c in instance.MEPModel.ConnectorManager.Connectors:
                if c.ConnectorType == ConnectorType.End:
                    default_dir = c.CoordinateSystem.BasisZ.Normalize()
                    break
        except: pass
        if default_dir is None:
            default_dir = XYZ.BasisX

        # Truc xoay: BasisZ (ong nam ngang), BasisX (ong thang dung)
        if abs(pipe_dir.DotProduct(XYZ.BasisZ)) > 0.99:
            rot_axis = XYZ.BasisX
        else:
            rot_axis = XYZ.BasisZ

        rot_line = Line.CreateBound(insert_pt, insert_pt.Add(rot_axis))

        def flatProj(v, n):
            r = v - n.Multiply(v.DotProduct(n))
            return r if r.GetLength() > 1e-6 else None

        d_f = flatProj(default_dir, rot_axis)
        p_f = flatProj(pipe_dir,    rot_axis)

        if d_f and p_f:
            d_f   = d_f.Normalize()
            p_f   = p_f.Normalize()
            angle = d_f.AngleTo(p_f)
            if d_f.CrossProduct(p_f).DotProduct(rot_axis) < 0:
                angle = -angle
            if abs(angle) > 1e-6:
                ElementTransformUtils.RotateElement(
                    doc, instance.Id, rot_line, angle)
                doc.Regenerate()

        TransactionManager.Instance.TransactionTaskDone()
        return instance, None

    except Exception as ex:
        try: TransactionManager.Instance.TransactionTaskDone()
        except: pass
        return None, str(ex)
#endregion

#region ___UI
class AccessoryPlacerForm(Form):
    W, H   = 560, 630
    PAD    = 12
    BTN_W  = 88
    BTN_H  = 34

    def __init__(self):
        self._pipe        = None
        self._dn_mm       = 0.0
        self._symbols     = []
        self._insertPt    = None
        self._logLines    = []
        self._clb_lastIdx = -1
        self.InitializeComponent()

    def InitializeComponent(self):
        p, w, h = self.PAD, self.W, self.H

        # ── Title ─────────────────────────────────────────────
        self._lb_title           = Label()
        self._lb_title.Text      = "Pipe Accessory Placer"
        self._lb_title.Font      = Font("Meiryo UI", 11,
                                        System.Drawing.FontStyle.Bold)
        self._lb_title.ForeColor = Color.Red
        self._lb_title.Location  = Point(p, 10)
        self._lb_title.Size      = Size(w - p*2, 24)

        # ── GroupBox: Pipe ─────────────────────────────────────
        self._grb_pipe           = GroupBox()
        self._grb_pipe.Text      = "Pipe"
        self._grb_pipe.Font      = Font("Meiryo UI", 8,
                                        System.Drawing.FontStyle.Bold)
        self._grb_pipe.ForeColor = Color.Red
        self._grb_pipe.Location  = Point(p, 38)
        self._grb_pipe.Size      = Size(w - p*2, 54)

        self._btt_pickPipe           = Button()
        self._btt_pickPipe.Text      = "Pick Pipe"
        self._btt_pickPipe.Font      = Font("Meiryo UI", 8,
                                            System.Drawing.FontStyle.Bold)
        self._btt_pickPipe.ForeColor = Color.Blue
        self._btt_pickPipe.Location  = Point(6, 18)
        self._btt_pickPipe.Size      = Size(90, 26)
        self._btt_pickPipe.UseVisualStyleBackColor = True
        self._btt_pickPipe.Click    += self.Btt_pickPipeClick

        self._lb_pipe            = Label()
        self._lb_pipe.Text       = "No pipe selected"
        self._lb_pipe.Font       = Font("Meiryo UI", 8,
                                        System.Drawing.FontStyle.Bold)
        self._lb_pipe.ForeColor  = Color.DimGray
        self._lb_pipe.Location   = Point(104, 22)
        self._lb_pipe.Size       = Size(w - p*2 - 108, 20)
        self._lb_pipe.AutoSize   = False

        self._grb_pipe.Controls.Add(self._btt_pickPipe)
        self._grb_pipe.Controls.Add(self._lb_pipe)

        # ── GroupBox: Accessories ──────────────────────────────
        self._grb_acc            = GroupBox()
        self._grb_acc.Text       = "Accessories"
        self._grb_acc.Font       = Font("Meiryo UI", 8,
                                        System.Drawing.FontStyle.Bold)
        self._grb_acc.ForeColor  = Color.Red
        self._grb_acc.Location   = Point(p, 98)
        self._grb_acc.Size       = Size(w - p*2, 284)

        self._lb_filterLbl           = Label()
        self._lb_filterLbl.Text      = "Filter:"
        self._lb_filterLbl.Font      = Font("Meiryo UI", 8,
                                            System.Drawing.FontStyle.Bold)
        self._lb_filterLbl.ForeColor = Color.Black
        self._lb_filterLbl.Location  = Point(6, 22)
        self._lb_filterLbl.Size      = Size(40, 20)

        self._txb_filter          = TextBox()
        self._txb_filter.Font     = Font("Meiryo UI", 8,
                                         System.Drawing.FontStyle.Bold)
        self._txb_filter.Location = Point(50, 20)
        self._txb_filter.Size     = Size(w - p*2 - 58, 22)
        self._txb_filter.TextChanged += self.Txb_filterChanged

        self._lb_count           = Label()
        self._lb_count.Text      = "0 items"
        self._lb_count.Font      = Font("Meiryo UI", 7,
                                        System.Drawing.FontStyle.Italic)
        self._lb_count.ForeColor = Color.Gray
        self._lb_count.Location  = Point(6, 46)
        self._lb_count.Size      = Size(w - p*2 - 14, 16)

        self._clb                    = CheckedListBox()
        self._clb.DisplayMember       = 'Name'
        self._clb.Font                = Font("Meiryo UI", 7.5,
                                             System.Drawing.FontStyle.Bold)
        self._clb.ForeColor           = Color.DarkBlue
        self._clb.FormattingEnabled   = True
        self._clb.CheckOnClick         = False
        self._clb.HorizontalScrollbar  = True
        self._clb.Location            = Point(6, 64)
        self._clb.Size                = Size(w - p*2 - 14, 210)
        self._clb.MouseDown           += self.Clb_MouseDown

        self._grb_acc.Controls.Add(self._lb_filterLbl)
        self._grb_acc.Controls.Add(self._txb_filter)
        self._grb_acc.Controls.Add(self._lb_count)
        self._grb_acc.Controls.Add(self._clb)

        # ── GroupBox: Placement Point ──────────────────────────
        self._grb_point           = GroupBox()
        self._grb_point.Text      = "Placement Point"
        self._grb_point.Font      = Font("Meiryo UI", 8,
                                         System.Drawing.FontStyle.Bold)
        self._grb_point.ForeColor = Color.Red
        self._grb_point.Location  = Point(p, 390)
        self._grb_point.Size      = Size(w - p*2, 72)

        self._btt_pickPt           = Button()
        self._btt_pickPt.Text      = "Pick on Pipe"
        self._btt_pickPt.Font      = Font("Meiryo UI", 8,
                                          System.Drawing.FontStyle.Bold)
        self._btt_pickPt.ForeColor = Color.Blue
        self._btt_pickPt.Location  = Point(6, 18)
        self._btt_pickPt.Size      = Size(100, 26)
        self._btt_pickPt.UseVisualStyleBackColor = True
        self._btt_pickPt.Click    += self.Btt_pickPointClick

        self._lb_point_raw        = Label()
        self._lb_point_raw.Text   = "Surface : -"
        self._lb_point_raw.Font   = Font("Meiryo UI", 7.5,
                                         System.Drawing.FontStyle.Regular)
        self._lb_point_raw.ForeColor = Color.DimGray
        self._lb_point_raw.Location  = Point(114, 18)
        self._lb_point_raw.Size      = Size(w - p*2 - 118, 18)

        self._lb_point_proj        = Label()
        self._lb_point_proj.Text   = "Centerline : -"
        self._lb_point_proj.Font   = Font("Meiryo UI", 7.5,
                                          System.Drawing.FontStyle.Bold)
        self._lb_point_proj.ForeColor = Color.DimGray
        self._lb_point_proj.Location  = Point(114, 40)
        self._lb_point_proj.Size      = Size(w - p*2 - 118, 18)

        self._grb_point.Controls.Add(self._btt_pickPt)
        self._grb_point.Controls.Add(self._lb_point_raw)
        self._grb_point.Controls.Add(self._lb_point_proj)

        # ── GroupBox: Log ──────────────────────────────────────
        self._grb_log            = GroupBox()
        self._grb_log.Text       = "Log"
        self._grb_log.Font       = Font("Meiryo UI", 8,
                                        System.Drawing.FontStyle.Bold)
        self._grb_log.ForeColor  = Color.Red
        self._grb_log.Location   = Point(p, 470)
        self._grb_log.Size       = Size(w - p*2, 80)

        self._txb_log            = TextBox()
        self._txb_log.Multiline  = True
        self._txb_log.ReadOnly   = True
        self._txb_log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        self._txb_log.Font       = Font("Consolas", 7)
        self._txb_log.BackColor  = Color.White
        self._txb_log.Location   = Point(6, 18)
        self._txb_log.Size       = Size(w - p*2 - 14, 54)
        self._grb_log.Controls.Add(self._txb_log)

        # ── Status ────────────────────────────────────────────
        self._lb_status           = Label()
        self._lb_status.Text      = "Start: click 'Pick Pipe'."
        self._lb_status.Font      = Font("Meiryo UI", 7.5,
                                         System.Drawing.FontStyle.Bold)
        self._lb_status.ForeColor = Color.Gray
        self._lb_status.Location  = Point(p, 556)
        self._lb_status.Size      = Size(w - p*2, 18)

        # ── Buttons ───────────────────────────────────────────
        bw, bh = self.BTN_W, self.BTN_H
        by     = h - bh - 14

        self._btt_RUN            = Button()
        self._btt_RUN.Text       = "RUN"
        self._btt_RUN.Font       = Font("Meiryo UI", 9,
                                        System.Drawing.FontStyle.Bold)
        self._btt_RUN.ForeColor  = Color.Red
        self._btt_RUN.Location   = Point(w - bw*2 - p - 6, by)
        self._btt_RUN.Size       = Size(bw, bh)
        self._btt_RUN.UseVisualStyleBackColor = True
        self._btt_RUN.Click     += self.Btt_RUNClick

        self._btt_CLOSE           = Button()
        self._btt_CLOSE.Text      = "CANCEL"
        self._btt_CLOSE.Font      = Font("Meiryo UI", 9,
                                         System.Drawing.FontStyle.Bold)
        self._btt_CLOSE.ForeColor = Color.Red
        self._btt_CLOSE.Location  = Point(w - bw - p, by)
        self._btt_CLOSE.Size      = Size(bw, bh)
        self._btt_CLOSE.UseVisualStyleBackColor = True
        self._btt_CLOSE.Click    += lambda s, e: self.Close()

        self._lb_fvc             = Label()
        self._lb_fvc.Text        = "@FVC"
        self._lb_fvc.Font        = Font("Meiryo UI", 4.8,
                                        System.Drawing.FontStyle.Bold)
        self._lb_fvc.ForeColor   = Color.LightGray
        self._lb_fvc.Location    = Point(p, by + 10)
        self._lb_fvc.Size        = Size(48, 16)

        self.Text            = "Pipe Accessory Placer"
        self.ClientSize      = Size(w, h)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.TopMost         = True
        self.Controls.Add(self._lb_title)
        self.Controls.Add(self._grb_pipe)
        self.Controls.Add(self._grb_acc)
        self.Controls.Add(self._grb_point)
        self.Controls.Add(self._grb_log)
        self.Controls.Add(self._lb_status)
        self.Controls.Add(self._btt_RUN)
        self.Controls.Add(self._btt_CLOSE)
        self.Controls.Add(self._lb_fvc)

    # ── Helpers ───────────────────────────────────────────────
    def _setStatus(self, text, color=None):
        self._lb_status.Text      = text
        self._lb_status.ForeColor = color or Color.Gray
        System.Windows.Forms.Application.DoEvents()

    def _addLog(self, text):
        self._logLines.append(text)
        if len(self._logLines) > 100:
            self._logLines = self._logLines[-100:]
        self._txb_log.Text = "\r\n".join(reversed(self._logLines))

    def _refreshList(self):
        ft = self._txb_filter.Text.strip().lower()
        self._clb.BeginUpdate()
        self._clb.Items.Clear()
        self._clb_lastIdx = -1
        for wrapper in self._symbols:
            if ft == "" or ft in wrapper.Name.lower():
                self._clb.Items.Add(wrapper)
        self._clb.EndUpdate()
        matched = sum(1 for i in range(self._clb.Items.Count)
                      if not self._clb.Items[i].Name.startswith("[?]"))
        total   = self._clb.Items.Count
        self._lb_count.Text = "{0} matched DN  +  {1} unresolved  =  {2} total".format(
            matched, total - matched, total)
        self._updateHExtent()

    def _updateHExtent(self):
        max_w = 0
        g = self._clb.CreateGraphics()
        for i in range(self._clb.Items.Count):
            tw = int(g.MeasureString(
                self._clb.Items[i].Name, self._clb.Font).Width) + 20
            if tw > max_w: max_w = tw
        g.Dispose()
        self._clb.HorizontalExtent = max_w

    # ── Shift+Click ───────────────────────────────────────────
    def Clb_MouseDown(self, sender, e):
        idx = self._clb.IndexFromPoint(e.X, e.Y)
        if idx < 0: return
        isShift = (System.Windows.Forms.Control.ModifierKeys ==
                   System.Windows.Forms.Keys.Shift)
        if isShift and self._clb_lastIdx >= 0:
            target = self._clb.GetItemChecked(self._clb_lastIdx)
            for i in range(min(self._clb_lastIdx, idx),
                           max(self._clb_lastIdx, idx) + 1):
                self._clb.SetItemChecked(i, target)
        else:
            self._clb.SetItemChecked(idx, not self._clb.GetItemChecked(idx))
            self._clb_lastIdx = idx

    def Txb_filterChanged(self, sender, e):
        self._refreshList()

    # ── Pick Pipe ─────────────────────────────────────────────
    def Btt_pickPipeClick(self, sender, e):
        self.Hide()
        try:
            ref  = uidoc.Selection.PickObject(
                Autodesk.Revit.UI.Selection.ObjectType.Element,
                PipeFilter(), "Pick a pipe")
            pipe = doc.GetElement(ref.ElementId)
            self._pipe    = pipe
            self._dn_mm   = getPipeDiameter_mm(pipe)
            self._setStatus("Loading accessories...", Color.DarkOrange)
            self.Show()
            System.Windows.Forms.Application.DoEvents()
            self.Hide()

            self._symbols = loadAccessoriesByDN(self._dn_mm)

            pipe_type = ""
            try:
                pipe_type = " | {0}".format(
                    doc.GetElement(pipe.GetTypeId()).Name)
            except: pass

            matched = sum(1 for w in self._symbols
                          if not w.Name.startswith("[?]"))
            self._lb_pipe.Text      = "DN {0} mm{1}".format(
                self._dn_mm, pipe_type)
            self._lb_pipe.ForeColor = Color.DarkGreen
            self._refreshList()
            self._setStatus(
                "DN = {0} mm | {1} matched, {2} unresolved.".format(
                    self._dn_mm, matched,
                    len(self._symbols) - matched),
                Color.DarkGreen)
        except Exception as ex:
            err = str(ex)
            if "cancel" not in err.lower() and "operation" not in err.lower():
                self._setStatus("Error: {0}".format(err), Color.OrangeRed)
            else:
                self._setStatus("Pick cancelled.", Color.Gray)
        finally:
            self.Show()
            self.TopMost = True
            self.BringToFront()

    # ── Pick Point on Pipe Face ───────────────────────────────
    def Btt_pickPointClick(self, sender, e):
        if self._pipe is None:
            MessageBox.Show("Please pick a pipe first.", "No Pipe",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self.Hide()
        try:
            # PickObject PointOnElement: bat buoc click tren be mat element
            # Chi cho phep chon be mat ong (PipeFaceFilter)
            ref = uidoc.Selection.PickObject(
                Autodesk.Revit.UI.Selection.ObjectType.PointOnElement,
                PipeFaceFilter(),
                "Click on the pipe surface")

            # Diem tren be mat ong (raw)
            raw_xyz = ref.GlobalPoint

            # Chieu len duong tam ong
            self._insertPt = projectOntoCenterline(self._pipe, raw_xyz)

            # Hien thi ca 2 diem
            def fmt(pt):
                return "({0:.0f}, {1:.0f}, {2:.0f}) mm".format(
                    pt.X * 304.8, pt.Y * 304.8, pt.Z * 304.8)

            self._lb_point_raw.Text      = "Surface    : {0}".format(fmt(raw_xyz))
            self._lb_point_raw.ForeColor = Color.DimGray
            self._lb_point_proj.Text     = "Centerline : {0}".format(
                fmt(self._insertPt))
            self._lb_point_proj.ForeColor = Color.DarkGreen

            self._setStatus(
                "Point on centerline ready. Select accessory and click RUN.",
                Color.DarkOrange)
        except Exception as ex:
            err = str(ex)
            if "cancel" not in err.lower() and "operation" not in err.lower():
                self._setStatus("Error: {0}".format(err), Color.OrangeRed)
            else:
                self._setStatus("Pick cancelled.", Color.Gray)
        finally:
            self.Show()
            self.TopMost = True
            self.BringToFront()

    # ── RUN ───────────────────────────────────────────────────
    def Btt_RUNClick(self, sender, e):
        if self._pipe is None:
            MessageBox.Show("Please pick a pipe first.", "No Pipe",
                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        if self._insertPt is None:
            MessageBox.Show("Please pick a point on the pipe surface.",
                "No Point", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        checked = list(self._clb.CheckedItems)
        if not checked:
            MessageBox.Show("Please select at least one accessory.",
                "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return

        ok_n = fail_n = 0
        for wrapper in checked:
            inst, err = placeAccessory(
                self._pipe, self._insertPt, wrapper.Symbol)
            if inst:
                ok_n += 1
                self._addLog("[OK]   {0}".format(wrapper.Name))
            else:
                fail_n += 1
                self._addLog("[FAIL] {0} -> {1}".format(wrapper.Name, err))

        # Reset point
        self._insertPt = None
        self._lb_point_raw.Text       = "Surface    : -"
        self._lb_point_raw.ForeColor  = Color.DimGray
        self._lb_point_proj.Text      = "Centerline : -"
        self._lb_point_proj.ForeColor = Color.DimGray

        # Uncheck list
        for i in range(self._clb.Items.Count):
            self._clb.SetItemChecked(i, False)

        if fail_n == 0:
            self._setStatus(
                "Placed {0} ok. Pick next point to continue.".format(ok_n),
                Color.DarkGreen)
        else:
            self._setStatus(
                "Done: {0} ok, {1} failed. Pick next point.".format(
                    ok_n, fail_n), Color.OrangeRed)
#endregion

f = AccessoryPlacerForm()
Application.Run(f)
OUT = "Done"


