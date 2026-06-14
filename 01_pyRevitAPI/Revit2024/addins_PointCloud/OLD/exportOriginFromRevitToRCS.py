# -*- coding: utf-8 -*-
"""
Export Point Cloud Transform to RCS  -  IronPython Script for Revit (Dynamo)
=============================================================================
Purpose:
  - Read the current world-space Transform of each Point Cloud (.rcs) file
    linked into Revit.
  - Write that transform to a sidecar XML file (<name>.rcs_transform.xml)
    next to the original .rcs, so the user can re-apply coordinates in
    Autodesk ReCap or Navisworks.
  - The XML also contains a 4x4 row-major matrix for direct pipeline use.

Workflow:
  1. Script collects all PointCloudInstances linked into the host document.
  2. UI appears; user selects one or more files (with a "Select All" checkbox).
  3. User clicks Export  -> XML file(s) written next to each .rcs.
  4. User clicks Cancel  -> nothing happens.

Notes:
  - Transform is world-space: T_world = T_link * T_pc
  - Output file: <rcs_path>_transform.xml
  - The original .rcs file is never modified -> safe, no data loss.

Usage:
  - Run via Dynamo (Script node) or pyRevit.

Copyright: @V2D
"""

# ─────────────────────────────────────────────────────────────────────────────
#  IMPORTS
# ─────────────────────────────────────────────────────────────────────────────
import clr
import sys
import math

# ── Revit API ──
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    PointCloudInstance,
    PointCloudType,
    RevitLinkInstance,
    XYZ,
    Transform,
    BuiltInParameter,
)
from Autodesk.Revit.UI import TaskDialog

# ── Dynamo / RevitServices ──
clr.AddReference('RevitNodes')
clr.AddReference('RevitServices')

import RevitServices
from RevitServices.Persistence import DocumentManager

# ── System / IO ──
import System
import System.IO
from System.IO import Path, File, StreamWriter
from System.Text import StringBuilder

# ── WPF ──
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window, Thickness, Visibility,
    GridLength, GridUnitType,
    HorizontalAlignment, VerticalAlignment,
    FontWeights, TextWrapping,
    MessageBox, MessageBoxButton, MessageBoxImage,
)
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition,
    StackPanel, GroupBox,
    ComboBox, ComboBoxItem,
    Button, TextBlock, Border,
    ScrollViewer, ScrollBarVisibility,
    CheckBox, Orientation, Label,
    Separator,
)
from System.Windows.Media import SolidColorBrush, Color, FontFamily
from System.Windows.Input import Cursors
from System.Windows.Controls import TextBox as WpfTextBox

# ── WinForms (cho FolderBrowserDialog) ──
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import FolderBrowserDialog as WinFormsFBD, DialogResult

# ─────────────────────────────────────────────────────────────────────────────
#  Lấy Revit context từ DocumentManager (bắt buộc khi chạy trong Dynamo)
# ─────────────────────────────────────────────────────────────────────────────
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app   = uiapp.Application
uidoc = uiapp.ActiveUIDocument


# ═════════════════════════════════════════════════════════════════════════════
#  DATA MODEL
# ═════════════════════════════════════════════════════════════════════════════

class RcsEntry(object):
    """
    Wrapper for a single PointCloudInstance linked into Revit.

    Attributes:
        pc_instance    : The raw PointCloudInstance element
        link_instance  : RevitLinkInstance that contains this PC (None if local)
        link_doc       : Document of the linked file (None if local)
        is_linked      : True when the PC comes from a Revit Link file
        display_name   : Human-readable label (usually the .rcs filename)
        rcs_path       : Full path to the source .rcs file ('' if unresolvable)
        world_transform: Transform in world-space coordinates
    """

    def __init__(self, pc_instance, link_instance=None, link_doc=None):
        self.pc_instance   = pc_instance
        self.link_instance = link_instance
        self.link_doc      = link_doc
        self.is_linked     = (link_instance is not None)

        self.rcs_path        = self._resolve_rcs_path()
        self.pc_type_name    = self._resolve_pc_type_name()  # ten tu PointCloudType.Name
        self.display_name    = self._resolve_display_name()
        self.world_transform = self._compute_world_transform()

    # ── Private helpers ───────────────────────────────────────────────────────

    def _target_doc(self):
        # Trả về document chứa PC: link doc hoặc host doc
        return self.link_doc if self.is_linked else doc

    @staticmethod
    def _is_valid_file_path(s):
        """Kiem tra xem s co phai la duong dan file hop le khong.
        Phai chua path separator HOAC co extension .rcs/.rcp.
        Loai bo cac gia tri garbage nhu 'Autodesk.Revit.DB.FilePath'.
        """
        if not s:
            return False
        if s.startswith('<'):
            return False
        # Phai chua path separator
        if '\\' in s or ('/' in s and not s.startswith('http')):
            return True
        # Hoac co extension .rcs / .rcp
        sl = s.lower()
        if sl.endswith('.rcs') or sl.endswith('.rcp'):
            return True
        return False

    def _resolve_rcs_path(self):
        """
        Lay duong dan file .rcs tu PointCloudType.
        GetPath() trong Revit 2024 tra ve FilePath object - can thu nhieu
        property de lay duoc string thuc su.
        """
        try:
            target  = self._target_doc()
            pc_type = target.GetElement(self.pc_instance.GetTypeId())
            if pc_type is None:
                return u""

            p = pc_type.GetPath()
            if p is None:
                return u""

            # Neu p da la string thuan tuy
            if isinstance(p, (str, unicode)):
                s = p.strip()
                if self._is_valid_file_path(s):
                    return s

            # FilePath object - thu lan luot cac property
            for attr in ('OriginalPath', 'CentralPath', 'UserVisiblePath',
                         'OriginalFilePath', 'Path', 'LocalPath'):
                try:
                    val = getattr(p, attr, None)
                    if val is None:
                        continue
                    if isinstance(val, (str, unicode)):
                        s = val.strip()
                        if self._is_valid_file_path(s):
                            return s
                    else:
                        s = System.Convert.ToString(val).strip()
                        if self._is_valid_file_path(s):
                            return s
                except:
                    continue

            # Thu ToString() tren chinh doi tuong p
            try:
                s = System.Convert.ToString(p).strip()
                if self._is_valid_file_path(s):
                    return s
            except:
                pass

        except:
            pass
        return u""

    def _resolve_pc_type_name(self):
        """Lay Name cua PointCloudType - thuong la ten file khong co extension."""
        try:
            target  = self._target_doc()
            pc_type = target.GetElement(self.pc_instance.GetTypeId())
            if pc_type is not None:
                n = pc_type.Name
                if n and not n.startswith('<'):
                    return u"{}".format(n)
        except:
            pass
        # Fallback: ten instance
        try:
            n = self.pc_instance.Name
            if n and not n.startswith('<'):
                return u"{}".format(n)
        except:
            pass
        return u""

    def _resolve_display_name(self):
        """Uu tien filename tu rcs_path, sau do pc_type_name, cuoi la ElementId."""
        if self.rcs_path:
            try:
                fname = Path.GetFileName(self.rcs_path)
                if fname:
                    return fname
            except:
                pass
            return self.rcs_path

        # Fallback: ten type trong Revit
        if self.pc_type_name:
            return self.pc_type_name

        # Fallback cuoi: dung ElementId
        try:
            return u"PointCloud [ID: {}]".format(self.pc_instance.Id.IntegerValue)
        except:
            return u"PointCloud (unknown)"

    def _compute_world_transform(self):
        # World-space transform:
        #   Nếu PC trong file link: T_world = T_link * T_pc
        #   Nếu PC là local:        T_world = T_pc
        try:
            pc_tf = self.pc_instance.GetTransform()
            if self.is_linked:
                link_tf = self.link_instance.GetTransform()
                return link_tf.Multiply(pc_tf)
            return pc_tf
        except:
            # Trả về Identity nếu không lấy được transform
            return Transform.Identity

    # ── Public properties ─────────────────────────────────────────────────────

    @property
    def element_id(self):
        return self.pc_instance.Id

    @property
    def link_doc_name(self):
        # Tên file link (dùng để hiển thị nguồn gốc của PC)
        if not self.is_linked:
            return u""
        try:
            return self.link_doc.Title or u"Linked File"
        except:
            return u"Linked File"

    @property
    def xml_filename(self):
        """Ten file xml duy nhat cho entry nay (dung lam ten file output)."""
        safe_name = self.pc_type_name if self.pc_type_name else u"PointCloud"
        for ch in u'<>:"/\\|?* ':
            safe_name = safe_name.replace(ch, u'_')
        try:
            elem_id_str = u"_{}".format(self.pc_instance.Id.IntegerValue)
        except:
            elem_id_str = u""
        return safe_name + elem_id_str + u"_transform.xml"

    def get_output_xml_path(self, custom_folder=None):
        """
        Tinh duong dan file XML output.
        Thu tu uu tien:
          0. custom_folder (neu user da chon)
          1. Canh file .rcs goc
          2. Canh file .rvt host doc
          3. Windows TEMP folder
        """
        fname = self.xml_filename
        if custom_folder:
            try:
                return Path.Combine(custom_folder, fname)
            except:
                pass
        # 1. Canh file .rcs
        if self.rcs_path:
            try:
                folder = Path.GetDirectoryName(self.rcs_path)
                if folder:
                    return Path.Combine(folder, fname)
            except:
                pass
        # 2. Canh file .rvt
        try:
            host_path = doc.PathName
            if host_path:
                folder = Path.GetDirectoryName(host_path)
                if folder:
                    return Path.Combine(folder, fname)
        except:
            pass
        # 3. TEMP
        try:
            return Path.Combine(System.IO.Path.GetTempPath(), fname)
        except:
            pass
        return u""

    @property
    def output_xml_path(self):
        """Default output path (khong co custom folder)."""
        return self.get_output_xml_path()


# ─────────────────────────────────────────────────────────────────────────────
#  LINK ENTRY  (RevitLink / CAD Link)
# ─────────────────────────────────────────────────────────────────────────────

class LinkEntry(object):
    """
    Wrapper for a RevitLinkInstance or ImportInstance (CAD link).
    Exposes the same interface as RcsEntry so the UI can treat them uniformly.

    Attributes:
        display_name   : file title shown in UI
        rcs_path       : full path to .rvt / .dwg (may be empty for CAD)
        world_transform: Transform of this link in host coordinates
        link_type      : 'revit' or 'cad'
        is_linked      : False (this IS a top-level link, not inside one)
        is_link_file   : True (marker to distinguish from RcsEntry)
        pc_type_name   : '' (not applicable)
        link_doc_name  : '' (not applicable)
    """

    def __init__(self, name, path, element_id_int, transform, link_type='revit'):
        self.display_name    = name or u"Linked File"
        self.rcs_path        = path or u""
        self._element_id_int = element_id_int
        self.world_transform = transform
        self.link_type       = link_type   # 'revit' | 'cad'
        self.is_linked       = False       # NOT inside another link
        self.is_link_file    = True        # distinguish from RcsEntry
        self.pc_type_name    = u""
        self.link_doc_name   = u""

    @property
    def element_id(self):
        return self._element_id_int

    @property
    def xml_filename(self):
        safe = self.display_name
        for ch in u'<>:"/\\|?* ':
            safe = safe.replace(ch, u'_')
        return u"{}_{}_{}_transform.xml".format(
            safe, self.link_type.upper(), self._element_id_int)

    def get_output_xml_path(self, custom_folder=None):
        fname = self.xml_filename
        if custom_folder:
            try:
                return Path.Combine(custom_folder, fname)
            except:
                pass
        if self.rcs_path:
            try:
                folder = Path.GetDirectoryName(self.rcs_path)
                if folder:
                    return Path.Combine(folder, fname)
            except:
                pass
        try:
            host_path = doc.PathName
            if host_path:
                folder = Path.GetDirectoryName(host_path)
                if folder:
                    return Path.Combine(folder, fname)
        except:
            pass
        try:
            return Path.Combine(System.IO.Path.GetTempPath(), fname)
        except:
            pass
        return u""

    @property
    def output_xml_path(self):
        return self.get_output_xml_path()


# ═════════════════════════════════════════════════════════════════════════════
#  DATA COLLECTION
# ═════════════════════════════════════════════════════════════════════════════

def collect_rcs_entries():
    """
    Collect all PointCloudInstances from the host document:
      - Local  : directly in the host document
      - Linked : inside loaded RevitLinkInstances

    Returns:
        list[RcsEntry]
    """
    entries = []

    # 1. Thu thập PC trong host document (local)
    local_pcs = list(
        FilteredElementCollector(doc).OfClass(PointCloudInstance)
    )
    for pc in local_pcs:
        entries.append(RcsEntry(pc))

    # 2. Thu thập PC trong các file Revit Link đã được load
    link_instances = list(
        FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    )
    for link_inst in link_instances:
        try:
            link_doc = link_inst.GetLinkDocument()
            # Bỏ qua link chưa được load (GetLinkDocument trả về None)
            if link_doc is None:
                continue
            linked_pcs = list(
                FilteredElementCollector(link_doc).OfClass(PointCloudInstance)
            )
            for pc in linked_pcs:
                entries.append(RcsEntry(pc, link_inst, link_doc))
        except:
            pass

    return entries


def collect_link_entries():
    """
    Collect all loaded RevitLinkInstances and ImportInstances (CAD links)
    as LinkEntry objects ready for display in the UI.

    Returns:
        list[LinkEntry]
    """
    result = []

    # -- Revit Links --
    link_instances = list(
        FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    )
    for link_inst in link_instances:
        try:
            link_doc = link_inst.GetLinkDocument()
            if link_doc is None:
                continue
            tf   = link_inst.GetTransform()
            path = u""
            try:
                path = link_doc.PathName or u""
            except:
                pass
            name = link_doc.Title or u"Linked File"
            result.append(LinkEntry(name, path,
                                    link_inst.Id.IntegerValue, tf, 'revit'))
        except:
            pass

    # -- CAD Links (ImportInstance where IsLinked == True) --
    try:
        from Autodesk.Revit.DB import ImportInstance
        import_instances = list(
            FilteredElementCollector(doc).OfClass(ImportInstance)
        )
        for imp in import_instances:
            try:
                if not imp.IsLinked:
                    continue
                tf       = imp.GetTransform()
                cad_name = u""
                try:
                    p = imp.LookupParameter(u"File Name")
                    if p is not None:
                        cad_name = p.AsString() or u""
                except:
                    pass
                if not cad_name:
                    try:
                        cad_name = imp.Category.Name if imp.Category else u"CAD Link"
                    except:
                        cad_name = u"CAD Link [{0}]".format(imp.Id.IntegerValue)
                result.append(LinkEntry(cad_name, u"",
                                        imp.Id.IntegerValue, tf, 'cad'))
            except:
                pass
    except:
        pass

    return result


# ═════════════════════════════════════════════════════════════════════════════
#  TRANSFORM UTILITIES
# ═════════════════════════════════════════════════════════════════════════════

# Conversion factor: Revit internal unit = feet; output = millimeters (1 ft = 304.8 mm)
FT_TO_MM = 304.8


def transform_to_matrix4x4(tf):
    """
    Convert a Revit Transform to a 4x4 row-major matrix (translation in mm).
    BasisX/Y/Z are NORMALIZED to strip any scale factor embedded by Revit
    (linked transforms often carry a ~3.2808 scale == 1/FT_TO_MM * 1000).
    """
    bx = tf.BasisX.Normalize()
    by = tf.BasisY.Normalize()
    bz = tf.BasisZ.Normalize()
    o  = tf.Origin
    ox = o.X * FT_TO_MM
    oy = o.Y * FT_TO_MM
    oz = o.Z * FT_TO_MM
    m = [
        [bx.X, by.X, bz.X, ox],
        [bx.Y, by.Y, bz.Y, oy],
        [bx.Z, by.Z, bz.Z, oz],
        [0.0,  0.0,  0.0,  1.0],
    ]
    return m


def extract_translation_ft(tf):
    """
    Return translation components (X, Y, Z) in feet from a Revit Transform.
    Revit stores translation in feet internally.
    """
    o = tf.Origin
    return o.X, o.Y, o.Z


def extract_rotation_deg(tf):
    """
    Rodrigues axis-angle extraction.
    CRITICAL: normalize BasisX/Y/Z first -- linked transforms carry a
    ~3.2808 scale factor (1/FT_TO_M) that inflates the trace and makes
    cos_a clamp to 1.0, producing angle = 0 for every file.
    Returns: (axis_x, axis_y, axis_z, angle_degrees)
    """
    bx = tf.BasisX.Normalize()
    by = tf.BasisY.Normalize()
    bz = tf.BasisZ.Normalize()

    r00, r10, r20 = bx.X, bx.Y, bx.Z
    r01, r11, r21 = by.X, by.Y, by.Z
    r02, r12, r22 = bz.X, bz.Y, bz.Z

    trace = r00 + r11 + r22
    cos_a = max(-1.0, min(1.0, (trace - 1.0) / 2.0))
    angle     = math.acos(cos_a)
    angle_deg = math.degrees(angle)

    TOL = 1e-9
    if abs(angle) < TOL:
        return 0.0, 0.0, 1.0, 0.0
    sin_a = math.sin(angle)
    if abs(sin_a) < TOL:
        return 0.0, 0.0, 1.0, angle_deg

    ax = (r21 - r12) / (2.0 * sin_a)
    ay = (r02 - r20) / (2.0 * sin_a)
    az = (r10 - r01) / (2.0 * sin_a)

    length = math.sqrt(ax*ax + ay*ay + az*az)
    if length < 1e-12:
        return 0.0, 0.0, 1.0, angle_deg
    return ax/length, ay/length, az/length, angle_deg


def extract_euler_zyx_deg(tf):
    """
    ZYX Tait-Bryan Euler angles in degrees (roll, pitch, yaw).
    Useful as an alternative rotation representation.
    Returns: (rx_deg, ry_deg, rz_deg)
    """
    bx = tf.BasisX.Normalize()
    by = tf.BasisY.Normalize()
    bz = tf.BasisZ.Normalize()
    # R[row][col]: col0=bx, col1=by, col2=bz
    r20 = bx.Z
    ry  = math.asin(max(-1.0, min(1.0, -r20)))
    cy  = math.cos(ry)
    TOL = 1e-6
    if abs(cy) > TOL:
        rx = math.atan2(by.Z, bz.Z)
        rz = math.atan2(bx.Y, bx.X)
    else:
        rz = math.atan2(-by.X, by.Y)
        rx = 0.0
    return math.degrees(rx), math.degrees(ry), math.degrees(rz)


# ═════════════════════════════════════════════════════════════════════════════
#  XML WRITER
# ═════════════════════════════════════════════════════════════════════════════

def _get_combined_output_path(custom_folder=None):
    """
    Tinh duong dan file XML gop chung.
    Thu tu: custom_folder -> thu muc .rvt -> TEMP.
    Ten file: PointCloudTransforms.xml
    """
    fname = u"PointCloudTransforms.xml"
    if custom_folder:
        try:
            return Path.Combine(custom_folder, fname)
        except:
            pass
    try:
        host_path = doc.PathName
        if host_path:
            folder = Path.GetDirectoryName(host_path)
            if folder:
                return Path.Combine(folder, fname)
    except:
        pass
    try:
        return Path.Combine(System.IO.Path.GetTempPath(), fname)
    except:
        pass
    return u""


def _build_entry_xml_block(sb, entry, index):
    """
    Append XML block cho 1 PointCloud entry vao StringBuilder sb.
    """
    tf  = entry.world_transform
    mat = transform_to_matrix4x4(tf)

    tx, ty, tz      = extract_translation_ft(tf)
    ax, ay, az, ang = extract_rotation_deg(tf)

    tx_mm = tx * FT_TO_MM
    ty_mm = ty * FT_TO_MM
    tz_mm = tz * FT_TO_MM

    # Quick-reference comment for copy-paste
    sb.AppendLine(u'')
    sb.AppendLine(u'  <!--')
    sb.AppendLine(u'    [NAVISWORKS] File Units and Transform:')
    sb.AppendLine(u'      Units  = Millimeters')
    sb.AppendLine(u'      Origin X={:.3f}  Y={:.3f}  Z={:.3f}'.format(tx_mm, ty_mm, tz_mm))
    sb.AppendLine(u'      Rotation {:.4f} deg  about axis ({:.6f}, {:.6f}, {:.6f})'.format(ang, ax, ay, az))
    sb.AppendLine(u'    [RECAP] Edit > Registration > apply Translation + 4x4 Matrix below.')
    sb.AppendLine(u'  -->')

    sb.AppendLine(u'  <PointCloud index="{}">'.format(index))

    # Source
    sb.AppendLine(u'    <Source>')
    sb.AppendLine(u'      <FileName>{}</FileName>'.format(entry.display_name))
    sb.AppendLine(u'      <TypeName>{}</TypeName>'.format(entry.pc_type_name))
    sb.AppendLine(u'      <RcsFile>{}</RcsFile>'.format(entry.rcs_path))
    sb.AppendLine(u'      <ElementId>{}</ElementId>'.format(entry.element_id))
    if entry.is_linked:
        sb.AppendLine(u'      <LinkedFrom>{}</LinkedFrom>'.format(entry.link_doc_name))
    sb.AppendLine(u'    </Source>')

    # Translation - feet
    sb.AppendLine(u'')
    sb.AppendLine(u'    <!-- Translation (Revit internal: feet) -->')
    sb.AppendLine(u'    <Translation unit="feet">')
    sb.AppendLine(u'      <X>{:.8f}</X>'.format(tx))
    sb.AppendLine(u'      <Y>{:.8f}</Y>'.format(ty))
    sb.AppendLine(u'      <Z>{:.8f}</Z>'.format(tz))
    sb.AppendLine(u'    </Translation>')

    # Translation - mm
    sb.AppendLine(u'')
    sb.AppendLine(u'    <!-- Translation (mm, for ReCap / Navisworks) -->')
    sb.AppendLine(u'    <Translation unit="mm">')
    sb.AppendLine(u'      <X>{:.4f}</X>'.format(tx_mm))
    sb.AppendLine(u'      <Y>{:.4f}</Y>'.format(ty_mm))
    sb.AppendLine(u'      <Z>{:.4f}</Z>'.format(tz_mm))
    sb.AppendLine(u'    </Translation>')

    # Rotation (axis-angle)
    sb.AppendLine(u'')
    sb.AppendLine(u'    <!-- Rotation: axis-angle -->')
    sb.AppendLine(u'    <!-- NAVISWORKS: File Units and Transform > Rotation field = AngleDegrees, Axis X/Y/Z below -->')
    sb.AppendLine(u'    <Rotation>')
    sb.AppendLine(u'      <Axis X="{:.8f}" Y="{:.8f}" Z="{:.8f}"/>'.format(ax, ay, az))
    sb.AppendLine(u'      <AngleDegrees>{:.6f}</AngleDegrees>'.format(ang))
    sb.AppendLine(u'    </Rotation>')

    # Euler angles ZYX
    rx, ry, rz = extract_euler_zyx_deg(tf)
    sb.AppendLine(u'')
    sb.AppendLine(u'    <!-- Euler ZYX: roll={:.4f} pitch={:.4f} yaw={:.4f} deg (alternative format) -->'.format(rx, ry, rz))
    sb.AppendLine(u'    <EulerAnglesZYX roll="{:.6f}" pitch="{:.6f}" yaw="{:.6f}"/>'.format(rx, ry, rz))

    # 4x4 matrix
    sb.AppendLine(u'')
    sb.AppendLine(u'    <!-- 4x4 row-major matrix (mm) -->')
    sb.AppendLine(u'    <!-- RECAP: Edit > Registration > Transform Matrix - enter 16 values row by row -->')
    sb.AppendLine(u'    <Matrix4x4 unit="mm" row_major="true">')
    for i, row in enumerate(mat):
        sb.AppendLine(u'      <Row index="{}">{:.10f} {:.10f} {:.10f} {:.10f}</Row>'.format(
            i, row[0], row[1], row[2], row[3]))
    sb.AppendLine(u'    </Matrix4x4>')

    # Basis vectors (normalized unit vectors)
    sb.AppendLine(u'')
    sb.AppendLine(u'    <!-- Normalized rotation basis vectors (unit length) -->')
    sb.AppendLine(u'    <BasisVectors>')
    bx_v = tf.BasisX.Normalize()
    by_v = tf.BasisY.Normalize()
    bz_v = tf.BasisZ.Normalize()
    sb.AppendLine(u'      <BasisX X="{:.10f}" Y="{:.10f}" Z="{:.10f}"/>'.format(bx_v.X, bx_v.Y, bx_v.Z))
    sb.AppendLine(u'      <BasisY X="{:.10f}" Y="{:.10f}" Z="{:.10f}"/>'.format(by_v.X, by_v.Y, by_v.Z))
    sb.AppendLine(u'      <BasisZ X="{:.10f}" Y="{:.10f}" Z="{:.10f}"/>'.format(bz_v.X, bz_v.Y, bz_v.Z))
    sb.AppendLine(u'    </BasisVectors>')

    sb.AppendLine(u'  </PointCloud>')


def _build_link_xml_block(sb, link_data, index):
    """
    Write a <RevitLink> block describing the transform of a Revit Link file.
    """
    tf  = link_data[u'transform']
    mat = transform_to_matrix4x4(tf)
    ax, ay, az, ang = extract_rotation_deg(tf)
    rx, ry, rz      = extract_euler_zyx_deg(tf)
    o   = tf.Origin
    tx_mm = o.X * FT_TO_MM
    ty_mm = o.Y * FT_TO_MM
    tz_mm = o.Z * FT_TO_MM

    ltype = link_data.get(u'link_type', u'revit').upper()
    sb.AppendLine(u'  <LinkedFile index="{}" name="{}" type="{}">'.format(index, link_data[u'name'], ltype))
    sb.AppendLine(u'    <ElementId>{}</ElementId>'.format(link_data[u'element_id']))
    sb.AppendLine(u'    <FilePath>{}</FilePath>'.format(link_data[u'path']))
    sb.AppendLine(u'')
    sb.AppendLine(u'    <!-- Translation (mm) -->')
    sb.AppendLine(u'    <Translation unit="mm">')
    sb.AppendLine(u'      <X>{:.4f}</X>'.format(tx_mm))
    sb.AppendLine(u'      <Y>{:.4f}</Y>'.format(ty_mm))
    sb.AppendLine(u'      <Z>{:.4f}</Z>'.format(tz_mm))
    sb.AppendLine(u'    </Translation>')
    sb.AppendLine(u'')
    sb.AppendLine(u'    <!-- Rotation axis-angle (for Navisworks Units and Transform dialog) -->')
    sb.AppendLine(u'    <Rotation>')
    sb.AppendLine(u'      <Axis X="{:.8f}" Y="{:.8f}" Z="{:.8f}"/>'.format(ax, ay, az))
    sb.AppendLine(u'      <AngleDegrees>{:.6f}</AngleDegrees>'.format(ang))
    sb.AppendLine(u'    </Rotation>')
    sb.AppendLine(u'')
    sb.AppendLine(u'    <!-- Euler ZYX (roll/pitch/yaw) in degrees -->')
    sb.AppendLine(u'    <EulerAnglesZYX roll="{:.6f}" pitch="{:.6f}" yaw="{:.6f}"/>'.format(rx, ry, rz))
    sb.AppendLine(u'')
    sb.AppendLine(u'    <!-- 4x4 row-major matrix (mm) -->')
    sb.AppendLine(u'    <Matrix4x4 unit="mm" row_major="true">')
    for i, row in enumerate(mat):
        sb.AppendLine(u'      <Row index="{}">{:.10f} {:.10f} {:.10f} {:.10f}</Row>'.format(
            i, row[0], row[1], row[2], row[3]))
    sb.AppendLine(u'    </Matrix4x4>')
    sb.AppendLine(u'  </LinkedFile>')


def write_all_transforms_xml(entries, custom_folder=None):
    """
    Write all selected entries (RcsEntry and/or LinkEntry) to one XML file.

    Args:
        entries       : list[RcsEntry | LinkEntry]
        custom_folder : output directory override; None = auto

    Returns:
        (True,  output_path)   -- success
        (False, error_message) -- failure
    """
    out_path = _get_combined_output_path(custom_folder)
    if not out_path:
        return False, u"Cannot determine output path."

    # Separate into PointCloud entries and Link entries
    pc_entries   = [e for e in entries if not getattr(e, 'is_link_file', False)]
    lnk_entries  = [e for e in entries if     getattr(e, 'is_link_file', False)]

    sb = StringBuilder()
    sb.AppendLine(u'<?xml version="1.0" encoding="utf-8"?>')
    sb.AppendLine(u'<!-- Generated by Export Point Cloud Transform  @V2D -->')
    sb.AppendLine(u'<!-- World-space coordinates extracted from Revit (internal unit: feet, output: mm) -->')
    sb.AppendLine(u'')
    sb.AppendLine(u'<!--')
    sb.AppendLine(u'  HOW TO USE IN NAVISWORKS (per PointCloud entry):')
    sb.AppendLine(u'    1. Append the .rcs file to your Navisworks scene.')
    sb.AppendLine(u'    2. In Selection Tree, right-click file > File Units and Transform.')
    sb.AppendLine(u'    3. Units = Millimeters.')
    sb.AppendLine(u'    4. Origin X/Y/Z  <-- Translation unit=mm values.')
    sb.AppendLine(u'    5. Rotation angle <-- AngleDegrees value.')
    sb.AppendLine(u'       Axis X/Y/Z     <-- Axis X/Y/Z values.')
    sb.AppendLine(u'    6. Scale = 1 / 1 / 1.  Click OK.')
    sb.AppendLine(u'')
    sb.AppendLine(u'  HOW TO USE IN RECAP (per PointCloud entry):')
    sb.AppendLine(u'    1. Open project in ReCap Pro.')
    sb.AppendLine(u'    2. Edit > Registration > select scan > Transform.')
    sb.AppendLine(u'    3. Enter the 4x4 Matrix4x4 values (row-major, mm).')
    sb.AppendLine(u'       OR apply Translation X/Y/Z (mm) + Rotation separately.')
    sb.AppendLine(u'-->')
    sb.AppendLine(u'')
    sb.AppendLine(u'<PointCloudTransforms pc_count="{}" link_count="{}">'.format(
        len(pc_entries), len(lnk_entries)))

    # ── Point Cloud blocks ──
    if pc_entries:
        sb.AppendLine(u'')
        sb.AppendLine(u'  <!-- === POINT CLOUD TRANSFORMS === -->')
        for idx, entry in enumerate(pc_entries):
            _build_entry_xml_block(sb, entry, idx)
            sb.AppendLine(u'')

    # ── Linked file blocks ──
    if lnk_entries:
        sb.AppendLine(u'')
        sb.AppendLine(u'  <!-- === LINKED FILE TRANSFORMS (Revit / CAD) === -->')
        sb.AppendLine(u'  <!-- Host-space positions of each linked .rvt / .dwg file. -->')
        for idx, lentry in enumerate(lnk_entries):
            # Build a dict as _build_link_xml_block expects
            ldata = {
                u'name':       lentry.display_name,
                u'path':       lentry.rcs_path,
                u'element_id': lentry.element_id,
                u'transform':  lentry.world_transform,
                u'link_type':  lentry.link_type,
            }
            _build_link_xml_block(sb, ldata, idx)
            sb.AppendLine(u'')

    sb.AppendLine(u'</PointCloudTransforms>')

    try:
        sw = StreamWriter(out_path, False, System.Text.Encoding.UTF8)
        sw.Write(sb.ToString())
        sw.Close()
        return True, out_path
    except Exception as ex:
        return False, u"Write error: " + str(ex)


# ═════════════════════════════════════════════════════════════════════════════
#  UI: WPF Dialog
# ═════════════════════════════════════════════════════════════════════════════

class ExportPCTransformDialog(Window):
    """
    WPF dialog that lets the user select Point Cloud instances and export
    their world-space transforms to sidecar XML files.

    Allows the user to select one or more PCs and export their transforms to XML.
    """

    # ── Color palette ─────────────────────────────────────────────────────────
    # Bảng màu chủ đạo của V2D
    _CLR_PRIMARY     = Color.FromRgb(0,   114, 178)   # V2D blue
    _CLR_PRIMARY_DRK = Color.FromRgb(0,    86, 143)   # darker blue (hover/pressed)
    _CLR_ACCENT      = Color.FromRgb(230, 159,   0)   # V2D yellow accent
    _CLR_BG          = Color.FromRgb(245, 246, 248)   # window background
    _CLR_WHITE       = Color.FromRgb(255, 255, 255)
    _CLR_PANEL_BG    = Color.FromRgb(250, 251, 253)
    _CLR_BORDER      = Color.FromRgb(210, 215, 220)
    _CLR_TEXT        = Color.FromRgb(30,   40,  50)
    _CLR_TEXT_GRAY   = Color.FromRgb(120, 130, 140)
    _CLR_WARN_BG     = Color.FromRgb(255, 248, 225)
    _CLR_WARN_BD     = Color.FromRgb(255, 193,   7)
    _CLR_SUCCESS_BG  = Color.FromRgb(232, 245, 233)
    _CLR_SUCCESS_BD  = Color.FromRgb(102, 187, 106)
    _CLR_LOCAL_BG    = Color.FromRgb(232, 244, 255)
    _CLR_LINK_BG     = Color.FromRgb(240, 255, 240)
    _CLR_LOCAL_BADGE = Color.FromRgb(0,   114, 178)
    _CLR_LINK_BADGE  = Color.FromRgb(34,  139,  34)

    def __init__(self, all_entries):
        """
        Args:
            all_entries: list[RcsEntry | LinkEntry] -- combined PC + Link entries
        """
        self._all_entries          = all_entries
        self._check_items          = []
        self.ExportedPaths         = []
        self.Confirmed             = False
        self._custom_output_folder = None  # None = auto

        self._build_window()
        self._populate_list()

    # ─── Window shell ─────────────────────────────────────────────────────────

    def _build_window(self):
        # Thiet lap cua so chinh
        self.Title                 = u"Export Point Cloud Transform  -  @V2D"
        self.Width                 = 600
        self.Height                = 640
        self.MinWidth              = 480
        self.MinHeight             = 500
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background            = SolidColorBrush(self._CLR_BG)
        self.FontFamily            = FontFamily(u"Segoe UI")
        self.FontSize              = 12
        self.ResizeMode            = System.Windows.ResizeMode.CanResizeWithGrip

        root = Grid()
        self.Content = root

        # Layout: 7 rows
        for h in [
            GridLength.Auto,                      # 0: Header bar
            GridLength.Auto,                      # 1: Info bar
            GridLength(1, GridUnitType.Star),     # 2: PC list
            GridLength.Auto,                      # 3: Transform preview
            GridLength.Auto,                      # 4: Warning bar
            GridLength.Auto,                      # 5: Output folder selector (NEW)
            GridLength.Auto,                      # 6: Button bar
        ]:
            rd = RowDefinition()
            rd.Height = h
            root.RowDefinitions.Add(rd)

        self._build_header(root)
        self._build_info_bar(root)
        self._build_list_group(root)
        self._build_preview(root)
        self._build_warning(root)
        self._build_folder_selector(root)
        self._build_button_bar(root)

    # ── Section builders ──────────────────────────────────────────────────────

    def _build_header(self, root):
        # Header màu xanh chính, chứa tiêu đề và badge version
        header = Border()
        header.Background = SolidColorBrush(self._CLR_PRIMARY)
        header.Padding    = Thickness(18, 14, 18, 14)
        Grid.SetRow(header, 0)

        hgrid = Grid()
        cd0 = ColumnDefinition(); cd0.Width = GridLength(1, GridUnitType.Star)
        cd1 = ColumnDefinition(); cd1.Width = GridLength.Auto
        hgrid.ColumnDefinitions.Add(cd0)
        hgrid.ColumnDefinitions.Add(cd1)

        # Tiêu đề + subtitle bên trái
        title_sp = StackPanel()

        t1 = TextBlock()
        t1.Text       = u"[PC]  Export Point Cloud Transform"
        t1.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        t1.FontSize   = 16
        t1.FontWeight = FontWeights.Bold

        t2 = TextBlock()
        t2.Text         = (u"Extract world-space coordinates from Revit "
                           u"and write to XML for ReCap / Navisworks")
        t2.Foreground   = SolidColorBrush(Color.FromRgb(180, 215, 255))
        t2.FontSize     = 10
        t2.Margin       = Thickness(0, 4, 0, 0)
        t2.TextWrapping = TextWrapping.Wrap

        title_sp.Children.Add(t1)
        title_sp.Children.Add(t2)
        Grid.SetColumn(title_sp, 0)
        hgrid.Children.Add(title_sp)

        # Badge version bên phải
        ver_border = Border()
        ver_border.Background        = SolidColorBrush(Color.FromRgb(0, 86, 143))
        ver_border.CornerRadius      = System.Windows.CornerRadius(4)
        ver_border.Padding           = Thickness(8, 4, 8, 4)
        ver_border.VerticalAlignment = VerticalAlignment.Top
        ver_tb = TextBlock()
        ver_tb.Text       = u"v1.0"
        ver_tb.FontSize   = 10
        ver_tb.FontWeight = FontWeights.Bold
        ver_tb.Foreground = SolidColorBrush(Color.FromRgb(180, 215, 255))
        ver_border.Child  = ver_tb
        Grid.SetColumn(ver_border, 1)
        hgrid.Children.Add(ver_border)

        header.Child = hgrid
        root.Children.Add(header)

    def _build_info_bar(self, root):
        # Thanh thông tin nhỏ bên dưới header
        bar = Border()
        bar.Background      = SolidColorBrush(Color.FromRgb(240, 245, 255))
        bar.BorderBrush     = SolidColorBrush(self._CLR_BORDER)
        bar.BorderThickness = Thickness(0, 0, 0, 1)
        bar.Padding         = Thickness(16, 7, 16, 7)
        Grid.SetRow(bar, 1)

        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal

        icon = TextBlock()
        icon.Text              = u"[i]  "
        icon.FontSize          = 13
        icon.Foreground        = SolidColorBrush(self._CLR_PRIMARY)
        icon.VerticalAlignment = VerticalAlignment.Center

        info = TextBlock()
        info.Text              = (u"The XML file will be written next to the original .rcs "
                                  u"(original file is untouched). "
                                  u"Translation units: feet (Revit) and meters (ReCap/Navisworks).")
        info.FontSize          = 10
        info.Foreground        = SolidColorBrush(self._CLR_TEXT_GRAY)
        info.TextWrapping      = TextWrapping.Wrap
        info.VerticalAlignment = VerticalAlignment.Center

        sp.Children.Add(icon)
        sp.Children.Add(info)
        bar.Child = sp
        root.Children.Add(bar)

    def _build_list_group(self, root):
        # GroupBox chứa danh sách PC để user chọn
        gb = GroupBox()
        gb.Margin      = Thickness(12, 8, 12, 4)
        gb.Padding     = Thickness(0)
        gb.BorderBrush = SolidColorBrush(self._CLR_BORDER)
        gb.Background  = SolidColorBrush(self._CLR_WHITE)
        Grid.SetRow(gb, 2)

        # Custom header cho GroupBox: tên + badge đếm số lượng đã chọn
        hdr_sp = StackPanel()
        hdr_sp.Orientation = Orientation.Horizontal

        hdr_tb = TextBlock()
        hdr_tb.Text       = u"  [PC]  Point Cloud Files  "
        hdr_tb.FontWeight = FontWeights.SemiBold
        hdr_tb.Foreground = SolidColorBrush(self._CLR_TEXT)

        # Badge hiển thị số item đang được chọn
        self._count_badge = Border()
        self._count_badge.Background        = SolidColorBrush(self._CLR_PRIMARY)
        self._count_badge.CornerRadius      = System.Windows.CornerRadius(8)
        self._count_badge.Padding           = Thickness(6, 1, 6, 1)
        self._count_badge.VerticalAlignment = VerticalAlignment.Center

        self._count_badge_tb = TextBlock()
        self._count_badge_tb.Text       = u"0"
        self._count_badge_tb.FontSize   = 9
        self._count_badge_tb.FontWeight = FontWeights.Bold
        self._count_badge_tb.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        self._count_badge.Child = self._count_badge_tb

        hdr_sp.Children.Add(hdr_tb)
        hdr_sp.Children.Add(self._count_badge)
        gb.Header = hdr_sp

        # Inner grid: 2 rows — select-all bar + scrollable list
        inner = Grid()
        inner.Margin = Thickness(8, 6, 8, 8)
        for h in [
            GridLength.Auto,                  # 0: "Select All" checkbox bar
            GridLength(1, GridUnitType.Star), # 1: Scrollable PC list
        ]:
            rd = RowDefinition(); rd.Height = h
            inner.RowDefinitions.Add(rd)

        # ── "Select All" checkbox bar ────────────────────────────────────────
        sel_row = Border()
        sel_row.Background      = SolidColorBrush(Color.FromRgb(245, 247, 252))
        sel_row.BorderBrush     = SolidColorBrush(self._CLR_BORDER)
        sel_row.BorderThickness = Thickness(1)
        sel_row.CornerRadius    = System.Windows.CornerRadius(3)
        sel_row.Padding         = Thickness(10, 6, 10, 6)
        sel_row.Margin          = Thickness(0, 0, 0, 6)
        Grid.SetRow(sel_row, 0)

        sel_sp = StackPanel()
        sel_sp.Orientation = Orientation.Horizontal

        self._chk_all = CheckBox()
        self._chk_all.Content           = u"  Select All"
        self._chk_all.FontWeight        = FontWeights.SemiBold
        self._chk_all.Foreground        = SolidColorBrush(self._CLR_PRIMARY)
        self._chk_all.VerticalAlignment = VerticalAlignment.Center
        # Khi check/uncheck → đồng bộ tất cả item bên dưới
        self._chk_all.Checked          += self._on_check_all_changed
        self._chk_all.Unchecked        += self._on_check_all_changed

        sel_sp.Children.Add(self._chk_all)
        sel_row.Child = sel_sp
        inner.Children.Add(sel_row)

        # ── Scrollable PC list ───────────────────────────────────────────────
        scr = ScrollViewer()
        scr.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scr.BorderBrush     = SolidColorBrush(self._CLR_BORDER)
        scr.BorderThickness = Thickness(1)
        scr.Background      = SolidColorBrush(self._CLR_WHITE)
        Grid.SetRow(scr, 1)

        self._list_panel = StackPanel()
        self._list_panel.Margin = Thickness(0)
        scr.Content = self._list_panel
        inner.Children.Add(scr)

        gb.Content = inner
        root.Children.Add(gb)

    def _build_preview(self, root):
        # Panel preview transform — chỉ hiển thị khi đúng 1 item được chọn
        self._preview_border = Border()
        self._preview_border.Background      = SolidColorBrush(self._CLR_SUCCESS_BG)
        self._preview_border.BorderBrush     = SolidColorBrush(self._CLR_SUCCESS_BD)
        self._preview_border.BorderThickness = Thickness(1)
        self._preview_border.CornerRadius    = System.Windows.CornerRadius(4)
        self._preview_border.Margin          = Thickness(12, 0, 12, 4)
        self._preview_border.Padding         = Thickness(12, 8, 12, 8)
        self._preview_border.Visibility      = Visibility.Collapsed  # ẩn mặc định
        Grid.SetRow(self._preview_border, 3)

        self._preview_tb = TextBlock()
        self._preview_tb.FontSize     = 10
        self._preview_tb.FontFamily   = FontFamily(u"Consolas")
        self._preview_tb.Foreground   = SolidColorBrush(Color.FromRgb(30, 100, 30))
        self._preview_tb.TextWrapping = TextWrapping.Wrap
        self._preview_border.Child    = self._preview_tb
        root.Children.Add(self._preview_border)

    def _build_warning(self, root):
        # Thanh cảnh báo màu vàng
        warn = Border()
        warn.Background      = SolidColorBrush(self._CLR_WARN_BG)
        warn.BorderBrush     = SolidColorBrush(self._CLR_WARN_BD)
        warn.BorderThickness = Thickness(1)
        warn.CornerRadius    = System.Windows.CornerRadius(3)
        warn.Margin          = Thickness(12, 0, 12, 4)
        warn.Padding         = Thickness(10, 7, 10, 7)
        Grid.SetRow(warn, 4)

        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal

        wi = TextBlock()
        wi.Text              = u"[!]  "
        wi.FontSize          = 13
        wi.Foreground        = SolidColorBrush(Color.FromRgb(180, 120, 0))
        wi.VerticalAlignment = VerticalAlignment.Top

        wt = TextBlock()
        wt.Text              = (u"If .rcs path is resolvable, XML is written next to the .rcs file. "
                                u"Otherwise it is written next to the .rvt document or to the TEMP folder. "
                                u"Coordinates are world-space (Revit Link transform is included if applicable).")
        wt.FontSize          = 10
        wt.Foreground        = SolidColorBrush(Color.FromRgb(130, 80, 0))
        wt.TextWrapping      = TextWrapping.Wrap
        wt.VerticalAlignment = VerticalAlignment.Center

        sp.Children.Add(wi)
        sp.Children.Add(wt)
        warn.Child = sp
        root.Children.Add(warn)

    def _build_folder_selector(self, root):
        """Row chon thu muc output - row 5."""
        fs_border = Border()
        fs_border.Background      = SolidColorBrush(Color.FromRgb(240, 244, 252))
        fs_border.BorderBrush     = SolidColorBrush(self._CLR_BORDER)
        fs_border.BorderThickness = Thickness(1)
        fs_border.CornerRadius    = System.Windows.CornerRadius(3)
        fs_border.Margin          = Thickness(12, 0, 12, 4)
        fs_border.Padding         = Thickness(10, 7, 10, 7)
        Grid.SetRow(fs_border, 5)

        fs_grid = Grid()
        lbl_col = ColumnDefinition(); lbl_col.Width = GridLength.Auto
        tb_col  = ColumnDefinition(); tb_col.Width  = GridLength(1, GridUnitType.Star)
        br_col  = ColumnDefinition(); br_col.Width  = GridLength.Auto
        rs_col  = ColumnDefinition(); rs_col.Width  = GridLength.Auto
        fs_grid.ColumnDefinitions.Add(lbl_col)
        fs_grid.ColumnDefinitions.Add(tb_col)
        fs_grid.ColumnDefinitions.Add(br_col)
        fs_grid.ColumnDefinitions.Add(rs_col)

        lbl = TextBlock()
        lbl.Text                = u"Output Folder:"
        lbl.FontSize            = 10
        lbl.FontWeight          = FontWeights.SemiBold
        lbl.Foreground          = SolidColorBrush(self._CLR_TEXT)
        lbl.VerticalAlignment   = VerticalAlignment.Center
        lbl.Margin              = Thickness(0, 0, 8, 0)
        Grid.SetColumn(lbl, 0)

        self._folder_tb = WpfTextBox()
        self._folder_tb.Text                   = u"(auto - same folder as .rcs or host .rvt)"
        self._folder_tb.FontSize               = 9
        self._folder_tb.IsReadOnly             = True
        self._folder_tb.Background             = SolidColorBrush(Color.FromRgb(235, 238, 246))
        self._folder_tb.BorderThickness        = Thickness(1)
        self._folder_tb.Padding                = Thickness(4, 2, 4, 2)
        self._folder_tb.VerticalContentAlignment = VerticalAlignment.Center
        Grid.SetColumn(self._folder_tb, 1)

        browse_btn = self._make_button(u"[...]  Browse", is_primary=False)
        browse_btn.FontSize = 10
        browse_btn.Padding  = Thickness(10, 4, 10, 4)
        browse_btn.Margin   = Thickness(6, 0, 0, 0)
        browse_btn.Click   += self._on_browse_folder
        Grid.SetColumn(browse_btn, 2)

        reset_btn = self._make_button(u"[x] Auto", is_primary=False)
        reset_btn.FontSize = 10
        reset_btn.Padding  = Thickness(8, 4, 8, 4)
        reset_btn.Margin   = Thickness(4, 0, 0, 0)
        reset_btn.Click   += self._on_reset_folder
        Grid.SetColumn(reset_btn, 3)

        fs_grid.Children.Add(lbl)
        fs_grid.Children.Add(self._folder_tb)
        fs_grid.Children.Add(browse_btn)
        fs_grid.Children.Add(reset_btn)
        fs_border.Child = fs_grid
        root.Children.Add(fs_border)

    def _build_button_bar(self, root):
        # Thanh dưới cùng: logo @V2D bên trái, nút Cancel + Export bên phải
        bottom = Grid()
        bottom.Margin = Thickness(12, 4, 12, 12)

        cd_logo = ColumnDefinition(); cd_logo.Width = GridLength(1, GridUnitType.Star)
        cd_btns = ColumnDefinition(); cd_btns.Width = GridLength.Auto
        bottom.ColumnDefinitions.Add(cd_logo)
        bottom.ColumnDefinitions.Add(cd_btns)
        Grid.SetRow(bottom, 6)

        # ── Logo @V2D (copyright, góc dưới trái) ────────────────────────────
        logo_border = Border()
        logo_border.VerticalAlignment = VerticalAlignment.Center

        logo_sp = StackPanel()
        logo_sp.Orientation = Orientation.Horizontal

        # Thanh dọc màu vàng accent bên cạnh logo text
        logo_bar = Border()
        logo_bar.Width        = 3
        logo_bar.Background   = SolidColorBrush(self._CLR_ACCENT)
        logo_bar.CornerRadius = System.Windows.CornerRadius(2)
        logo_bar.Margin       = Thickness(0, 0, 8, 0)

        logo_txt_sp = StackPanel()

        logo_main = TextBlock()
        logo_main.Text       = u"@V2D"
        logo_main.FontSize   = 13
        logo_main.FontWeight = FontWeights.Bold
        logo_main.Foreground = SolidColorBrush(self._CLR_PRIMARY)

        logo_sub = TextBlock()
        logo_sub.Text       = u"Point Cloud Tools"
        logo_sub.FontSize   = 9
        logo_sub.Foreground = SolidColorBrush(self._CLR_TEXT_GRAY)

        logo_txt_sp.Children.Add(logo_main)
        logo_txt_sp.Children.Add(logo_sub)
        logo_sp.Children.Add(logo_bar)
        logo_sp.Children.Add(logo_txt_sp)
        logo_border.Child = logo_sp

        Grid.SetColumn(logo_border, 0)
        bottom.Children.Add(logo_border)

        # ── Action buttons ────────────────────────────────────────────────────
        btn_sp = StackPanel()
        btn_sp.Orientation       = Orientation.Horizontal
        btn_sp.VerticalAlignment = VerticalAlignment.Center

        self._cancel_btn = self._make_button(u"[X]  Cancel", is_primary=False)
        self._cancel_btn.Margin = Thickness(0, 0, 10, 0)
        self._cancel_btn.Click += self._on_cancel

        self._export_btn = self._make_button(u"[v]  Export Transform", is_primary=True)
        self._export_btn.IsEnabled = False   # bat sau khi user chon it nhat 1 item
        self._export_btn.Click    += self._on_export

        btn_sp.Children.Add(self._cancel_btn)
        btn_sp.Children.Add(self._export_btn)
        Grid.SetColumn(btn_sp, 1)
        bottom.Children.Add(btn_sp)

        root.Children.Add(bottom)

    # ─── List population ──────────────────────────────────────────────────────

    def _populate_list(self):
        # Xóa list cũ và điền lại từ _all_entries
        self._list_panel.Children.Clear()
        self._check_items = []

        # Group: Local PC -> Linked PC (inside Revit links) -> Link Files
        local_pc  = [e for e in self._all_entries
                     if not getattr(e, 'is_link_file', False) and not e.is_linked]
        linked_pc = [e for e in self._all_entries
                     if not getattr(e, 'is_link_file', False) and e.is_linked]
        link_files = [e for e in self._all_entries
                      if getattr(e, 'is_link_file', False)]

        if local_pc:
            self._add_separator(u"  -- Point Clouds (Local) --")
            for e in local_pc:
                self._add_entry_row(e)

        if linked_pc:
            self._add_separator(u"  -- Point Clouds (inside Linked Files) --")
            for e in linked_pc:
                self._add_entry_row(e)

        if link_files:
            self._add_separator(u"  -- Linked Files (Revit Model / CAD) --")
            for e in link_files:
                self._add_entry_row(e)

        if not local_pc and not linked_pc and not link_files:
            empty = TextBlock()
            empty.Text       = u"  (No Point Cloud or Linked File found)"
            empty.FontSize   = 11
            empty.Foreground = SolidColorBrush(self._CLR_TEXT_GRAY)
            empty.Margin     = Thickness(12, 16, 12, 16)
            self._list_panel.Children.Add(empty)

        self._update_count_badge()

    def _add_separator(self, text):
        # Dòng phân nhóm (không thể chọn)
        sep_b = Border()
        sep_b.Background      = SolidColorBrush(Color.FromRgb(238, 242, 248))
        sep_b.BorderBrush     = SolidColorBrush(self._CLR_BORDER)
        sep_b.BorderThickness = Thickness(0, 0, 0, 1)
        sep_b.Padding         = Thickness(8, 4, 8, 4)

        sep_tb = TextBlock()
        sep_tb.Text       = text
        sep_tb.FontSize   = 9
        sep_tb.FontWeight = FontWeights.SemiBold
        sep_tb.Foreground = SolidColorBrush(self._CLR_TEXT_GRAY)
        sep_b.Child = sep_tb
        self._list_panel.Children.Add(sep_b)

    def _add_entry_row(self, entry):
        # Một hàng trong danh sách: checkbox + thông tin PC
        row_border = Border()
        row_border.Padding         = Thickness(10, 8, 10, 8)
        row_border.BorderBrush     = SolidColorBrush(Color.FromRgb(235, 238, 242))
        row_border.BorderThickness = Thickness(0, 0, 0, 1)
        # Nền xanh lá nhạt cho linked item, trắng cho local
        row_border.Background = SolidColorBrush(
            self._CLR_LINK_BG if entry.is_linked else self._CLR_WHITE)

        row_grid = Grid()
        cd_cb   = ColumnDefinition(); cd_cb.Width   = GridLength.Auto
        cd_info = ColumnDefinition(); cd_info.Width = GridLength(1, GridUnitType.Star)
        row_grid.ColumnDefinitions.Add(cd_cb)
        row_grid.ColumnDefinitions.Add(cd_info)

        # CheckBox bên trái
        cb = CheckBox()
        cb.VerticalAlignment = VerticalAlignment.Center
        cb.Margin            = Thickness(0, 0, 10, 0)
        cb.Checked          += self._on_item_check_changed
        cb.Unchecked        += self._on_item_check_changed
        cb.Tag               = entry
        Grid.SetColumn(cb, 0)
        row_grid.Children.Add(cb)

        # Info panel bên phải
        info_sp = StackPanel()

        # Hang tren: badge + ten file
        name_row = StackPanel()
        name_row.Orientation = Orientation.Horizontal

        # Badge: RVT / CAD / LINK (pc inside link) / LOCAL
        is_link_file = getattr(entry, 'is_link_file', False)
        if is_link_file:
            lt = getattr(entry, 'link_type', 'revit').upper()
            badge_txt = lt                                # 'RVT' or 'CAD'
            badge_clr = (Color.FromRgb(100, 60, 180)     # purple for RVT
                         if lt == 'REVIT'
                         else Color.FromRgb(180, 90, 0))  # orange for CAD
        elif entry.is_linked:
            badge_txt = u"PC-LINK"
            badge_clr = self._CLR_LINK_BADGE
        else:
            badge_txt = u"PC-LOCAL"
            badge_clr = self._CLR_LOCAL_BADGE
        badge = self._make_badge(badge_txt, badge_clr, Color.FromRgb(255, 255, 255))
        badge.Margin = Thickness(0, 0, 6, 0)

        # File name label
        name_tb = TextBlock()
        name_tb.Text              = entry.display_name
        name_tb.FontWeight        = FontWeights.SemiBold
        name_tb.Foreground        = SolidColorBrush(self._CLR_TEXT)
        name_tb.VerticalAlignment = VerticalAlignment.Center
        name_tb.TextWrapping      = TextWrapping.Wrap

        name_row.Children.Add(badge)
        name_row.Children.Add(name_tb)

        # Neu display_name khac pc_type_name thi hien thi them type name
        if entry.pc_type_name and entry.pc_type_name != entry.display_name:
            type_badge = self._make_badge(
                u"Type: " + entry.pc_type_name,
                Color.FromRgb(100, 100, 120),
                Color.FromRgb(255, 255, 255)
            )
            type_badge.Margin = Thickness(6, 0, 0, 0)
            name_row.Children.Add(type_badge)

        # Hàng dưới: đường dẫn + link source + tọa độ preview
        sub_sp = StackPanel()
        sub_sp.Margin = Thickness(0, 3, 0, 0)

        if entry.rcs_path:
            path_tb = TextBlock()
            path_tb.Text         = u"  [path] " + entry.rcs_path
            path_tb.FontSize     = 9
            path_tb.Foreground   = SolidColorBrush(self._CLR_TEXT_GRAY)
            path_tb.TextWrapping = TextWrapping.Wrap
            sub_sp.Children.Add(path_tb)
        elif entry.output_xml_path:
            # rcs_path khong resolve duoc nhung co fallback output path
            info_tb = TextBlock()
            info_tb.Text         = u"  [i] No .rcs path - output: " + entry.output_xml_path
            info_tb.FontSize     = 9
            info_tb.Foreground   = SolidColorBrush(Color.FromRgb(100, 100, 160))
            info_tb.TextWrapping = TextWrapping.Wrap
            sub_sp.Children.Add(info_tb)
        else:
            # Khong co bat ky path nao - entry se bi skip
            no_path = TextBlock()
            no_path.Text       = u"  [!] Path not found - this entry will be skipped"
            no_path.FontSize   = 9
            no_path.Foreground = SolidColorBrush(Color.FromRgb(180, 60, 0))
            sub_sp.Children.Add(no_path)

        if entry.is_linked:
            # Hien thi ten file link nguon
            link_src = TextBlock()
            link_src.Text       = u"  <- From linked file: " + entry.link_doc_name
            link_src.FontSize   = 9
            link_src.Foreground = SolidColorBrush(self._CLR_LINK_BADGE)
            link_src.FontWeight = FontWeights.SemiBold
            sub_sp.Children.Add(link_src)

        # Quick transform preview: offset (mm) + rotation angle
        tf = entry.world_transform
        o  = tf.Origin
        tx_mm = o.X * FT_TO_MM
        ty_mm = o.Y * FT_TO_MM
        tz_mm = o.Z * FT_TO_MM
        _, _, _, ang = extract_rotation_deg(tf)
        coord_tb = TextBlock()
        coord_tb.Text       = u"  /> Offset (mm): X={:.1f}  Y={:.1f}  Z={:.1f}   Rot={:.3f}deg".format(
            tx_mm, ty_mm, tz_mm, ang)
        coord_tb.FontFamily = FontFamily(u"Consolas")
        coord_tb.FontSize   = 9
        coord_tb.Foreground = SolidColorBrush(Color.FromRgb(0, 100, 60))
        sub_sp.Children.Add(coord_tb)

        info_sp.Children.Add(name_row)
        info_sp.Children.Add(sub_sp)
        Grid.SetColumn(info_sp, 1)
        row_grid.Children.Add(info_sp)

        row_border.Child = row_grid
        self._list_panel.Children.Add(row_border)
        self._check_items.append((cb, row_border, entry))

    # ─── UI helpers ──────────────────────────────────────────────────────────

    def _make_button(self, text, is_primary=True):
        btn = Button()
        btn.Content    = text
        btn.FontWeight = FontWeights.SemiBold
        btn.FontSize   = 12
        btn.Cursor     = Cursors.Hand
        btn.Padding    = Thickness(18, 9, 18, 9)
        if is_primary:
            # Nút chính: nền xanh V2D, chữ trắng
            btn.Background      = SolidColorBrush(self._CLR_PRIMARY)
            btn.Foreground      = SolidColorBrush(Color.FromRgb(255, 255, 255))
            btn.BorderThickness = Thickness(0)
        else:
            # Nút phụ: nền xám nhạt, chữ tối
            btn.Background      = SolidColorBrush(Color.FromRgb(220, 224, 230))
            btn.Foreground      = SolidColorBrush(Color.FromRgb(50, 55, 65))
            btn.BorderThickness = Thickness(0)
        return btn

    def _make_badge(self, text, bg, fg):
        # Tạo badge nhỏ có nền màu (LOCAL / LINK)
        b = Border()
        b.Background        = SolidColorBrush(bg)
        b.CornerRadius      = System.Windows.CornerRadius(3)
        b.Padding           = Thickness(4, 1, 4, 1)
        b.VerticalAlignment = VerticalAlignment.Center
        tb = TextBlock()
        tb.Text       = text
        tb.FontSize   = 8
        tb.FontWeight = FontWeights.Bold
        tb.Foreground = SolidColorBrush(fg)
        b.Child = tb
        return b

    # ─── Event handlers ───────────────────────────────────────────────────────

    def _on_check_all_changed(self, sender, e):
        # Đồng bộ trạng thái checked của tất cả item theo "Select All"
        checked = (self._chk_all.IsChecked == True)
        for cb, row_b, entry in self._check_items:
            cb.IsChecked = checked
        self._update_export_button()
        self._update_count_badge()

    def _on_item_check_changed(self, sender, e):
        self._update_export_button()
        self._update_count_badge()

        # Đồng bộ ngược checkbox "Select All" theo trạng thái các item
        n_checked = sum(1 for cb, _, __ in self._check_items if cb.IsChecked == True)
        n_total   = len(self._check_items)
        if n_checked == 0:
            self._chk_all.IsChecked = False
        elif n_checked == n_total:
            self._chk_all.IsChecked = True
        else:
            # Trạng thái indeterminate khi chọn một phần
            self._chk_all.IsChecked = None

        # Hiển thị preview nếu đúng 1 item được chọn
        checked_entries = [e for cb, _, e in self._check_items if cb.IsChecked == True]
        if len(checked_entries) == 1:
            self._show_preview(checked_entries[0])
        else:
            self._preview_border.Visibility = Visibility.Collapsed

    def _show_preview(self, entry):
        # Show transform summary of selected entry
        tf = entry.world_transform
        o  = tf.Origin
        tx_mm = o.X * FT_TO_MM
        ty_mm = o.Y * FT_TO_MM
        tz_mm = o.Z * FT_TO_MM
        ax, ay, az, ang = extract_rotation_deg(tf)

        lines = []
        lines.append(u"[Preview] {}".format(entry.display_name))
        lines.append(u"   Translation (mm): X={:.2f}  Y={:.2f}  Z={:.2f}".format(tx_mm, ty_mm, tz_mm))
        lines.append(u"   Rotation: {:.4f} deg  Axis=({:.4f}, {:.4f}, {:.4f})".format(ang, ax, ay, az))
        lines.append(u"   Output -> {}".format(
            entry.get_output_xml_path(self._custom_output_folder) or u"(path not available)"))

        self._preview_tb.Text = u"\n".join(lines)
        self._preview_border.Visibility = Visibility.Visible

    def _update_export_button(self):
        # Cập nhật trạng thái và label của nút Export
        n = sum(1 for cb, _, __ in self._check_items if cb.IsChecked == True)
        self._export_btn.IsEnabled = (n > 0)
        if n > 0:
            self._export_btn.Content = u"[v]  Export Transform ({} file{})".format(
                n, u"s" if n > 1 else u"")
        else:
            self._export_btn.Content = u"[v]  Export Transform"

    def _update_count_badge(self):
        # Cập nhật badge đếm số item đang được chọn
        n = sum(1 for cb, _, __ in self._check_items if cb.IsChecked == True)
        self._count_badge_tb.Text = str(n)
        self._count_badge.Background = SolidColorBrush(
            self._CLR_PRIMARY if n > 0 else Color.FromRgb(160, 170, 180))

    def _on_export(self, sender, e):
        selected = [entry for cb, _, entry in self._check_items if cb.IsChecked == True]
        if not selected:
            return

        folder = self._custom_output_folder  # None = auto
        ok, result = write_all_transforms_xml(selected, folder)

        self.Confirmed = True
        if ok:
            self.ExportedPaths = [result]
        self.Close()

        msg_lines = []
        if ok:
            msg_lines.append(u"[OK] Exported {} point cloud(s) into 1 file:".format(len(selected)))
            msg_lines.append(u"   " + result)
            msg_lines.append(u"")
            msg_lines.append(u"Included:")
            for entry in selected:
                msg_lines.append(u"   - " + entry.display_name)
        else:
            msg_lines.append(u"[X] Export failed:")
            msg_lines.append(u"   " + result)

        TaskDialog.Show(
            u"Export Point Cloud Transform  -  @V2D",
            u"\n".join(msg_lines)
        )

    def _on_browse_folder(self, sender, e):
        dlg = WinFormsFBD()
        dlg.Description      = u"Select output folder for XML files"
        dlg.ShowNewFolderButton = True
        # Mo o vi tri hien tai neu co
        if self._custom_output_folder:
            dlg.SelectedPath = self._custom_output_folder
        result = dlg.ShowDialog()
        if result == DialogResult.OK:
            self._custom_output_folder = dlg.SelectedPath
            self._folder_tb.Text       = self._custom_output_folder

    def _on_reset_folder(self, sender, e):
        self._custom_output_folder = None
        self._folder_tb.Text       = u"(auto - same folder as .rcs or host .rvt)"

    def _on_cancel(self, sender, e):
        self.Confirmed = False
        self.Close()


# ═════════════════════════════════════════════════════════════════════════════
#  MAIN ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

def main():
    # Step 1: Collect PointCloud entries (local + inside Revit links)
    rcs_entries  = collect_rcs_entries()

    # Step 2: Collect Revit Link and CAD Link entries
    link_entries = collect_link_entries()

    all_entries = rcs_entries + link_entries

    if not all_entries:
        TaskDialog.Show(
            u"Export Point Cloud Transform  -  @V2D",
            u"No Point Cloud or Linked File found in the current project."
        )
        return

    # Step 3: Show dialog for user to select and export
    dialog = ExportPCTransformDialog(all_entries)
    dialog.ShowDialog()

    # Bước 3: Kiểm tra kết quả — nếu user nhấn Cancel thì Confirmed = False
    if not dialog.Confirmed:
        return

    if not dialog.ExportedPaths:
        TaskDialog.Show(
            u"Export Point Cloud Transform  -  @V2D",
            u"No files were exported. All selected entries may have missing paths."
        )


# ── Script entry point (Dynamo gọi trực tiếp hoặc qua __main__) ──────────────
if __name__ == "__main__":
    main()
else:
    main()