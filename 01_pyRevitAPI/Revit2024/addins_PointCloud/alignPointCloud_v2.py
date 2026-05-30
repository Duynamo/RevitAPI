# -*- coding: utf-8 -*-
"""
Point Cloud Aligner v2 - IronPython Script for Revit (Dynamo)
=============================================================
Changes from v1:
  - SOURCE Point Cloud can now be selected from a LINKED Revit file
  - The world-space transform of a linked PC = LinkInstance.GetTransform() * PC.GetTransform()
  - CHILD Point Clouds are still from the current document only
    (only elements in the host doc can be moved)

Usage:
  - Run via Dynamo (Script node)
"""

#region ── Import ──────────────────────────────────────────────────────────────
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
    Transaction,
    ElementTransformUtils,
    Line,
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
from RevitServices.Transactions import TransactionManager

# ── System ──
import System
from System.Collections.Generic import List
import System.IO

# ── WPF ──
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Windows import (
    Window, Thickness, Visibility,
    GridLength, GridUnitType,
    HorizontalAlignment, VerticalAlignment,
    FontWeights, TextWrapping,
)
from System.Windows.Controls import (
    Grid, RowDefinition, ColumnDefinition,
    StackPanel, GroupBox,
    ListBox, ListBoxItem, SelectionMode,
    Button, TextBlock, Border,
    ScrollViewer, ScrollBarVisibility,
    CheckBox, Orientation,
)
from System.Windows.Media import SolidColorBrush, Color, FontFamily
from System.Windows.Input import Cursors
#endregion

# ─────────────────────────────────────────────────────────────────────────────
# Lấy context Revit từ Dynamo DocumentManager
# ─────────────────────────────────────────────────────────────────────────────
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app   = uiapp.Application
uidoc = uiapp.ActiveUIDocument


# ═════════════════════════════════════════════════════════════════════════════
# PHẦN 0 - DATA MODEL: Wrapper cho Point Cloud (local hoặc linked)
# ═════════════════════════════════════════════════════════════════════════════

class PointCloudEntry(object):
    """
    Wrapper thống nhất cho cả local PC và linked PC.

    Attributes:
        instance        - PointCloudInstance gốc
        is_linked       - True nếu PC đến từ file link
        link_instance   - RevitLinkInstance nếu is_linked, else None
        link_doc        - Document của file link nếu is_linked, else None
        display_name    - Tên hiển thị trên UI
        world_transform - Transform trong không gian thực tế (world-space)
    """

    def __init__(self, pc_instance, link_instance=None, link_doc=None):
        self.instance      = pc_instance
        self.link_instance = link_instance
        self.link_doc      = link_doc
        self.is_linked     = (link_instance is not None)

        # ── Tên hiển thị ──────────────────────────────────────────────────
        self.display_name = self._resolve_name()

        # ── World-space Transform ─────────────────────────────────────────
        # Nếu PC trong file link:  T_world = T_link * T_pc
        # Nếu PC trong host doc:   T_world = T_pc
        if self.is_linked:
            link_tf = link_instance.GetTransform()
            pc_tf   = pc_instance.GetTransform()
            self.world_transform = link_tf.Multiply(pc_tf)
        else:
            self.world_transform = pc_instance.GetTransform()

    def _resolve_name(self):
        """Lấy tên file Point Cloud."""
        try:
            target_doc = self.link_doc if self.is_linked else doc

            # Ưu tiên: Tên instance
            name = self.instance.Name
            if name and name != u"Point Cloud":
                if name.lower().endswith(('.rcs', '.rcp')):
                    return System.IO.Path.GetFileNameWithoutExtension(name)
                return name

            # Fallback: Path từ PointCloudType
            pc_type = target_doc.GetElement(self.instance.GetTypeId())
            if pc_type:
                path = pc_type.GetPath()
                if path:
                    return System.IO.Path.GetFileNameWithoutExtension(path)
                return pc_type.Name
        except:
            pass
        return u"PointCloud [{}]".format(self.instance.Id)

    @property
    def element_id(self):
        return self.instance.Id

    @property
    def link_doc_name(self):
        """Tên file link (dùng để hiển thị nguồn gốc)."""
        if not self.is_linked:
            return u""
        try:
            title = self.link_doc.Title
            return title if title else u"Linked File"
        except:
            return u"Linked File"


# ═════════════════════════════════════════════════════════════════════════════
# PHẦN 1 - HELPERS: Thu thập dữ liệu + Tính toán Transform
# ═════════════════════════════════════════════════════════════════════════════

ANGLE_TOL     = 1e-9
TRANSLATE_TOL = 1e-6


def collect_all_point_cloud_entries():
    """
    Thu thập tất cả PointCloudEntry:
      - Local PC: từ host document
      - Linked PC: từ tất cả RevitLinkInstance trong host document
    
    Returns:
        local_entries  - list[PointCloudEntry]  (chỉ local, dùng cho target)
        source_entries - list[PointCloudEntry]  (local + linked, dùng cho source)
    """
    local_entries  = []
    source_entries = []

    # ── 1. Local Point Clouds ──────────────────────────────────────────────
    local_pcs = list(
        FilteredElementCollector(doc)
        .OfClass(PointCloudInstance)
    )
    for pc in local_pcs:
        entry = PointCloudEntry(pc)
        local_entries.append(entry)
        source_entries.append(entry)

    # ── 2. Linked Point Clouds ─────────────────────────────────────────────
    link_instances = list(
        FilteredElementCollector(doc)
        .OfClass(RevitLinkInstance)
    )
    for link_inst in link_instances:
        try:
            link_doc = link_inst.GetLinkDocument()
            if link_doc is None:
                continue  # File link chưa load

            linked_pcs = list(
                FilteredElementCollector(link_doc)
                .OfClass(PointCloudInstance)
            )
            for pc in linked_pcs:
                entry = PointCloudEntry(pc, link_instance=link_inst, link_doc=link_doc)
                source_entries.append(entry)
        except:
            pass

    return local_entries, source_entries


def extract_axis_angle(transform):
    """
    Extract rotation axis and angle from a Transform.
    Uses Rodrigues' rotation formula.
    Returns: (XYZ axis, float angle_radians)
    """
    bx = transform.BasisX
    by = transform.BasisY
    bz = transform.BasisZ

    r00, r10, r20 = bx.X, bx.Y, bx.Z
    r01, r11, r21 = by.X, by.Y, by.Z
    r02, r12, r22 = bz.X, bz.Y, bz.Z

    trace = r00 + r11 + r22
    cos_a = max(-1.0, min(1.0, (trace - 1.0) / 2.0))
    angle = math.acos(cos_a)

    if abs(angle) < ANGLE_TOL:
        return XYZ.BasisZ, 0.0

    if abs(angle - math.pi) < ANGLE_TOL:
        # Special case: 180° rotation
        xx = (r00 + 1) / 2.0
        yy = (r11 + 1) / 2.0
        zz = (r22 + 1) / 2.0
        if xx >= yy and xx >= zz:
            x = math.sqrt(max(0, xx))
            y = r01 / (2.0 * x) if x > 1e-12 else 0
            z = r02 / (2.0 * x) if x > 1e-12 else 0
        elif yy >= xx and yy >= zz:
            y = math.sqrt(max(0, yy))
            x = r01 / (2.0 * y) if y > 1e-12 else 0
            z = r12 / (2.0 * y) if y > 1e-12 else 0
        else:
            z = math.sqrt(max(0, zz))
            x = r02 / (2.0 * z) if z > 1e-12 else 0
            y = r12 / (2.0 * z) if z > 1e-12 else 0
        length = math.sqrt(x*x + y*y + z*z)
        if length < 1e-12:
            return XYZ.BasisZ, math.pi
        return XYZ(x/length, y/length, z/length), angle

    sin_a = math.sin(angle)
    ax = (r21 - r12) / (2.0 * sin_a)
    ay = (r02 - r20) / (2.0 * sin_a)
    az = (r10 - r01) / (2.0 * sin_a)
    length = math.sqrt(ax*ax + ay*ay + az*az)
    if length < 1e-12:
        return XYZ.BasisZ, 0.0

    axis = XYZ(ax/length, ay/length, az/length)

    # Ensure axis points in positive direction
    if (axis.Z < 0 or
        (abs(axis.Z) < ANGLE_TOL and axis.Y < 0) or
        (abs(axis.Z) < ANGLE_TOL and abs(axis.Y) < ANGLE_TOL and axis.X < 0)):
        axis  = axis.Negate()
        angle = -angle

    return axis, angle


def describe_transform(t):
    """Mô tả transform dạng chuỗi đọc được."""
    origin = t.Origin
    axis, angle = extract_axis_angle(t)
    angle_deg = math.degrees(angle)

    lines = []
    lines.append(u"  Translation  X: {:.4f}  Y: {:.4f}  Z: {:.4f} (ft)".format(
        origin.X, origin.Y, origin.Z))
    if abs(angle) < ANGLE_TOL:
        lines.append(u"  Rotation     No rotation")
    else:
        lines.append(u"  Rotation     {:.3f}\u00b0  |  Axis: ({:.3f}, {:.3f}, {:.3f})".format(
            angle_deg, axis.X, axis.Y, axis.Z))
    return u"\n".join(lines)


def apply_transform_to_target(target_entry, source_world_transform):
    """
    Áp dụng world-space transform của source lên target local PC.
    delta = source_world_transform × target_world_transform⁻¹
    Thực hiện theo thứ tự: xoay → tịnh tiến

    Args:
        target_entry         - PointCloudEntry (local only)
        source_world_transform - Transform (world-space của source)
    """
    target_instance = target_entry.instance
    tgt_tf = target_instance.GetTransform()

    # delta transform (world-space)
    delta = source_world_transform.Multiply(tgt_tf.Inverse)

    axis, angle = extract_axis_angle(delta)

    # 1. Xoay
    if abs(angle) > ANGLE_TOL:
        tgt_origin    = tgt_tf.Origin
        rot_axis_line = Line.CreateUnbound(tgt_origin, axis)
        ElementTransformUtils.RotateElement(doc, target_instance.Id, rot_axis_line, angle)
        # Refresh transform sau khi xoay
        tgt_tf = target_instance.GetTransform()

    # 2. Tịnh tiến
    translation = source_world_transform.Origin - tgt_tf.Origin
    if translation.GetLength() > TRANSLATE_TOL:
        ElementTransformUtils.MoveElement(doc, target_instance.Id, translation)


# ═════════════════════════════════════════════════════════════════════════════
# PHẦN 2 - UI: WPF Dialog
# ═════════════════════════════════════════════════════════════════════════════

class AlignPointCloudDialog(Window):

    # ── Màu sắc ──────────────────────────────────────────────────────────────
    _CLR_PRIMARY     = Color.FromRgb(0,   102, 204)   # #0066CC
    _CLR_PRIMARY_DRK = Color.FromRgb(0,   82,  163)   # #0052A3
    _CLR_BG          = Color.FromRgb(245, 245, 245)   # #F5F5F5
    _CLR_WHITE       = Color.FromRgb(255, 255, 255)
    _CLR_WARN_BG     = Color.FromRgb(255, 248, 230)   # #FFF8E6
    _CLR_WARN_BD     = Color.FromRgb(255, 184,  77)   # #FFB84D
    _CLR_INFO_BG     = Color.FromRgb(240, 247, 255)   # #F0F7FF
    _CLR_INFO_BD     = Color.FromRgb(179, 209, 255)   # #B3D1FF
    _CLR_LINK_BG     = Color.FromRgb(240, 255, 240)   # #F0FFF0  (xanh lá nhạt cho linked)
    _CLR_LINK_BD     = Color.FromRgb(144, 238, 144)   # #90EE90
    _CLR_LINK_BADGE  = Color.FromRgb( 34, 139,  34)   # #228B22
    _CLR_GRAY        = Color.FromRgb(200, 200, 200)
    _CLR_TEXT_GRAY   = Color.FromRgb(120, 120, 120)

    def __init__(self, local_entries, source_entries):
        """
        Args:
            local_entries   - list[PointCloudEntry]  PC trong host doc (dùng cho target)
            source_entries  - list[PointCloudEntry]  PC local + linked (dùng cho source)
        """
        self._local_entries  = local_entries    # dùng cho danh sách target
        self._source_entries = source_entries   # dùng cho danh sách source

        self._source_entry = None               # PointCloudEntry được chọn làm source
        self._target_items = []                 # list of (CheckBox, PointCloudEntry)

        self.SelectedSource  = None
        self.SelectedTargets = []
        self.Confirmed       = False

        self._build_window()
        self._populate_source_list()

    # ─── Window shell ────────────────────────────────────────────────────────

    def _build_window(self):
        self.Title                 = u"Point Cloud Aligner v2"
        self.Width                 = 530
        self.Height                = 640
        self.MinWidth              = 440
        self.MinHeight             = 500
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background            = SolidColorBrush(self._CLR_BG)
        self.FontFamily            = FontFamily(u"Segoe UI")
        self.FontSize              = 12

        root = Grid()
        self.Content = root

        # Row definitions
        for h in [
            GridLength.Auto,                      # 0 Header
            GridLength.Auto,                      # 1 Legend
            GridLength(200, GridUnitType.Pixel),  # 2 Group: source  (fixed 200px)
            GridLength(1,   GridUnitType.Star),   # 3 Group: targets (fill)
            GridLength.Auto,                      # 4 Warning
            GridLength.Auto,                      # 5 Buttons
        ]:
            rd = RowDefinition()
            rd.Height = h
            root.RowDefinitions.Add(rd)

        # ── Header ──────────────────────────────────────────────────────────
        header = Border()
        header.Background = SolidColorBrush(self._CLR_PRIMARY)
        header.Padding    = Thickness(15, 12, 15, 12)
        Grid.SetRow(header, 0)

        hsp = StackPanel()
        t1  = TextBlock()
        t1.Text       = u"\u2601 Point Cloud Aligner  \u00b7  v2"
        t1.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        t1.FontSize   = 16
        t1.FontWeight = FontWeights.Bold

        t2  = TextBlock()
        t2.Text       = u"Source can be from a Linked Revit file \u2014 Child files must be in the current document"
        t2.Foreground = SolidColorBrush(Color.FromRgb(204, 224, 255))
        t2.FontSize   = 11
        t2.Margin     = Thickness(0, 3, 0, 0)
        t2.TextWrapping = TextWrapping.Wrap

        hsp.Children.Add(t1)
        hsp.Children.Add(t2)
        header.Child = hsp
        root.Children.Add(header)

        # ── Legend (chú giải badge) ──────────────────────────────────────────
        legend_border = Border()
        legend_border.Background     = SolidColorBrush(Color.FromRgb(250, 250, 250))
        legend_border.BorderBrush    = SolidColorBrush(Color.FromRgb(220, 220, 220))
        legend_border.BorderThickness = Thickness(0, 0, 0, 1)
        legend_border.Padding        = Thickness(12, 5, 12, 5)
        Grid.SetRow(legend_border, 1)

        legend_sp = StackPanel()
        legend_sp.Orientation = Orientation.Horizontal

        l1 = self._make_badge(u" LOCAL ", self._CLR_PRIMARY, Color.FromRgb(255, 255, 255))
        l1.Margin = Thickness(0, 0, 6, 0)
        l2 = TextBlock()
        l2.Text = u"= Point Cloud in current document"
        l2.FontSize = 10
        l2.VerticalAlignment = VerticalAlignment.Center
        l2.Foreground = SolidColorBrush(Color.FromRgb(80, 80, 80))
        l2.Margin = Thickness(0, 0, 16, 0)

        l3 = self._make_badge(u" LINK ", self._CLR_LINK_BADGE, Color.FromRgb(255, 255, 255))
        l3.Margin = Thickness(0, 0, 6, 0)
        l4 = TextBlock()
        l4.Text = u"= Point Cloud from Linked Revit file (read-only reference)"
        l4.FontSize = 10
        l4.VerticalAlignment = VerticalAlignment.Center
        l4.Foreground = SolidColorBrush(Color.FromRgb(80, 80, 80))

        legend_sp.Children.Add(l1)
        legend_sp.Children.Add(l2)
        legend_sp.Children.Add(l3)
        legend_sp.Children.Add(l4)
        legend_border.Child = legend_sp
        root.Children.Add(legend_border)

        # ── Group: Source ────────────────────────────────────────────────────
        src_group = self._make_groupbox(u"\u2460  Select SOURCE Point Cloud  (can be from Linked file)", row=2)
        src_outer = Grid()
        src_outer.Margin = Thickness(4)

        for h in [GridLength.Auto, GridLength(1, GridUnitType.Star), GridLength.Auto]:
            rd2 = RowDefinition(); rd2.Height = h
            src_outer.RowDefinitions.Add(rd2)

        src_hint = TextBlock()
        src_hint.Text       = u"Select the reference file (already moved / rotated). Linked files are shown with [LINK] badge:"
        src_hint.FontSize   = 11
        src_hint.Foreground = SolidColorBrush(Color.FromRgb(80, 80, 80))
        src_hint.Margin     = Thickness(0, 0, 0, 6)
        src_hint.TextWrapping = TextWrapping.Wrap
        Grid.SetRow(src_hint, 0)
        src_outer.Children.Add(src_hint)

        # ListBox nguồn (single select)
        src_scr = ScrollViewer()
        src_scr.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        src_scr.BorderBrush     = SolidColorBrush(self._CLR_GRAY)
        src_scr.BorderThickness = Thickness(1)
        src_scr.Background      = SolidColorBrush(self._CLR_WHITE)
        Grid.SetRow(src_scr, 1)

        self._src_listbox = ListBox()
        self._src_listbox.SelectionMode   = SelectionMode.Single
        self._src_listbox.BorderThickness = Thickness(0)
        self._src_listbox.SelectionChanged += self._on_source_changed
        src_scr.Content = self._src_listbox
        src_outer.Children.Add(src_scr)

        # Info transform
        self._info_border = Border()
        self._info_border.Background      = SolidColorBrush(self._CLR_INFO_BG)
        self._info_border.BorderBrush     = SolidColorBrush(self._CLR_INFO_BD)
        self._info_border.BorderThickness = Thickness(1)
        self._info_border.CornerRadius    = System.Windows.CornerRadius(3)
        self._info_border.Padding         = Thickness(8, 6, 8, 6)
        self._info_border.Margin          = Thickness(0, 6, 0, 0)
        self._info_border.Visibility      = Visibility.Collapsed
        Grid.SetRow(self._info_border, 2)

        info_sp    = StackPanel()
        info_title = TextBlock()
        info_title.Text       = u"World-space transform of source:"
        info_title.FontWeight = FontWeights.SemiBold
        info_title.FontSize   = 11
        info_title.Foreground = SolidColorBrush(self._CLR_PRIMARY_DRK)

        self._info_text = TextBlock()
        self._info_text.FontSize    = 11
        self._info_text.Foreground  = SolidColorBrush(Color.FromRgb(50, 50, 50))
        self._info_text.Margin      = Thickness(0, 3, 0, 0)
        self._info_text.FontFamily  = FontFamily(u"Consolas")

        info_sp.Children.Add(info_title)
        info_sp.Children.Add(self._info_text)
        self._info_border.Child = info_sp
        src_outer.Children.Add(self._info_border)

        src_group.Content = src_outer
        root.Children.Add(src_group)

        # ── Group: Targets ───────────────────────────────────────────────────
        tgt_group = self._make_groupbox(u"\u2461  Select CHILD Point Clouds to move  (current document only)", row=3)
        tgt_inner = Grid()
        tgt_inner.Margin = Thickness(4)

        for h in [GridLength.Auto, GridLength(1, GridUnitType.Star), GridLength.Auto]:
            rd3 = RowDefinition(); rd3.Height = h
            tgt_inner.RowDefinitions.Add(rd3)

        # Hint + count
        hint_row = StackPanel()
        hint_row.Orientation = Orientation.Horizontal
        hint_row.Margin      = Thickness(0, 0, 0, 6)
        Grid.SetRow(hint_row, 0)

        tgt_hint = TextBlock()
        tgt_hint.Text               = u"Select one or more local Point Clouds to move:"
        tgt_hint.FontSize           = 11
        tgt_hint.Foreground         = SolidColorBrush(Color.FromRgb(80, 80, 80))
        tgt_hint.VerticalAlignment  = VerticalAlignment.Center

        self._count_text = TextBlock()
        self._count_text.FontSize           = 11
        self._count_text.FontWeight         = FontWeights.SemiBold
        self._count_text.Foreground         = SolidColorBrush(self._CLR_PRIMARY)
        self._count_text.Margin             = Thickness(8, 0, 0, 0)
        self._count_text.VerticalAlignment  = VerticalAlignment.Center

        hint_row.Children.Add(tgt_hint)
        hint_row.Children.Add(self._count_text)
        tgt_inner.Children.Add(hint_row)

        # ScrollViewer chứa các CheckBox item
        scr = ScrollViewer()
        scr.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scr.BorderBrush     = SolidColorBrush(self._CLR_GRAY)
        scr.BorderThickness = Thickness(1)
        scr.Background      = SolidColorBrush(self._CLR_WHITE)
        Grid.SetRow(scr, 1)

        self._tgt_panel = StackPanel()
        self._tgt_panel.Margin = Thickness(4)
        scr.Content = self._tgt_panel
        tgt_inner.Children.Add(scr)

        # Select All / Deselect All
        btn_row = StackPanel()
        btn_row.Orientation         = Orientation.Horizontal
        btn_row.HorizontalAlignment = HorizontalAlignment.Right
        btn_row.Margin              = Thickness(0, 6, 0, 0)
        Grid.SetRow(btn_row, 2)

        btn_all  = self._make_button(u"Select All",   is_primary=False, small=True)
        btn_none = self._make_button(u"Deselect All", is_primary=False, small=True)
        btn_all.Margin  = Thickness(0, 0, 8, 0)
        btn_all.Click  += self._on_select_all
        btn_none.Click += self._on_deselect_all

        btn_row.Children.Add(btn_all)
        btn_row.Children.Add(btn_none)
        tgt_inner.Children.Add(btn_row)

        tgt_group.Content = tgt_inner
        root.Children.Add(tgt_group)

        # ── Warning ──────────────────────────────────────────────────────────
        warn = Border()
        warn.Background      = SolidColorBrush(self._CLR_WARN_BG)
        warn.BorderBrush     = SolidColorBrush(self._CLR_WARN_BD)
        warn.BorderThickness = Thickness(1)
        warn.CornerRadius    = System.Windows.CornerRadius(3)
        warn.Margin          = Thickness(10, 4, 10, 4)
        warn.Padding         = Thickness(10, 7, 10, 7)
        Grid.SetRow(warn, 4)

        warn_sp = StackPanel()
        warn_sp.Orientation = Orientation.Horizontal

        w_icon = TextBlock()
        w_icon.Text              = u"\u26a0"
        w_icon.FontSize          = 14
        w_icon.Foreground        = SolidColorBrush(Color.FromRgb(204, 119, 0))
        w_icon.VerticalAlignment = VerticalAlignment.Top

        w_text = TextBlock()
        w_text.Text         = (u"  This action cannot be undone after saving. "
                                u"Child PCs will be moved to match the world-space "
                                u"position of the source.")
        w_text.FontSize     = 11
        w_text.Foreground   = SolidColorBrush(Color.FromRgb(128, 85, 0))
        w_text.TextWrapping = TextWrapping.Wrap

        warn_sp.Children.Add(w_icon)
        warn_sp.Children.Add(w_text)
        warn.Child = warn_sp
        root.Children.Add(warn)

        # ── Buttons ──────────────────────────────────────────────────────────
        btn_bar = StackPanel()
        btn_bar.Orientation         = Orientation.Horizontal
        btn_bar.HorizontalAlignment = HorizontalAlignment.Right
        btn_bar.Margin              = Thickness(10, 4, 10, 12)
        Grid.SetRow(btn_bar, 5)

        self._cancel_btn = self._make_button(u"Cancel", is_primary=False)
        self._cancel_btn.Margin = Thickness(0, 0, 10, 0)
        self._cancel_btn.Click += self._on_cancel

        self._ok_btn = self._make_button(u"\u2713 Apply", is_primary=True)
        self._ok_btn.IsEnabled = False
        self._ok_btn.Click    += self._on_ok

        btn_bar.Children.Add(self._cancel_btn)
        btn_bar.Children.Add(self._ok_btn)
        root.Children.Add(btn_bar)

    # ─── UI Helpers ──────────────────────────────────────────────────────────

    def _make_groupbox(self, header_text, row):
        gb = GroupBox()
        gb.Header         = header_text
        gb.Margin         = Thickness(10, 6, 10, 4)
        gb.Padding        = Thickness(8)
        gb.BorderBrush    = SolidColorBrush(self._CLR_GRAY)
        gb.Background     = SolidColorBrush(self._CLR_WHITE)
        Grid.SetRow(gb, row)
        return gb

    def _make_button(self, text, is_primary=True, small=False):
        btn = Button()
        btn.Content    = text
        btn.FontWeight = FontWeights.SemiBold
        btn.Cursor     = Cursors.Hand
        btn.FontSize   = 11 if small else 12
        btn.Padding    = Thickness(12, 5, 12, 5) if small else Thickness(20, 8, 20, 8)
        if is_primary:
            btn.Background      = SolidColorBrush(self._CLR_PRIMARY)
            btn.Foreground      = SolidColorBrush(Color.FromRgb(255, 255, 255))
            btn.BorderThickness = Thickness(0)
        else:
            btn.Background      = SolidColorBrush(Color.FromRgb(224, 224, 224))
            btn.Foreground      = SolidColorBrush(Color.FromRgb(50, 50, 50))
            btn.BorderThickness = Thickness(0)
        return btn

    def _make_badge(self, text, bg_color, fg_color):
        """Tạo badge nhỏ có nền màu."""
        b = Border()
        b.Background      = SolidColorBrush(bg_color)
        b.CornerRadius    = System.Windows.CornerRadius(3)
        b.Padding         = Thickness(4, 1, 4, 1)
        b.VerticalAlignment = VerticalAlignment.Center

        tb = TextBlock()
        tb.Text       = text
        tb.FontSize   = 9
        tb.FontWeight = FontWeights.Bold
        tb.Foreground = SolidColorBrush(fg_color)
        b.Child = tb
        return b

    def _make_pc_row(self, entry, for_source=True):
        """
        Tạo nội dung hiển thị cho một PointCloudEntry.
        Hiển thị: [badge] ☁ Tên  [ID: xxx]  (+ tên file link nếu có)
        """
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal

        # Badge LOCAL / LINK
        if entry.is_linked:
            badge = self._make_badge(u" LINK ", self._CLR_LINK_BADGE, Color.FromRgb(255, 255, 255))
        else:
            badge = self._make_badge(u" LOCAL ", self._CLR_PRIMARY, Color.FromRgb(255, 255, 255))
        badge.Margin = Thickness(0, 0, 6, 0)

        # Cloud icon
        icon = TextBlock()
        icon.Text              = u"\u2601 "
        icon.FontSize          = 13
        icon.Foreground        = SolidColorBrush(
            Color.FromRgb(34, 139, 34) if entry.is_linked else Color.FromRgb(70, 130, 180))
        icon.VerticalAlignment = VerticalAlignment.Center

        # Tên PC
        name_tb = TextBlock()
        name_tb.Text              = entry.display_name
        name_tb.FontWeight        = FontWeights.SemiBold
        name_tb.VerticalAlignment = VerticalAlignment.Center

        # ID
        id_tb = TextBlock()
        id_tb.Text              = u"  [ID: {}]".format(entry.element_id)
        id_tb.FontSize          = 10
        id_tb.Foreground        = SolidColorBrush(self._CLR_TEXT_GRAY)
        id_tb.VerticalAlignment = VerticalAlignment.Center

        sp.Children.Add(badge)
        sp.Children.Add(icon)
        sp.Children.Add(name_tb)
        sp.Children.Add(id_tb)

        # Nếu là PC từ file link → hiển thị thêm tên file link
        if entry.is_linked:
            link_tb = TextBlock()
            link_tb.Text              = u"  \u2190 {}".format(entry.link_doc_name)
            link_tb.FontSize          = 10
            link_tb.Foreground        = SolidColorBrush(self._CLR_LINK_BADGE)
            link_tb.FontWeight        = FontWeights.SemiBold
            link_tb.VerticalAlignment = VerticalAlignment.Center
            sp.Children.Add(link_tb)

        return sp

    # ─── Populate Lists ───────────────────────────────────────────────────────

    def _populate_source_list(self):
        """Đổ dữ liệu vào danh sách Source (local + linked)."""
        self._src_listbox.Items.Clear()

        # Nhóm: local trước, linked sau
        local_group  = [e for e in self._source_entries if not e.is_linked]
        linked_group = [e for e in self._source_entries if e.is_linked]

        if local_group:
            self._add_list_separator(u"── Current Document ──")
            for entry in local_group:
                self._add_source_item(entry)

        if linked_group:
            self._add_list_separator(u"── Linked Files ──")
            for entry in linked_group:
                self._add_source_item(entry)

        if not local_group and not linked_group:
            sep = ListBoxItem()
            sep.Content    = u"(No Point Cloud found)"
            sep.IsEnabled  = False
            sep.Foreground = SolidColorBrush(self._CLR_TEXT_GRAY)
            self._src_listbox.Items.Add(sep)

    def _add_list_separator(self, text):
        """Thêm dòng tiêu đề nhóm (không chọn được)."""
        sep = ListBoxItem()
        tb  = TextBlock()
        tb.Text       = text
        tb.FontSize   = 10
        tb.FontWeight = FontWeights.SemiBold
        tb.Foreground = SolidColorBrush(self._CLR_TEXT_GRAY)
        tb.Margin     = Thickness(4, 4, 4, 2)
        sep.Content   = tb
        sep.IsEnabled = False
        sep.Padding   = Thickness(4, 0, 4, 0)
        self._src_listbox.Items.Add(sep)

    def _add_source_item(self, entry):
        item = ListBoxItem()
        item.Content = self._make_pc_row(entry, for_source=True)
        item.Tag     = entry
        item.Padding = Thickness(8, 6, 8, 6)
        # Nền xanh lá nhạt cho linked item
        if entry.is_linked:
            item.Background = SolidColorBrush(self._CLR_LINK_BG)
        self._src_listbox.Items.Add(item)

    def _populate_target_list(self, exclude_entry):
        """
        Đổ dữ liệu vào danh sách Target.
        Chỉ hiển thị LOCAL PC, loại trừ entry đang được chọn làm source (nếu là local).
        """
        self._tgt_panel.Children.Clear()
        self._target_items = []

        exclude_id = exclude_entry.element_id if (not exclude_entry.is_linked) else None

        for entry in self._local_entries:
            if entry.element_id == exclude_id:
                continue

            row_border = Border()
            row_border.Padding         = Thickness(6, 5, 6, 5)
            row_border.BorderBrush     = SolidColorBrush(Color.FromRgb(230, 230, 230))
            row_border.BorderThickness = Thickness(0, 0, 0, 1)

            cb = CheckBox()
            cb.Content           = self._make_pc_row(entry, for_source=False)
            cb.Tag               = entry
            cb.VerticalAlignment = VerticalAlignment.Center
            cb.Checked          += self._on_checkbox_changed
            cb.Unchecked        += self._on_checkbox_changed

            row_border.Child = cb
            self._tgt_panel.Children.Add(row_border)
            self._target_items.append((cb, entry))

        # Nếu không có local PC nào để làm target
        if not self._target_items:
            no_item = TextBlock()
            no_item.Text       = u"(No local Point Cloud available for child)"
            no_item.FontSize   = 11
            no_item.Foreground = SolidColorBrush(self._CLR_TEXT_GRAY)
            no_item.Margin     = Thickness(8, 8, 8, 8)
            self._tgt_panel.Children.Add(no_item)

        self._update_count()

    # ─── Event Handlers ───────────────────────────────────────────────────────

    def _on_source_changed(self, sender, e):
        item = self._src_listbox.SelectedItem
        if item is None or not hasattr(item, 'Tag') or item.Tag is None:
            self._source_entry = None
            self._info_border.Visibility = Visibility.Collapsed
            self._tgt_panel.Children.Clear()
            self._target_items = []
            self._update_ok_button()
            return

        self._source_entry = item.Tag  # PointCloudEntry

        # Hiển thị world-space transform info
        tf = self._source_entry.world_transform
        extra = u""
        if self._source_entry.is_linked:
            extra = u"  [Source: {}]\n".format(self._source_entry.link_doc_name)
        self._info_text.Text    = extra + describe_transform(tf)
        self._info_border.Visibility = Visibility.Visible

        # Cập nhật danh sách target
        self._populate_target_list(self._source_entry)
        self._update_ok_button()

    def _on_checkbox_changed(self, sender, e):
        self._update_count()
        self._update_ok_button()

    def _on_select_all(self, sender, e):
        for cb, _ in self._target_items:
            cb.IsChecked = True

    def _on_deselect_all(self, sender, e):
        for cb, _ in self._target_items:
            cb.IsChecked = False

    def _on_ok(self, sender, e):
        self.SelectedSource  = self._source_entry
        self.SelectedTargets = [entry for cb, entry in self._target_items if cb.IsChecked == True]
        self.Confirmed       = True
        self.Close()

    def _on_cancel(self, sender, e):
        self.Confirmed = False
        self.Close()

    # ─── State helpers ────────────────────────────────────────────────────────

    def _get_checked_count(self):
        return sum(1 for cb, _ in self._target_items if cb.IsChecked == True)

    def _update_count(self):
        n = self._get_checked_count()
        self._count_text.Text = u"(Selected: {})".format(n) if n > 0 else u""

    def _update_ok_button(self):
        has_source  = self._source_entry is not None
        has_targets = self._get_checked_count() > 0
        self._ok_btn.IsEnabled = has_source and has_targets


# ═════════════════════════════════════════════════════════════════════════════
# PHẦN 3 - MAIN: Chạy tool
# ═════════════════════════════════════════════════════════════════════════════

def main():
    # 1. Thu thập Point Cloud Entries
    local_entries, source_entries = collect_all_point_cloud_entries()

    # Cần ít nhất 1 source entry và 1 local entry để làm target
    if not source_entries:
        TaskDialog.Show(
            u"Notice",
            u"No Point Cloud found in the current document or any loaded linked file."
        )
        return

    if not local_entries:
        TaskDialog.Show(
            u"Notice",
            u"No local Point Cloud found in the current document to use as a child target."
        )
        return

    if len(source_entries) < 1 or len(local_entries) < 1:
        TaskDialog.Show(
            u"Notice",
            u"At least 1 source and 1 child Point Cloud are required."
        )
        return

    # 2. Mở dialog
    dialog = AlignPointCloudDialog(local_entries, source_entries)
    dialog.ShowDialog()

    if not dialog.Confirmed:
        return

    source_entry  = dialog.SelectedSource
    target_entries = dialog.SelectedTargets

    if not source_entry or not target_entries:
        TaskDialog.Show(u"Error", u"Please select both a source and at least one child Point Cloud.")
        return

    # 3. Lấy world-space transform của source
    source_world_transform = source_entry.world_transform

    # 4. Áp dụng trong Transaction
    success_count = 0
    errors        = []

    with Transaction(doc, u"Align Point Clouds (v2)") as trans:
        trans.Start()
        for entry in target_entries:
            try:
                apply_transform_to_target(entry, source_world_transform)
                success_count += 1
            except Exception as ex:
                errors.append(u"\u2022 {}: {}".format(entry.display_name, str(ex)))
        trans.Commit()

    # 5. Thông báo kết quả
    src_label = source_entry.display_name
    if source_entry.is_linked:
        src_label += u" [from {}]".format(source_entry.link_doc_name)

    msg = u"\u2713 Successfully moved {}/{} Point Cloud(s).\n\nSource: {}".format(
        success_count, len(target_entries), src_label)

    if errors:
        msg += u"\n\nErrors occurred:\n" + u"\n".join(errors)

    TaskDialog.Show(u"Done", msg)


# ─── Entry point ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()
else:
    main()
