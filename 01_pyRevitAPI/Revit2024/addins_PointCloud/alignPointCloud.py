# -*- coding: utf-8 -*-
"""
Point Cloud Aligner - IronPython Script for Revit (Dynamo)
===========================================================
Purpose:
  - Select the SOURCE Point Cloud (file A that has been moved / rotated)
  - Select CHILD Point Clouds (B, C, ...)
  - Tool automatically applies the same transform to all child files

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
doc    = DocumentManager.Instance.CurrentDBDocument
uiapp  = DocumentManager.Instance.CurrentUIApplication
app    = uiapp.Application
uidoc  = uiapp.ActiveUIDocument


# ═════════════════════════════════════════════════════════════════════════════
# PHẦN 1 - HELPERS: Tính toán Transform
# ═════════════════════════════════════════════════════════════════════════════

ANGLE_TOL       = 1e-9
TRANSLATE_TOL   = 1e-6


def get_point_cloud_name(instance):
    """Lấy tên file Point Cloud từ PointCloudInstance."""
    try:
        # Ưu tiên: Tên instance (thường là tên file RCS/RCP)
        name = instance.Name
        if name and name != "Point Cloud":
            import System.IO.Path as Path
            # Lấy tên file không extension nếu có path
            if name.lower().endswith(('.rcs', '.rcp')):
                return System.IO.Path.GetFileNameWithoutExtension(name)
            return name
        
        # Fallback: Path từ PointCloudType
        pc_type = doc.GetElement(instance.GetTypeId())
        if pc_type:
            path = pc_type.GetPath()  # Revit API: PointCloudType.GetPath()
            if path:
                return System.IO.Path.GetFileNameWithoutExtension(path)
            return pc_type.Name
    except:
        pass
    return "PointCloud [{}]".format(instance.Id)


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


def apply_transform_to_target(target_instance, source_transform):
    """
    Áp dụng transform của file gốc lên file con.
    delta = source_transform × target_transform⁻¹
    Sau đó: rotate → translate
    """
    tgt_tf = target_instance.GetTransform()

    # delta transform
    delta = source_transform.Multiply(tgt_tf.Inverse)

    axis, angle = extract_axis_angle(delta)

    # 1. Xoay
    if abs(angle) > ANGLE_TOL:
        tgt_origin = tgt_tf.Origin
        rot_axis_line = Line.CreateUnbound(tgt_origin, axis)
        ElementTransformUtils.RotateElement(doc, target_instance.Id, rot_axis_line, angle)
        # Refresh
        tgt_tf = target_instance.GetTransform()

    # 2. Dịch chuyển
    translation = source_transform.Origin - tgt_tf.Origin
    if translation.GetLength() > TRANSLATE_TOL:
        ElementTransformUtils.MoveElement(doc, target_instance.Id, translation)


# ═════════════════════════════════════════════════════════════════════════════
# PHẦN 2 - UI: WPF Dialog
# ═════════════════════════════════════════════════════════════════════════════

class AlignPointCloudDialog(Window):

    # ── Màu sắc ──────────────────────────────────────────────────────────────
    _CLR_PRIMARY     = Color.FromRgb(0,   102, 204)  # #0066CC
    _CLR_PRIMARY_DRK = Color.FromRgb(0,   82,  163)  # #0052A3
    _CLR_BG          = Color.FromRgb(245, 245, 245)  # #F5F5F5
    _CLR_WHITE       = Color.FromRgb(255, 255, 255)
    _CLR_WARN_BG     = Color.FromRgb(255, 248, 230)  # #FFF8E6
    _CLR_WARN_BD     = Color.FromRgb(255, 184,  77)  # #FFB84D
    _CLR_INFO_BG     = Color.FromRgb(240, 247, 255)  # #F0F7FF
    _CLR_INFO_BD     = Color.FromRgb(179, 209, 255)  # #B3D1FF
    _CLR_GRAY        = Color.FromRgb(200, 200, 200)
    _CLR_TEXT_GRAY   = Color.FromRgb(120, 120, 120)

    def __init__(self, all_point_clouds):
        self._all_pcs     = all_point_clouds      # list of PointCloudInstance
        self._source_pc   = None
        self._target_items = []                   # list of (CheckBox, PointCloudInstance)

        self.SelectedSource  = None
        self.SelectedTargets = []
        self.Confirmed       = False

        self._build_window()
        self._populate_source_list()

    # ─── Window shell ────────────────────────────────────────────────────────

    def _build_window(self):
        self.Title                 = u"Point Cloud Aligner"
        self.Width                 = 500
        self.Height                = 590
        self.MinWidth              = 420
        self.MinHeight             = 480
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background            = SolidColorBrush(self._CLR_BG)
        self.FontFamily            = FontFamily("Segoe UI")
        self.FontSize              = 12

        root = Grid()
        self.Content = root

        # Row definitions
        for h in [GridLength.Auto,          # Header
                  GridLength.Auto,          # Group: source
                  GridLength(1, GridUnitType.Star),  # Group: targets
                  GridLength.Auto,          # Warning
                  GridLength.Auto]:         # Buttons
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
        t1.Text       = u"\u2601 Point Cloud Aligner"
        t1.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        t1.FontSize   = 16
        t1.FontWeight = FontWeights.Bold

        t2  = TextBlock()
        t2.Text       = u"Move child Point Clouds to match the position of the source file"
        t2.Foreground = SolidColorBrush(Color.FromRgb(204, 224, 255))
        t2.FontSize   = 11
        t2.Margin     = Thickness(0, 3, 0, 0)

        hsp.Children.Add(t1)
        hsp.Children.Add(t2)
        header.Child = hsp
        root.Children.Add(header)

        # ── Group: Source ────────────────────────────────────────────────────
        src_group = self._make_groupbox(u"①  Select SOURCE Point Cloud  (file A)", row=1)
        src_inner = StackPanel()
        src_inner.Margin = Thickness(4)

        src_hint = TextBlock()
        src_hint.Text      = u"Select the file that has already been moved / rotated as reference:"
        src_hint.FontSize  = 11
        src_hint.Foreground = SolidColorBrush(Color.FromRgb(80, 80, 80))
        src_hint.Margin    = Thickness(0, 0, 0, 6)
        src_inner.Children.Add(src_hint)

        # ListBox nguồn (single select)
        self._src_listbox = ListBox()
        self._src_listbox.Height          = 110
        self._src_listbox.SelectionMode   = SelectionMode.Single
        self._src_listbox.BorderBrush     = SolidColorBrush(self._CLR_GRAY)
        self._src_listbox.BorderThickness = Thickness(1)
        self._src_listbox.SelectionChanged += self._on_source_changed
        src_inner.Children.Add(self._src_listbox)

        # Info transform
        self._info_border = Border()
        self._info_border.Background     = SolidColorBrush(self._CLR_INFO_BG)
        self._info_border.BorderBrush    = SolidColorBrush(self._CLR_INFO_BD)
        self._info_border.BorderThickness = Thickness(1)
        self._info_border.CornerRadius   = System.Windows.CornerRadius(3)
        self._info_border.Padding        = Thickness(8, 6, 8, 6)
        self._info_border.Margin         = Thickness(0, 7, 0, 0)
        self._info_border.Visibility     = Visibility.Collapsed

        info_sp = StackPanel()
        info_title = TextBlock()
        info_title.Text      = u"Transform of source file:"
        info_title.FontWeight = FontWeights.SemiBold
        info_title.FontSize   = 11
        info_title.Foreground = SolidColorBrush(self._CLR_PRIMARY_DRK)

        self._info_text = TextBlock()
        self._info_text.FontSize   = 11
        self._info_text.Foreground = SolidColorBrush(Color.FromRgb(50, 50, 50))
        self._info_text.Margin     = Thickness(0, 4, 0, 0)
        self._info_text.FontFamily = FontFamily("Consolas")

        info_sp.Children.Add(info_title)
        info_sp.Children.Add(self._info_text)
        self._info_border.Child = info_sp
        src_inner.Children.Add(self._info_border)

        src_group.Content = src_inner
        root.Children.Add(src_group)

        # ── Group: Targets ───────────────────────────────────────────────────
        tgt_group = self._make_groupbox(u"Select CHILD Point Clouds to move", row=2)
        tgt_inner = Grid()
        tgt_inner.Margin = Thickness(4)

        for h in [GridLength.Auto, GridLength(1, GridUnitType.Star), GridLength.Auto]:
            rd2 = RowDefinition(); rd2.Height = h
            tgt_inner.RowDefinitions.Add(rd2)

        # Hint + count
        hint_row = StackPanel()
        hint_row.Orientation = Orientation.Horizontal
        hint_row.Margin      = Thickness(0, 0, 0, 6)
        Grid.SetRow(hint_row, 0)

        tgt_hint = TextBlock()
        tgt_hint.Text      = u"Select one or more files:"
        tgt_hint.FontSize  = 11
        tgt_hint.Foreground = SolidColorBrush(Color.FromRgb(80, 80, 80))
        tgt_hint.VerticalAlignment = VerticalAlignment.Center

        self._count_text = TextBlock()
        self._count_text.FontSize   = 11
        self._count_text.FontWeight = FontWeights.SemiBold
        self._count_text.Foreground = SolidColorBrush(self._CLR_PRIMARY)
        self._count_text.Margin     = Thickness(8, 0, 0, 0)
        self._count_text.VerticalAlignment = VerticalAlignment.Center

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
        btn_row.Orientation          = Orientation.Horizontal
        btn_row.HorizontalAlignment  = HorizontalAlignment.Right
        btn_row.Margin               = Thickness(0, 6, 0, 0)
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
        warn.Background     = SolidColorBrush(self._CLR_WARN_BG)
        warn.BorderBrush    = SolidColorBrush(self._CLR_WARN_BD)
        warn.BorderThickness = Thickness(1)
        warn.CornerRadius   = System.Windows.CornerRadius(3)
        warn.Margin         = Thickness(10, 4, 10, 4)
        warn.Padding        = Thickness(10, 7, 10, 7)
        Grid.SetRow(warn, 3)

        warn_sp = StackPanel()
        warn_sp.Orientation = Orientation.Horizontal

        w_icon = TextBlock()
        w_icon.Text                = u"\u26a0"
        w_icon.FontSize            = 14
        w_icon.Foreground          = SolidColorBrush(Color.FromRgb(204, 119, 0))
        w_icon.VerticalAlignment   = VerticalAlignment.Top

        w_text = TextBlock()
        w_text.Text        = u"  This action cannot be undone after saving the file. Please review carefully before applying."
        w_text.FontSize    = 11
        w_text.Foreground  = SolidColorBrush(Color.FromRgb(128, 85, 0))
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
        Grid.SetRow(btn_bar, 4)

        self._cancel_btn = self._make_button(u"Cancel", is_primary=False)
        self._cancel_btn.Margin  = Thickness(0, 0, 10, 0)
        self._cancel_btn.Click  += self._on_cancel

        self._ok_btn = self._make_button(u"\u2713 Apply", is_primary=True)
        self._ok_btn.IsEnabled = False
        self._ok_btn.Click    += self._on_ok

        btn_bar.Children.Add(self._cancel_btn)
        btn_bar.Children.Add(self._ok_btn)
        root.Children.Add(btn_bar)

    # ─── UI Helpers ──────────────────────────────────────────────────────────

    def _make_groupbox(self, header_text, row):
        gb = GroupBox()
        gb.Header          = header_text
        gb.Margin          = Thickness(10, 6, 10, 4)
        gb.Padding         = Thickness(8)
        gb.BorderBrush     = SolidColorBrush(self._CLR_GRAY)
        gb.Background      = SolidColorBrush(self._CLR_WHITE)
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
            btn.Background = SolidColorBrush(self._CLR_PRIMARY)
            btn.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
            btn.BorderThickness = Thickness(0)
        else:
            btn.Background = SolidColorBrush(Color.FromRgb(224, 224, 224))
            btn.Foreground = SolidColorBrush(Color.FromRgb(50, 50, 50))
            btn.BorderThickness = Thickness(0)
        return btn

    def _make_pc_item(self, pc, for_source=True):
        """Tạo một row item hiển thị tên Point Cloud."""
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal

        icon = TextBlock()
        icon.Text      = u"\u2601 "
        icon.FontSize  = 13
        icon.Foreground = SolidColorBrush(
            Color.FromRgb(70, 130, 180) if for_source else Color.FromRgb(100, 149, 237))
        icon.VerticalAlignment = VerticalAlignment.Center

        name = TextBlock()
        name.Text       = get_point_cloud_name(pc)
        name.FontWeight = FontWeights.SemiBold
        name.VerticalAlignment = VerticalAlignment.Center

        id_text = TextBlock()
        id_text.Text      = u"  [ID: {}]".format(pc.Id)
        id_text.FontSize  = 10
        id_text.Foreground = SolidColorBrush(self._CLR_TEXT_GRAY)
        id_text.VerticalAlignment = VerticalAlignment.Center

        sp.Children.Add(icon)
        sp.Children.Add(name)
        sp.Children.Add(id_text)
        return sp

    # ─── Populate Lists ───────────────────────────────────────────────────────

    def _populate_source_list(self):
        self._src_listbox.Items.Clear()
        for pc in self._all_pcs:
            item = ListBoxItem()
            item.Content = self._make_pc_item(pc, for_source=True)
            item.Tag     = pc
            item.Padding = Thickness(8, 6, 8, 6)
            self._src_listbox.Items.Add(item)

    def _populate_target_list(self, exclude_id):
        self._tgt_panel.Children.Clear()
        self._target_items = []

        for pc in self._all_pcs:
            if pc.Id == exclude_id:
                continue

            # Mỗi item = Border chứa CheckBox + StackPanel
            row_border = Border()
            row_border.Padding         = Thickness(6, 5, 6, 5)
            row_border.BorderBrush     = SolidColorBrush(Color.FromRgb(230, 230, 230))
            row_border.BorderThickness = Thickness(0, 0, 0, 1)

            cb = CheckBox()
            cb.Content             = self._make_pc_item(pc, for_source=False)
            cb.Tag                 = pc
            cb.VerticalAlignment   = VerticalAlignment.Center
            cb.Checked            += self._on_checkbox_changed
            cb.Unchecked          += self._on_checkbox_changed

            row_border.Child = cb
            self._tgt_panel.Children.Add(row_border)
            self._target_items.append((cb, pc))

        self._update_count()

    # ─── Event Handlers ───────────────────────────────────────────────────────

    def _on_source_changed(self, sender, e):
        item = self._src_listbox.SelectedItem
        if item is None:
            self._source_pc = None
            self._info_border.Visibility = Visibility.Collapsed
            self._tgt_panel.Children.Clear()
            self._target_items = []
            self._update_ok_button()
            return

        self._source_pc = item.Tag

        # Hiển thị transform info
        tf = self._source_pc.GetTransform()
        self._info_text.Text    = describe_transform(tf)
        self._info_border.Visibility = Visibility.Visible

        # Cập nhật danh sách target
        self._populate_target_list(self._source_pc.Id)
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
        self.SelectedSource  = self._source_pc
        self.SelectedTargets = [pc for cb, pc in self._target_items if cb.IsChecked == True]
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
        has_source  = self._source_pc is not None
        has_targets = self._get_checked_count() > 0
        self._ok_btn.IsEnabled = has_source and has_targets


# ═════════════════════════════════════════════════════════════════════════════
# PHẦN 3 - MAIN: Chạy tool
# ═════════════════════════════════════════════════════════════════════════════

def main():
    # 1. Lấy tất cả Point Cloud Instance
    all_pcs = list(
        FilteredElementCollector(doc)
        .OfClass(PointCloudInstance)
    )

    if len(all_pcs) < 2:
        TaskDialog.Show(
            u"Notice",
            u"At least 2 Point Clouds are required in the project.\n(1 source file + 1 child file)"
        )
        return

    # 2. Mở dialog
    dialog = AlignPointCloudDialog(all_pcs)
    dialog.ShowDialog()

    if not dialog.Confirmed:
        return

    source_pc    = dialog.SelectedSource
    target_pcs   = dialog.SelectedTargets

    if not source_pc or not target_pcs:
        TaskDialog.Show(u"Error", u"Please select both a source and at least one child Point Cloud.")
        return

    # 3. Lấy transform gốc
    source_transform = source_pc.GetTransform()

    # 4. Áp dụng trong Transaction
    success_count = 0
    errors        = []

    with Transaction(doc, u"Align Point Clouds") as trans:
        trans.Start()
        for pc in target_pcs:
            try:
                apply_transform_to_target(pc, source_transform)
                success_count += 1
            except Exception as ex:
                errors.append(u"• {}: {}".format(get_point_cloud_name(pc), str(ex)))
        trans.Commit()

    # 5. Thông báo kết quả
    msg = u"\u2713 Successfully moved {}/{} Point Cloud(s).".format(
        success_count, len(target_pcs))
    if errors:
        msg += u"\n\nErrors occurred:\n" + u"\n".join(errors)

    TaskDialog.Show(u"Done", msg)


# ─── Entry point ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()
else:
    main()