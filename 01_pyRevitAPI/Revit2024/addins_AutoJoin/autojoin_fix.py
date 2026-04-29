# -*- coding: utf-8 -*-
"""
Alpha BIM - Auto Join Tool
Fixed & Optimized by Gemini Code Assist
IronPython for Dynamo / Revit
"""

# ════════════════════════════════════════════════════════
#  IMPORTS
# ════════════════════════════════════════════════════════
import clr
import System
import json
import io

# ── Revit API ──────────────────────────────────────────
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import (
    FilteredElementCollector, ElementId, Transaction,
    JoinGeometryUtils, BuiltInCategory,
    FailureHandlingOptions, FailureProcessingResult
)
from Autodesk.Revit.DB import IFailuresPreprocessor

# ── Dynamo / RevitServices ─────────────────────────────
clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager

# ── WPF ────────────────────────────────────────────────
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Data")

from System.Data import DataTable
from System.Collections.Generic import List as NetList
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage, PropertyPath
from System.Windows.Controls import (
    DataGridComboBoxColumn, DataGridCheckBoxColumn,
    DataGridLength, DataGridLengthUnitType, MenuItem
)
from System.Windows.Data import Binding
from System.Windows.Markup import XamlReader

# ── WinForms (file dialogs) ────────────────────────────
clr.AddReference("System.Windows.Forms")
import System.Windows.Forms as WinForms

# ════════════════════════════════════════════════════════
#  REVIT CONTEXT  (Dynamo style)
# ════════════════════════════════════════════════════════
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
uidoc = uiapp.ActiveUIDocument

# ════════════════════════════════════════════════════════
#  CATEGORY REGISTRY
# ════════════════════════════════════════════════════════
CATEGORY_MAP = {
    "Beam":          BuiltInCategory.OST_StructuralFraming,
    "Ceiling":       BuiltInCategory.OST_Ceilings,
    "Column":        BuiltInCategory.OST_Columns,
    "Floor":         BuiltInCategory.OST_Floors,
    "Foundation":    BuiltInCategory.OST_StructuralFoundation,
    "Generic Model": BuiltInCategory.OST_GenericModel,
    "Roof":          BuiltInCategory.OST_Roofs,
    "Wall":          BuiltInCategory.OST_Walls,
}

PRIORITY_CATS  = sorted(CATEGORY_MAP.keys())
JOIN_WITH_CATS = ["<All>"] + sorted(CATEGORY_MAP.keys())

# ════════════════════════════════════════════════════════
#  DATA TABLE HELPERS
# ════════════════════════════════════════════════════════
def make_table():
    t = DataTable("Rules")
    t.Columns.Add("PriorityCategory", System.String)
    t.Columns.Add("JoinWithCategory",  System.String)
    t.Columns.Add("Reverse",           System.Boolean)
    return t

def table_add_row(tbl, pc="Wall", jw="<All>", rev=False):
    r = tbl.NewRow()
    r["PriorityCategory"] = pc
    r["JoinWithCategory"]  = jw
    r["Reverse"]           = rev
    tbl.Rows.Add(r)

def table_to_rules(tbl):
    return [
        {
            "priority":  str(row["PriorityCategory"]),
            "join_with": str(row["JoinWithCategory"]),
            "reverse":   bool(row["Reverse"]),
        }
        for row in tbl.Rows
    ]

def rules_to_table(rules, tbl):
    tbl.Clear()
    for r in rules:
        table_add_row(
            tbl,
            r.get("priority",  "Wall"),
            r.get("join_with", "<All>"),
            bool(r.get("reverse", False))
        )

# ════════════════════════════════════════════════════════
#  REVIT JOIN LOGIC & OPTIMIZATION CACHE
# ════════════════════════════════════════════════════════
def get_id_val(eid):
    """Helper for Revit 2023 (IntegerValue) and 2024+ (Value) compatibility"""
    try: return eid.Value
    except: return eid.IntegerValue

_elem_cache = {}
def collect_elements(bic, scope_ids=None):
    cache_key = (bic, scope_ids is not None)
    if cache_key not in _elem_cache:
        try:
            if scope_ids:
                col = FilteredElementCollector(doc, NetList[ElementId](scope_ids))
            else:
                col = FilteredElementCollector(doc, doc.ActiveView.Id)
            _elem_cache[cache_key] = list(col.OfCategory(bic).WhereElementIsNotElementType())
        except Exception:
            _elem_cache[cache_key] = []
    return _elem_cache[cache_key]

_bbox_cache = {}
def bboxes_overlap(e1, e2, tol=0.01):
    try:
        e1_id = get_id_val(e1.Id)
        e2_id = get_id_val(e2.Id)
        
        if e1_id not in _bbox_cache:
            _bbox_cache[e1_id] = e1.get_BoundingBox(None)
        if e2_id not in _bbox_cache:
            _bbox_cache[e2_id] = e2.get_BoundingBox(None)
            
        b1 = _bbox_cache[e1_id]
        b2 = _bbox_cache[e2_id]
        
        if not b1 or not b2:
            return False
        return not (
            b1.Max.X + tol < b2.Min.X or b2.Max.X + tol < b1.Min.X or
            b1.Max.Y + tol < b2.Min.Y or b2.Max.Y + tol < b1.Min.Y or
            b1.Max.Z + tol < b2.Min.Z or b2.Max.Z + tol < b1.Min.Z
        )
    except Exception:
        return False

# ────────────────────────────────────────────────────────
#  Failure preprocessor – silently dismiss non-fatal
#  Revit warnings such as "Can't cut joined element"
# ────────────────────────────────────────────────────────
class _SilentPreprocessor(IFailuresPreprocessor):
    def PreprocessFailures(self, fa):
        from Autodesk.Revit.DB import FailureSeverity
        for msg in list(fa.GetFailureMessages()):
            if msg.GetSeverity() == FailureSeverity.Warning:
                fa.DeleteWarning(msg)
        return FailureProcessingResult.Continue

def _join_pair(doc, e1, e2, rev, do_unjoin, tx):
    """
    Join e1–e2 inside an already-open transaction.
    Returns True if a join was created or already existed.
    """
    is_joined = JoinGeometryUtils.AreElementsJoined(doc, e1, e2)

    if do_unjoin and is_joined:
        JoinGeometryUtils.UnjoinGeometry(doc, e1, e2)
        is_joined = False

    if not is_joined:
        JoinGeometryUtils.JoinGeometry(doc, e1, e2)
        is_joined = True

    if is_joined:
        try:
            e1_cuts = JoinGeometryUtils.IsCuttingElementInJoin(doc, e1, e2)
            need_switch = (not rev and not e1_cuts) or (rev and e1_cuts)
            if need_switch:
                JoinGeometryUtils.SwitchJoinOrder(doc, e1, e2)
        except Exception:
            pass

    return is_joined

def run_autojoin(rules, scope="view", scope_ids=None, do_unjoin=False):
    """Execute all join rules."""
    ids          = scope_ids if (scope == "selected" and scope_ids) else None
    joined_count = 0
    warnings     = []
    
    # Clear caches at the start of autojoin
    _elem_cache.clear()
    _bbox_cache.clear()

    tx = Transaction(doc, "AutoJoin")

    # ── Suppress warnings ──
    fho = tx.GetFailureHandlingOptions()
    fho.SetFailuresPreprocessor(_SilentPreprocessor())
    fho.SetClearAfterRollback(True)
    tx.SetFailureHandlingOptions(fho)

    tx.Start()
    try:
        for rule in rules:
            pc  = rule["priority"]
            jw  = rule["join_with"]
            rev = rule["reverse"]
            if pc not in CATEGORY_MAP:
                continue

            elems1 = collect_elements(CATEGORY_MAP[pc], ids)

            if jw == "<All>":
                elems2 = []
                for cat_name, bic in CATEGORY_MAP.items():
                    if cat_name != pc:
                        elems2.extend(collect_elements(bic, ids))
            elif jw in CATEGORY_MAP:
                elems2 = collect_elements(CATEGORY_MAP[jw], ids)
            else:
                continue

            seen = set()
            for e1 in elems1:
                for e2 in elems2:
                    if e1.Id == e2.Id:
                        continue
                    id1 = get_id_val(e1.Id)
                    id2 = get_id_val(e2.Id)
                    key = (
                        min(id1, id2),
                        max(id1, id2),
                    )
                    if key in seen:
                        continue
                    seen.add(key)

                    if not bboxes_overlap(e1, e2):
                        continue

                    try:
                        if _join_pair(doc, e1, e2, rev, do_unjoin, tx):
                            joined_count += 1
                    except Exception as pair_ex:
                        warnings.append(u"[{}-{}] {}".format(
                            get_id_val(e1.Id),
                            get_id_val(e2.Id),
                            str(pair_ex)))

        tx.Commit()

    except Exception as ex:
        if tx.HasStarted() and not tx.HasEnded():
            tx.RollbackIfPossible()
        raise ex

    return joined_count, warnings

# ════════════════════════════════════════════════════════
#  WPF WINDOW XAML
# ════════════════════════════════════════════════════════
WINDOW_XAML = u"""<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="V2D - Auto Join"
    Width="600" Height="500"
    MinWidth="480" MinHeight="400"
    WindowStartupLocation="CenterScreen"
    FontFamily="Segoe UI" FontSize="12"
    Background="White">

  <Window.Resources>
    <Style x:Key="ToolBtn" TargetType="Button">
      <Setter Property="Background"  Value="#F0F0F0"/>
      <Setter Property="BorderBrush" Value="#BBBBBB"/>
      <Setter Property="Padding"     Value="8,0"/>
      <Setter Property="Height"      Value="26"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#E0E0E0"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
          <Setter Property="Background" Value="#D0D0D0"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="OkBtn" TargetType="Button">
      <Setter Property="Background"  Value="#43A047"/>
      <Setter Property="Foreground"  Value="White"/>
      <Setter Property="BorderBrush" Value="#2E7D32"/>
      <Setter Property="Cursor"      Value="Hand"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#2E7D32"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
          <Setter Property="Background" Value="#1B5E20"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="HeaderStyle" TargetType="DataGridColumnHeader">
      <Setter Property="Background"                Value="#43A047"/>
      <Setter Property="Foreground"                Value="White"/>
      <Setter Property="FontWeight"                Value="Bold"/>
      <Setter Property="Padding"                   Value="12,0"/>
      <Setter Property="Height"                    Value="32"/>
      <Setter Property="BorderBrush"               Value="#357A38"/>
      <Setter Property="BorderThickness"           Value="0,0,1,0"/>
      <Setter Property="HorizontalContentAlignment" Value="Left"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="38"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="52"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#F5F5F5" BorderBrush="#DDDDDD" BorderThickness="0,0,0,1">
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="8,0">
        <Button Name="btnImport" Content="Import" Style="{StaticResource ToolBtn}" Width="82" Margin="2,0" ToolTip="Import rules from JSON file"/>
        <Button Name="btnExport" Content="Export" Style="{StaticResource ToolBtn}" Width="82" Margin="2,0" ToolTip="Save current rules to JSON file"/>
        <Rectangle Width="1" Fill="#CCCCCC" Margin="8,5"/>
        <Button Name="btnAdd"    Content="+ Add Rule" Style="{StaticResource ToolBtn}" Width="90" Margin="2,0"/>
        <Button Name="btnDelete" Content="Delete" Style="{StaticResource ToolBtn}" Width="78" Margin="2,0"/>
      </StackPanel>
    </Border>

    <DataGrid Name="dgRules" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" SelectionMode="Extended" SelectionUnit="FullRow" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#E8E8E8" AlternatingRowBackground="#F1FFF2" RowBackground="White" BorderThickness="0" RowHeight="28" ColumnHeaderStyle="{StaticResource HeaderStyle}">
      <DataGrid.RowStyle>
        <Style TargetType="DataGridRow">
          <Style.Triggers>
            <Trigger Property="IsSelected" Value="True">
              <Setter Property="Background" Value="#BBDEFB"/>
              <Setter Property="Foreground" Value="#0D47A1"/>
            </Trigger>
          </Style.Triggers>
        </Style>
      </DataGrid.RowStyle>
      <DataGrid.ContextMenu>
        <ContextMenu FontFamily="Segoe UI" FontSize="12">
          <MenuItem Name="miNewRule"    Header="New Rule" ToolTip="Duplicate selected row and append"/>
          <MenuItem Name="miDeleteRule" Header="Delete selected Rules" ToolTip="Remove highlighted rows"/>
        </ContextMenu>
      </DataGrid.ContextMenu>
    </DataGrid>

    <Border Grid.Row="2" BorderBrush="#DDDDDD" BorderThickness="0,1,0,0" Background="#FAFAFA" Padding="12,10">
      <StackPanel>
        <TextBlock Text="PHAM VI DOI TUONG UU TIEN" FontWeight="Bold" FontSize="10" Foreground="#666666" Margin="0,0,0,6"/>
        <StackPanel Orientation="Horizontal">
          <RadioButton Name="rbView" Content="Tren View" IsChecked="True" GroupName="Scope" Margin="0,0,28,0" VerticalContentAlignment="Center"/>
          <RadioButton Name="rbSelected" Content="Tren cac doi tuong duoc chon" GroupName="Scope" VerticalContentAlignment="Center"/>
        </StackPanel>
      </StackPanel>
    </Border>

    <Border Grid.Row="3" Background="#EFEFEF" BorderBrush="#DDDDDD" BorderThickness="0,1,0,0">
      <Grid Margin="12,0">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Grid.Column="0">
          <CheckBox Name="chkUnjoin" Content="Unjoin" Margin="0,0,16,0" VerticalContentAlignment="Center" ToolTip="Unjoin truoc khi join lai"/>
        </StackPanel>
        <Button Name="btnOK" Grid.Column="1" Content="OK" Width="90" Height="30" FontWeight="Bold" FontSize="13" Margin="0,11" Style="{StaticResource OkBtn}"/>
      </Grid>
    </Border>
  </Grid>
</Window>"""

# ════════════════════════════════════════════════════════
#  WINDOW CONTROLLER
# ════════════════════════════════════════════════════════
class AutoJoinWindow(object):

    def __init__(self, preselected_ids):
        self.preselected_ids = list(preselected_ids)
        self.table = make_table()
        table_add_row(self.table, "Column", "<All>", False)  # default row

        self.window = XamlReader.Parse(WINDOW_XAML)
        self._find_controls()
        self._setup_columns()
        self._bind_events()

    def _find(self, name):
        return self.window.FindName(name)

    def _make_binding(self, path_str):
        b = Binding()
        b.Path = PropertyPath(path_str)
        return b

    def _find_controls(self):
        self.dg       = self._find("dgRules")
        self.rb_view  = self._find("rbView")
        self.rb_sel   = self._find("rbSelected")
        self.chk_un   = self._find("chkUnjoin")
        self.btn_ok   = self._find("btnOK")
        self.btn_imp  = self._find("btnImport")
        self.btn_exp  = self._find("btnExport")
        self.btn_add  = self._find("btnAdd")
        self.btn_del  = self._find("btnDelete")

        self.mi_new      = None
        self.mi_del_rule = None
        ctx = self.dg.ContextMenu
        if ctx:
            for item in ctx.Items:
                if isinstance(item, MenuItem):
                    hdr = str(item.Header) if item.Header is not None else ""
                    if "New Rule" in hdr:
                        self.mi_new = item
                    elif "Delete selected Rules" in hdr:
                        self.mi_del_rule = item

    def _setup_columns(self):
        col1 = DataGridComboBoxColumn()
        col1.Header = "Priority Category"
        col1.ItemsSource = NetList[System.String](PRIORITY_CATS)
        col1.SelectedItemBinding = self._make_binding("PriorityCategory")
        col1.Width = DataGridLength(1, DataGridLengthUnitType.Star)
        self.dg.Columns.Add(col1)

        col2 = DataGridComboBoxColumn()
        col2.Header = "Join With Category"
        col2.ItemsSource = NetList[System.String](JOIN_WITH_CATS)
        col2.SelectedItemBinding = self._make_binding("JoinWithCategory")
        col2.Width = DataGridLength(1, DataGridLengthUnitType.Star)
        self.dg.Columns.Add(col2)

        col3 = DataGridCheckBoxColumn()
        col3.Header = "Reverse"
        col3.Binding = self._make_binding("Reverse")
        col3.Width = DataGridLength(75)
        self.dg.Columns.Add(col3)

        self.dg.ItemsSource = self.table.DefaultView

    def _bind_events(self):
        self.btn_ok.Click  += self._on_ok
        self.btn_imp.Click += self._on_import
        self.btn_exp.Click += self._on_export
        self.btn_add.Click += self._on_add_rule
        self.btn_del.Click += self._on_delete_rule
        if self.mi_new:
            self.mi_new.Click += self._on_ctx_new
        if self.mi_del_rule:
            self.mi_del_rule.Click += self._on_ctx_delete

    def _on_add_rule(self, s, e):
        table_add_row(self.table)
        self.dg.Items.Refresh()

    def _on_delete_rule(self, s, e):
        self._delete_selected_rows()

    def _on_import(self, s, e):
        dlg = WinForms.OpenFileDialog()
        dlg.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        dlg.Title  = "Import AutoJoin Rules"
        if dlg.ShowDialog() != WinForms.DialogResult.OK:
            return
        try:
            with io.open(dlg.FileName, 'r', encoding='utf-8') as f:
                data = json.load(f)
            rules = data.get("rules", [])
            if not rules:
                MessageBox.Show(u"File khong chua rules hop le.",
                                "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                return
            rules_to_table(rules, self.table)
            self.dg.Items.Refresh()
            MessageBox.Show(
                u"Da import {} rule(s) tu:\n{}".format(len(rules), dlg.FileName),
                "Import OK", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            MessageBox.Show(u"Loi import:\n" + str(ex), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_export(self, s, e):
        dlg = WinForms.SaveFileDialog()
        dlg.Filter   = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        dlg.FileName = "AutoJoin_Rules.json"
        dlg.Title    = "Export AutoJoin Rules"
        if dlg.ShowDialog() != WinForms.DialogResult.OK:
            return
        try:
            rules = table_to_rules(self.table)
            with io.open(dlg.FileName, 'w', encoding='utf-8') as f:
                json.dump({"version": "1.0", "rules": rules},
                          f, indent=2, ensure_ascii=False)
            MessageBox.Show(
                u"Da luu {} rule(s) vao:\n{}".format(len(rules), dlg.FileName),
                "Export OK", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            MessageBox.Show(u"Loi export:\n" + str(ex), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)

    def _on_ctx_new(self, s, e):
        sel = self.dg.SelectedItem
        if sel is not None:
            row = self.table.NewRow()
            row["PriorityCategory"] = sel["PriorityCategory"]
            row["JoinWithCategory"]  = sel["JoinWithCategory"]
            row["Reverse"]           = sel["Reverse"]
            self.table.Rows.Add(row)
        else:
            table_add_row(self.table)
        self.dg.Items.Refresh()

    def _on_ctx_delete(self, s, e):
        self._delete_selected_rows()

    def _delete_selected_rows(self):
        items = list(self.dg.SelectedItems)
        for item in items:
            try:
                item.Row.Delete()
            except Exception:
                pass
        self.table.AcceptChanges()
        self.dg.Items.Refresh()

    def _on_ok(self, s, e):
        rules = table_to_rules(self.table)
        if not rules:
            MessageBox.Show(u"Vui long them it nhat 1 rule truoc khi chay.",
                            "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        scope     = "view" if bool(self.rb_view.IsChecked) else "selected"
        do_unjoin = bool(self.chk_un.IsChecked)
        scope_ids = self.preselected_ids if scope == "selected" else None

        if scope == "selected" and not scope_ids:
            res = MessageBox.Show(
                u"Khong co doi tuong nao duoc chon truoc.\n"
                u"Chuyen sang che do 'Tren View' khong?",
                u"Xac nhan", MessageBoxButton.YesNo, MessageBoxImage.Question)
            if res == System.Windows.MessageBoxResult.Yes:
                scope     = "view"
                scope_ids = None
            else:
                return

        self.window.Hide()
        try:
            count, warns = run_autojoin(rules, scope, scope_ids, do_unjoin)
            msg = u"AutoJoin hoan tat!\nSo cap da join / chinh thu tu: {}".format(count)
            if warns:
                msg += u"\n\nCanh bao ({} cap bi bo qua):".format(len(warns))
                msg += u"\n" + u"\n".join(warns[:6])
                if len(warns) > 6:
                    msg += u"\n... va {} loi khac".format(len(warns) - 6)
            MessageBox.Show(msg, "AutoJoin",
                            MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            MessageBox.Show(u"Loi thuc thi:\n" + str(ex), "Error",
                            MessageBoxButton.OK, MessageBoxImage.Error)
        finally:
            self.window.Close()

    def show(self):
        self.window.ShowDialog()

# ════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════
def main():
    try:
        sel_ids = list(uidoc.Selection.GetElementIds())
    except Exception:
        sel_ids = []
    AutoJoinWindow(sel_ids).show()

main()
OUT = "Success"