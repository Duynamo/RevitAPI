"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys
import System

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

clr.AddReference("System")
clr.AddReference("System.Data")
from System.Collections.Generic import List
from System.Data import DataTable

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
from System.Windows.Markup import XamlReader
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage

# Khởi tạo các biến Revit cơ bản cho môi trường Dynamo (IronPython)
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc  = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app  = uiapp.Application
uidoc = uiapp.ActiveUIDocument

view = doc.ActiveView
#endregion

#region XAML UI Definition
WPF_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Filter Elements by Parameter"
        Width="480" Height="430"
        WindowStartupLocation="CenterScreen"
        Background="#F5F6F8" FontFamily="Segoe UI" FontSize="13"
        ResizeMode="NoResize"
        Topmost="True">
    <Window.Resources>
        <!-- Modern Button Style -->
        <Style x:Key="ModernBtn" TargetType="Button">
            <Setter Property="Foreground"       Value="White"/>
            <Setter Property="BorderThickness"  Value="0"/>
            <Setter Property="Padding"          Value="10,5"/>
            <Setter Property="Cursor"           Value="Hand"/>
            <Setter Property="FontWeight"       Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Opacity" Value="0.85"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Opacity" Value="0.70"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Label Style -->
        <Style x:Key="FieldLabel" TargetType="TextBlock">
            <Setter Property="FontWeight"   Value="SemiBold"/>
            <Setter Property="Foreground"   Value="#333333"/>
            <Setter Property="Margin"       Value="0,0,0,4"/>
        </Style>

        <!-- ComboBox Style -->
        <Style TargetType="ComboBox">
            <Setter Property="Height"           Value="30"/>
            <Setter Property="Padding"          Value="6,4"/>
            <Setter Property="Background"       Value="White"/>
            <Setter Property="BorderBrush"      Value="#CCCCCC"/>
            <Setter Property="BorderThickness"  Value="1"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <!-- Title -->
            <RowDefinition Height="Auto"/>
            <!-- Separator -->
            <RowDefinition Height="10"/>
            <!-- Category -->
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <!-- Gap -->
            <RowDefinition Height="12"/>
            <!-- Parameter Group -->
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <!-- Gap -->
            <RowDefinition Height="12"/>
            <!-- Parameter Name -->
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <!-- Gap -->
            <RowDefinition Height="12"/>
            <!-- Parameter Value -->
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <!-- Spacer -->
            <RowDefinition Height="*"/>
            <!-- Buttons + Footer -->
            <RowDefinition Height="Auto"/>
            <!-- Footer gap -->
            <RowDefinition Height="8"/>
        </Grid.RowDefinitions>

        <!-- ====== Title ====== -->
        <TextBlock Grid.Row="0"
                   Text="Filter Elements by Parameter"
                   FontSize="16" FontWeight="Bold"
                   Foreground="#0078D7"/>

        <!-- Separator Line -->
        <Rectangle Grid.Row="1" Height="1" Fill="#DDDDDD" Margin="0,4,0,0"/>

        <!-- ====== Category ====== -->
        <TextBlock Grid.Row="2" Text="Select Category" Style="{StaticResource FieldLabel}"/>
        <ComboBox  Name="cmbCategory" Grid.Row="3"/>

        <!-- ====== Parameter Group ====== -->
        <TextBlock Grid.Row="5" Text="Select Parameter Group" Style="{StaticResource FieldLabel}"/>
        <ComboBox  Name="cmbParamGroup" Grid.Row="6" IsEnabled="False"/>

        <!-- ====== Parameter Name ====== -->
        <TextBlock Grid.Row="8" Text="Select Parameter Name" Style="{StaticResource FieldLabel}"/>
        <ComboBox  Name="cmbParamName" Grid.Row="9" IsEnabled="False"/>

        <!-- ====== Parameter Value ====== -->
        <TextBlock Grid.Row="11" Text="Select Parameter Value" Style="{StaticResource FieldLabel}"/>
        <ComboBox  Name="cmbParamValue" Grid.Row="12" IsEnabled="False"/>

        <!-- ====== Action Buttons + Watermark ====== -->
        <Grid Grid.Row="14" Margin="0,12,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <!-- Watermark V2D at bottom-left -->
            <TextBlock Grid.Column="0"
                       Text="&#169; V2D"
                       FontSize="11" FontWeight="Bold"
                       Foreground="#AAAAAA"
                       VerticalAlignment="Center"
                       Margin="0,0,8,0"
                       ToolTip="Copyright by vudinhduybm@gmail.com"/>

            <Button Name="btnIsolate" Grid.Column="2"
                    Content="Isolate" Width="100" Height="35"
                    Style="{StaticResource ModernBtn}" Background="#107C10"/>

            <Button Name="btnHide"    Grid.Column="4"
                    Content="Hide"    Width="100" Height="35"
                    Style="{StaticResource ModernBtn}" Background="#FF8C00"/>

            <Button Name="btnCancel"  Grid.Column="6"
                    Content="Cancel"  Width="100" Height="35"
                    Style="{StaticResource ModernBtn}" Background="#A80000"/>
        </Grid>
    </Grid>
</Window>
"""
#endregion

#region ___Logic Backend
def getModelCategoriesInView(document, view_ref):
    """
    Lấy danh sách các Category thực sự có đối tượng trong View hiện tại.
    Trả về list tuple (category_name, BuiltInCategory) đã sắp xếp theo tên.
    """
    collector = FilteredElementCollector(document, view_ref.Id)\
                .WhereElementIsNotElementType()

    cat_set = {}
    for elem in collector:
        try:
            cat = elem.Category
            if cat is None:
                continue
            if cat.CategoryType != CategoryType.Model:
                continue
            if cat.Name not in cat_set:
                cat_set[cat.Name] = cat.Id
        except:
            pass

    result = sorted(cat_set.items(), key=lambda x: x[0])
    return result  # list of (name, ElementId)


def getParameterGroupsForCategory(document, view_ref, category_id):
    """
    Lấy tất cả các Parameter Group tồn tại trên các element thuộc category_id trong View.
    Trả về list string tên nhóm đã sắp xếp, và dict {group_name: group_enum_value}
    """
    cat_filter = ElementCategoryFilter(category_id)
    collector = FilteredElementCollector(document, view_ref.Id)\
                .WherePasses(cat_filter)\
                .WhereElementIsNotElementType()

    group_dict = {}  # { group_label_string: internal_group_value }
    for elem in collector:
        try:
            params = elem.Parameters
            for p in params:
                try:
                    # Lấy tên nhóm hiển thị
                    group_name = p.Definition.ParameterGroup.ToString()
                    # Thử lấy label đẹp hơn qua LabelUtils nếu có
                    try:
                        group_name = LabelUtils.GetLabelFor(p.Definition.ParameterGroup)
                    except:
                        pass
                    if group_name not in group_dict:
                        group_dict[group_name] = p.Definition.ParameterGroup
                except:
                    pass
        except:
            pass

    return group_dict  # { display_name: group_enum }


def getParameterNamesForGroup(document, view_ref, category_id, param_group_enum):
    """
    Lấy tất cả tên Parameter thuộc param_group_enum trên các element của category_id trong View.
    Trả về list string tên parameter đã sắp xếp.
    """
    cat_filter = ElementCategoryFilter(category_id)
    collector = FilteredElementCollector(document, view_ref.Id)\
                .WherePasses(cat_filter)\
                .WhereElementIsNotElementType()

    param_names = set()
    for elem in collector:
        try:
            params = elem.Parameters
            for p in params:
                try:
                    if p.Definition.ParameterGroup == param_group_enum:
                        param_names.add(p.Definition.Name)
                except:
                    pass
        except:
            pass

    return sorted(list(param_names))


def getParameterValuesForName(document, view_ref, category_id, param_name):
    """
    Lấy tất cả các giá trị (value string) của parameter param_name
    trên các element của category_id trong View.
    Trả về dict { value_string: [ElementId, ...] }
    """
    cat_filter = ElementCategoryFilter(category_id)
    collector = FilteredElementCollector(document, view_ref.Id)\
                .WherePasses(cat_filter)\
                .WhereElementIsNotElementType()

    value_dict = {}  # { value_string: [ElementId] }
    for elem in collector:
        try:
            param = elem.LookupParameter(param_name)
            if param is None:
                continue
            # Lấy giá trị string
            if param.StorageType == StorageType.String:
                val = param.AsString() or ""
            elif param.StorageType == StorageType.Integer:
                val = str(param.AsInteger())
            elif param.StorageType == StorageType.Double:
                val = str(round(param.AsDouble(), 4))
            elif param.StorageType == StorageType.ElementId:
                eid = param.AsElementId()
                val = str(eid.IntegerValue)
            else:
                val = "N/A"

            if val not in value_dict:
                value_dict[val] = []
            value_dict[val].append(elem.Id)
        except:
            pass

    return value_dict
#endregion

#region ___WPF UI Class
class FilterByParameterWindow(object):
    def __init__(self, doc_ref, uidoc_ref, view_ref):
        self.doc  = doc_ref
        self.uidoc = uidoc_ref
        self.view = view_ref

        # Internal state
        self._categories    = []   # list of (name, ElementId_of_category)
        self._group_dict    = {}   # { display_name: group_enum }
        self._param_names   = []   # list of string
        self._value_dict    = {}   # { value_string: [ElementId] }
        self._cat_id        = None # ElementId of selected category

        self.window = XamlReader.Parse(WPF_XAML)
        self._find_controls()
        self._bind_events()
        self._load_categories()

    # ── helpers ──────────────────────────────────────────────────────────────
    def _find(self, name):
        return self.window.FindName(name)

    def _find_controls(self):
        self.cmbCategory   = self._find("cmbCategory")
        self.cmbParamGroup = self._find("cmbParamGroup")
        self.cmbParamName  = self._find("cmbParamName")
        self.cmbParamValue = self._find("cmbParamValue")
        self.btnIsolate    = self._find("btnIsolate")
        self.btnHide       = self._find("btnHide")
        self.btnCancel     = self._find("btnCancel")

    def _bind_events(self):
        self.cmbCategory.SelectionChanged   += self.OnCategoryChanged
        self.cmbParamGroup.SelectionChanged += self.OnParamGroupChanged
        self.cmbParamName.SelectionChanged  += self.OnParamNameChanged
        self.btnIsolate.Click += self.OnIsolate
        self.btnHide.Click    += self.OnHide
        self.btnCancel.Click  += self.OnCancel

    # ── Load data helpers ────────────────────────────────────────────────────
    def _clear_combo(self, combo, placeholder, enable=False):
        """Xóa item, thêm placeholder và set trạng thái enable."""
        combo.Items.Clear()
        combo.Items.Add(placeholder)
        combo.SelectedIndex = 0
        combo.IsEnabled = enable

    def _load_categories(self):
        self._categories = getModelCategoriesInView(self.doc, self.view)
        self.cmbCategory.Items.Clear()
        self.cmbCategory.Items.Add("-- Select Category --")
        for cat_name, _ in self._categories:
            self.cmbCategory.Items.Add(cat_name)
        self.cmbCategory.SelectedIndex = 0

        # Reset các combo phụ
        self._clear_combo(self.cmbParamGroup, "-- Select Parameter Group --", False)
        self._clear_combo(self.cmbParamName,  "-- Select Parameter Name --",  False)
        self._clear_combo(self.cmbParamValue, "-- Select Parameter Value --", False)

    # ── Event handlers ───────────────────────────────────────────────────────
    def OnCategoryChanged(self, sender, e):
        idx = self.cmbCategory.SelectedIndex
        if idx <= 0:
            # Reset tất cả
            self._cat_id = None
            self._clear_combo(self.cmbParamGroup, "-- Select Parameter Group --", False)
            self._clear_combo(self.cmbParamName,  "-- Select Parameter Name --",  False)
            self._clear_combo(self.cmbParamValue, "-- Select Parameter Value --", False)
            return

        cat_name, cat_elem_id = self._categories[idx - 1]
        self._cat_id = cat_elem_id

        # Load Parameter Groups
        self._group_dict = getParameterGroupsForCategory(self.doc, self.view, self._cat_id)
        sorted_groups = sorted(self._group_dict.keys())

        self.cmbParamGroup.Items.Clear()
        self.cmbParamGroup.Items.Add("-- Select Parameter Group --")
        for g in sorted_groups:
            self.cmbParamGroup.Items.Add(g)
        self.cmbParamGroup.SelectedIndex = 0
        self.cmbParamGroup.IsEnabled = True

        # Reset downstream
        self._clear_combo(self.cmbParamName,  "-- Select Parameter Name --",  False)
        self._clear_combo(self.cmbParamValue, "-- Select Parameter Value --", False)

    def OnParamGroupChanged(self, sender, e):
        idx = self.cmbParamGroup.SelectedIndex
        if idx <= 0:
            self._clear_combo(self.cmbParamName,  "-- Select Parameter Name --",  False)
            self._clear_combo(self.cmbParamValue, "-- Select Parameter Value --", False)
            return

        group_name = self.cmbParamGroup.SelectedItem
        group_enum = self._group_dict.get(group_name, None)
        if group_enum is None:
            return

        # Load Parameter Names
        self._param_names = getParameterNamesForGroup(
            self.doc, self.view, self._cat_id, group_enum
        )

        self.cmbParamName.Items.Clear()
        self.cmbParamName.Items.Add("-- Select Parameter Name --")
        for n in self._param_names:
            self.cmbParamName.Items.Add(n)
        self.cmbParamName.SelectedIndex = 0
        self.cmbParamName.IsEnabled = True

        # Reset downstream
        self._clear_combo(self.cmbParamValue, "-- Select Parameter Value --", False)

    def OnParamNameChanged(self, sender, e):
        idx = self.cmbParamName.SelectedIndex
        if idx <= 0:
            self._clear_combo(self.cmbParamValue, "-- Select Parameter Value --", False)
            return

        param_name = self.cmbParamName.SelectedItem

        # Load Parameter Values
        self._value_dict = getParameterValuesForName(
            self.doc, self.view, self._cat_id, param_name
        )

        sorted_values = sorted(self._value_dict.keys())

        self.cmbParamValue.Items.Clear()
        self.cmbParamValue.Items.Add("-- Select Parameter Value --")
        for v in sorted_values:
            self.cmbParamValue.Items.Add(v)
        self.cmbParamValue.SelectedIndex = 0
        self.cmbParamValue.IsEnabled = True

    # ── Collect target element ids ───────────────────────────────────────────
    def _get_target_element_ids(self):
        """
        Trả về List[ElementId] ứng với giá trị parameter được chọn.
        Trả về None nếu chưa chọn đủ thông tin.
        """
        if self.cmbParamValue.SelectedIndex <= 0:
            return None

        value_str = self.cmbParamValue.SelectedItem
        raw_ids   = self._value_dict.get(value_str, [])
        ids = List[ElementId]()
        for eid in raw_ids:
            ids.Add(eid)
        return ids

    # ── Button handlers ──────────────────────────────────────────────────────
    def OnIsolate(self, sender, e):
        try:
            ids = self._get_target_element_ids()
            if ids is None or ids.Count == 0:
                MessageBox.Show(
                    "Vui lòng chọn đầy đủ Category, Parameter Group, Parameter Name và Parameter Value!",
                    "Warning", MessageBoxButton.OK, MessageBoxImage.Warning
                )
                return

            t = Transaction(self.doc, "Isolate Elements by Parameter")
            t.Start()
            try:
                self.view.IsolateElementsTemporary(ids)
                t.Commit()
                self.window.Close()
            except Exception as ex:
                t.RollBack()
                MessageBox.Show(
                    "Không thể Isolate:\n" + str(ex),
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error
                )
        except Exception as ex:
            MessageBox.Show(
                "Lỗi khi thực hiện Isolate:\n" + str(ex),
                "Error", MessageBoxButton.OK, MessageBoxImage.Error
            )

    def OnHide(self, sender, e):
        try:
            ids = self._get_target_element_ids()
            if ids is None or ids.Count == 0:
                MessageBox.Show(
                    "Vui lòng chọn đầy đủ Category, Parameter Group, Parameter Name và Parameter Value!",
                    "Warning", MessageBoxButton.OK, MessageBoxImage.Warning
                )
                return

            t = Transaction(self.doc, "Hide Elements by Parameter")
            t.Start()
            try:
                self.view.HideElements(ids)
                t.Commit()
                self.window.Close()
            except Exception as ex:
                t.RollBack()
                MessageBox.Show(
                    "Không thể Hide:\n" + str(ex),
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error
                )
        except Exception as ex:
            MessageBox.Show(
                "Lỗi khi thực hiện Hide:\n" + str(ex),
                "Error", MessageBoxButton.OK, MessageBoxImage.Error
            )

    def OnCancel(self, sender, e):
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
#endregion


def main():
    app = FilterByParameterWindow(doc, uidoc, view)
    app.show()

# Execute directly in Dynamo Environment
main()
OUT = "Success"