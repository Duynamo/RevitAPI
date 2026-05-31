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

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = uiapp.ActiveUIDocument

view = doc.ActiveView
#endregion

#region XAML UI Definition
WPF_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Category Visibility &amp; Selection Manager" 
        Width="450" Height="600" 
        WindowStartupLocation="CenterScreen" 
        Background="#F5F6F8" FontFamily="Segoe UI" FontSize="13"
        Topmost="True">
    <Window.Resources>
        <Style x:Key="ModernBtn" TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button Name="btnCheckAll" Grid.Column="0" Content="Check All" Style="{StaticResource ModernBtn}" Background="#0078D7" Width="90" Margin="0,0,10,0"/>
            <Button Name="btnUncheckAll" Grid.Column="1" Content="Uncheck All" Style="{StaticResource ModernBtn}" Background="#555555" Width="90" Margin="0,0,10,0"/>
            <TextBox Name="txtSearch" Grid.Column="2" Padding="5,0" VerticalContentAlignment="Center" ToolTip="Tìm kiếm Category..."/>
        </Grid>

        <DataGrid Name="dgCategories" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" HeadersVisibility="Column" SelectionMode="Single" SelectionUnit="FullRow" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#E8E8E8" RowBackground="White" AlternatingRowBackground="#F9F9F9" BorderBrush="#DDDDDD" BorderThickness="1">
            <DataGrid.Columns>
                <DataGridTemplateColumn Header="✓" Width="40">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding [IsChecked], UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Category Name" Binding="{Binding [Name]}" Width="*" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>

        <Grid Grid.Row="2" Margin="0,15,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            
            <Button Name="btnSelect" Grid.Column="1" Content="Select" Width="90" Height="35" Margin="0,0,10,0" Style="{StaticResource ModernBtn}" Background="#0078D7"/>
            <Button Name="btnIsolate" Grid.Column="2" Content="Isolate" Width="90" Height="35" Margin="0,0,10,0" Style="{StaticResource ModernBtn}" Background="#107C10"/>
            <Button Name="btnCancel" Grid.Column="3" Content="Cancel" Width="90" Height="35" Style="{StaticResource ModernBtn}" Background="#A80000"/>
        </Grid>
    </Grid>
</Window>
"""
#endregion

#region ___Logic Backend
def getAllModelCategories(document):
    """
    Lấy tất cả các Model Categories hợp lệ trong dự án và sắp xếp theo bảng chữ cái.
    """
    categories = document.Settings.Categories
    model_cats = []
    for cat in categories:
        try:
            # Lọc chỉ lấy CategoryType.Model và có thể gán Subcategory (loại bỏ các mục ảo)
            if cat.CategoryType == CategoryType.Model and cat.CanAddSubcategory:
                model_cats.append(cat)
        except:
            pass
    
    # Sắp xếp theo tên từ điển
    model_cats.sort(key=lambda c: c.Name)
    return model_cats
#endregion

#region ___WPF UI Class
class CategoryManagerWindow(object):
    def __init__(self, doc_ref, uidoc_ref, view_ref):
        self.doc = doc_ref
        self.uidoc = uidoc_ref
        self.view = view_ref
        
        self.window = XamlReader.Parse(WPF_XAML)
        self._find_controls()
        self._bind_events()
        
        # Setup DataTable
        self.dt_cats = DataTable("Categories")
        self.dt_cats.Columns.Add("IsChecked", System.Boolean)
        self.dt_cats.Columns.Add("Name", System.String)
        self.dt_cats.Columns.Add("Id", System.Object) # Lưu ElementId
        
        self.dgCategories.ItemsSource = self.dt_cats.DefaultView
        self.LoadCategories()

    def _find(self, name):
        return self.window.FindName(name)

    def _find_controls(self):
        self.btnCheckAll = self._find("btnCheckAll")
        self.btnUncheckAll = self._find("btnUncheckAll")
        self.txtSearch = self._find("txtSearch")
        self.dgCategories = self._find("dgCategories")
        self.btnSelect = self._find("btnSelect")
        self.btnIsolate = self._find("btnIsolate")
        self.btnCancel = self._find("btnCancel")

    def _bind_events(self):
        self.btnCheckAll.Click += self.OnCheckAll
        self.btnUncheckAll.Click += self.OnUncheckAll
        self.txtSearch.TextChanged += self.OnSearchTextChanged
        self.btnSelect.Click += self.OnSelect
        self.btnIsolate.Click += self.OnIsolate
        self.btnCancel.Click += self.OnCancel

    def LoadCategories(self):
        """Tải dữ liệu categories vào DataTable"""
        categories = getAllModelCategories(self.doc)
        for cat in categories:
            row = self.dt_cats.NewRow()
            row["IsChecked"] = False
            row["Name"] = cat.Name
            row["Id"] = cat.Id
            self.dt_cats.Rows.Add(row)

    def OnSearchTextChanged(self, sender, e):
        search_text = self.txtSearch.Text.strip()
        if search_text:
            # Lọc theo trường Name, hỗ trợ không phân biệt chữ hoa/thường
            self.dt_cats.DefaultView.RowFilter = "Name LIKE '%{}%'".format(search_text)
        else:
            self.dt_cats.DefaultView.RowFilter = ""

    def OnCheckAll(self, sender, e):
        for row_view in self.dt_cats.DefaultView:
            row_view.Row["IsChecked"] = True
        self.dgCategories.Items.Refresh()

    def OnUncheckAll(self, sender, e):
        for row_view in self.dt_cats.DefaultView:
            row_view.Row["IsChecked"] = False
        self.dgCategories.Items.Refresh()

    def GetSelectedCategoryIds(self):
        """Trả về danh sách List[ElementId] của các category được chọn."""
        cat_ids = List[ElementId]()
        for row in self.dt_cats.Rows:
            if row["IsChecked"]:
                cat_ids.Add(row["Id"])
        return cat_ids

    def OnSelect(self, sender, e):
        try:
            cat_ids = self.GetSelectedCategoryIds()
            if cat_ids.Count == 0:
                MessageBox.Show("Vui lòng tick chọn ít nhất một Category để Select!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                return

            # Tạo bộ lọc ElementMulticategoryFilter
            cat_filter = ElementMulticategoryFilter(cat_ids)
            collector = FilteredElementCollector(self.doc, self.view.Id)
            collector.WherePasses(cat_filter).WhereElementIsNotElementType()
            
            element_ids = collector.ToElementIds()

            if element_ids.Count == 0:
                MessageBox.Show("Không tìm thấy đối tượng nào thuộc các Categories đã chọn trong View hiện hành!", "Info", MessageBoxButton.OK, MessageBoxImage.Information)
            else:
                self.uidoc.Selection.SetElementIds(element_ids)
                self.window.Close()
        except Exception as ex:
            MessageBox.Show("Lỗi trong quá trình Select:\n" + str(ex), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def OnIsolate(self, sender, e):
        try:
            cat_ids = self.GetSelectedCategoryIds()
            if cat_ids.Count == 0:
                MessageBox.Show("Vui lòng tick chọn ít nhất một Category để Isolate!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                return

            # Isolate yêu cầu thay đổi View nên cần thực thi Transaction
            t = Transaction(self.doc, "Isolate Categories Temporary")
            t.Start()
            try:
                self.view.IsolateCategoriesTemporary(cat_ids)
                t.Commit()
                self.window.Close()
            except Exception as ex:
                t.RollBack()
                MessageBox.Show("Không thể Isolate:\n" + str(ex), "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        except Exception as ex:
            MessageBox.Show("Lỗi trong quá trình khởi tạo Isolate:\n" + str(ex), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def OnCancel(self, sender, e):
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
#endregion


def main():
    app = CategoryManagerWindow(doc, uidoc, view)
    app.show()

# Execute directly in Dynamo Environment
main()
OUT = "Success"