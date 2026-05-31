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
        Title="Generic Models Selection Manager" 
        Width="480" Height="600" 
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

        <!-- Header: Nút Check/Uncheck & Search Box -->
        <Grid Grid.Row="0" Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button Name="btnCheckAll" Grid.Column="0" Content="Check All" Style="{StaticResource ModernBtn}" Background="#0078D7" Width="90" Margin="0,0,10,0"/>
            <Button Name="btnUncheckAll" Grid.Column="1" Content="Uncheck All" Style="{StaticResource ModernBtn}" Background="#555555" Width="90" Margin="0,0,10,0"/>
            <TextBox Name="txtSearch" Grid.Column="2" Padding="5,0" VerticalContentAlignment="Center" ToolTip="Tìm kiếm theo tên Generic Model..."/>
        </Grid>

        <!-- DataGrid hiển thị danh sách -->
        <DataGrid Name="dgModels" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" HeadersVisibility="Column" SelectionMode="Single" SelectionUnit="FullRow" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#E8E8E8" RowBackground="White" AlternatingRowBackground="#F9F9F9" BorderBrush="#DDDDDD" BorderThickness="1">
            <DataGrid.Columns>
                <!-- Sử dụng Template Column để CheckBox hoạt động bằng 1 Click -->
                <DataGridTemplateColumn Header="✓" Width="40">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding [IsChecked], UpdateSourceTrigger=PropertyChanged}" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Generic Model Name" Binding="{Binding [Name]}" Width="*" IsReadOnly="True"/>
                <DataGridTextColumn Header="Count" Binding="{Binding [Count]}" Width="60" IsReadOnly="True">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="HorizontalAlignment" Value="Center"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Footer: Các nút Action -->
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
def getGenericModelsInView(document, view_ref):
    """
    Thu thập tất cả các đối tượng Generic Model trong View hiện tại và gom nhóm theo Tên (Family - Type).
    Trả về Dictionary dạng: { "FamilyName - TypeName": [ElementId1, ElementId2, ...] }
    """
    collector = FilteredElementCollector(document, view_ref.Id)\
                .OfCategory(BuiltInCategory.OST_GenericModel)\
                .WhereElementIsNotElementType()
    
    gm_dict = {}
    for elem in collector:
        display_name = "Unknown Generic Model"
        try:
            # Ưu tiên lấy định dạng "Tên Family - Tên Type"
            if hasattr(elem, "Symbol") and hasattr(elem.Symbol, "Family"):
                display_name = "{} - {}".format(elem.Symbol.Family.Name, elem.Name)
            else:
                # Fallback cho các DirectShape hoặc In-Place Model
                display_name = getattr(elem, "Name", "Unknown Generic Model")
        except:
            pass
        
        if display_name not in gm_dict:
            gm_dict[display_name] = []
        gm_dict[display_name].append(elem.Id)
        
    return gm_dict
#endregion

#region ___WPF UI Class
class GenericModelManagerWindow(object):
    def __init__(self, doc_ref, uidoc_ref, view_ref):
        self.doc = doc_ref
        self.uidoc = uidoc_ref
        self.view = view_ref
        self.gm_dict = {}
        
        self.window = XamlReader.Parse(WPF_XAML)
        self._find_controls()
        self._bind_events()
        
        # Setup DataTable
        self.dt_data = DataTable("GenericModels")
        self.dt_data.Columns.Add("IsChecked", System.Boolean)
        self.dt_data.Columns.Add("Name", System.String)
        self.dt_data.Columns.Add("Count", System.Int32)
        
        self.dgModels.ItemsSource = self.dt_data.DefaultView
        self.LoadData()

    def _find(self, name):
        return self.window.FindName(name)

    def _find_controls(self):
        self.btnCheckAll = self._find("btnCheckAll")
        self.btnUncheckAll = self._find("btnUncheckAll")
        self.txtSearch = self._find("txtSearch")
        self.dgModels = self._find("dgModels")
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

    def LoadData(self):
        self.gm_dict = getGenericModelsInView(self.doc, self.view)
        sorted_names = sorted(self.gm_dict.keys())
        
        for name in sorted_names:
            row = self.dt_data.NewRow()
            row["IsChecked"] = False
            row["Name"] = name
            row["Count"] = len(self.gm_dict[name])
            self.dt_data.Rows.Add(row)

    def OnSearchTextChanged(self, sender, e):
        search_text = self.txtSearch.Text.strip()
        if search_text:
            self.dt_data.DefaultView.RowFilter = "Name LIKE '%{}%'".format(search_text)
        else:
            self.dt_data.DefaultView.RowFilter = ""

    def OnCheckAll(self, sender, e):
        for row_view in self.dt_data.DefaultView:
            row_view.Row["IsChecked"] = True
        self.dgModels.Items.Refresh()

    def OnUncheckAll(self, sender, e):
        for row_view in self.dt_data.DefaultView:
            row_view.Row["IsChecked"] = False
        self.dgModels.Items.Refresh()

    def GetSelectedElementIds(self):
        """Gom toàn bộ ElementId của các đối tượng thuộc các tên được tick chọn"""
        ids = List[ElementId]()
        for row in self.dt_data.Rows:
            if row["IsChecked"]:
                name = row["Name"]
                if name in self.gm_dict:
                    for eid in self.gm_dict[name]:
                        ids.Add(eid)
        return ids

    def OnSelect(self, sender, e):
        try:
            element_ids = self.GetSelectedElementIds()
            if element_ids.Count == 0:
                MessageBox.Show("Vui lòng tick chọn ít nhất một Model để Select!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                return

            self.uidoc.Selection.SetElementIds(element_ids)
            self.window.Close()
        except Exception as ex:
            MessageBox.Show("Lỗi trong quá trình Select:\n" + str(ex), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def OnIsolate(self, sender, e):
        try:
            element_ids = self.GetSelectedElementIds()
            if element_ids.Count == 0:
                MessageBox.Show("Vui lòng tick chọn ít nhất một Model để Isolate!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
                return

            # Isolate Element yêu cầu View thay đổi -> cần Transaction
            t = Transaction(self.doc, "Isolate Generic Models Temporary")
            t.Start()
            try:
                self.view.IsolateElementsTemporary(element_ids)
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
    app = GenericModelManagerWindow(doc, uidoc, view)
    app.show()

# Execute directly in Dynamo Environment
main()
OUT = "Success"