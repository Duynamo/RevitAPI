# -*- coding: utf-8 -*-
"""
Change Color & Transparency for Structural Elements in Active View
Tuân thủ quy trình xử lý lỗi: Validate -> Transaction -> Try Apply -> Commit / Rollback
"""

import clr
import traceback
import sys

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
import System.Windows
from System.Windows.Markup import XamlReader
from System.Windows import MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows.Media import Color as WpfColor, SolidColorBrush

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager


# ─── REVIT CONTEXT ─────────────────────────────────────────────────────────
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app   = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view  = doc.ActiveView


# ─── XAML GIAO DIỆN ────────────────────────────────────────────────────────
XAML_UI = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Structural Graphics Override" Height="520" Width="420"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        FontFamily="Segoe UI" Background="#F5F5F5">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Category Selection -->
        <GroupBox Grid.Row="0" Header="1. Select Categories" Margin="0,0,0,10" FontWeight="Bold">
            <StackPanel Margin="10" TextElement.FontWeight="Normal">
                <CheckBox Name="chkColumns" Content="Structural Columns" IsChecked="True" Margin="0,2"/>
                <CheckBox Name="chkFraming" Content="Structural Framing" IsChecked="True" Margin="0,2"/>
                <CheckBox Name="chkWalls" Content="Walls" IsChecked="True" Margin="0,2"/>
                <CheckBox Name="chkFloors" Content="Floors" IsChecked="True" Margin="0,2"/>
                <CheckBox Name="chkRoofs" Content="Roofs" Margin="0,2"/>
                <CheckBox Name="chkFoundations" Content="Structural Foundations" IsChecked="True" Margin="0,2"/>
                <CheckBox Name="chkGenericModels" Content="Generic Models" IsChecked="True" Margin="0,2"/>
            </StackPanel>
        </GroupBox>
        
        <!-- Color & Transparency Settings -->
        <GroupBox Grid.Row="1" Header="2. Graphics Settings" Margin="0,0,0,10" FontWeight="Bold">
            <Grid Margin="10" TextElement.FontWeight="Normal">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="50"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <!-- R -->
                <TextBlock Grid.Row="0" Grid.Column="0" Text="Red:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                <Slider Name="slRed" Grid.Row="0" Grid.Column="1" Minimum="0" Maximum="255" Value="100" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center" Margin="0,0,10,10"/>
                <TextBlock Name="txtRed" Grid.Row="0" Grid.Column="2" Text="{Binding ElementName=slRed, Path=Value}" VerticalAlignment="Center" Margin="0,0,0,10"/>
                
                <!-- G -->
                <TextBlock Grid.Row="1" Grid.Column="0" Text="Green:" VerticalAlignment="Center" Margin="0,0,10,10"/>
                <Slider Name="slGreen" Grid.Row="1" Grid.Column="1" Minimum="0" Maximum="255" Value="150" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center" Margin="0,0,10,10"/>
                <TextBlock Name="txtGreen" Grid.Row="1" Grid.Column="2" Text="{Binding ElementName=slGreen, Path=Value}" VerticalAlignment="Center" Margin="0,0,0,10"/>
                
                <!-- B -->
                <TextBlock Grid.Row="2" Grid.Column="0" Text="Blue:" VerticalAlignment="Center" Margin="0,0,10,15"/>
                <Slider Name="slBlue" Grid.Row="2" Grid.Column="1" Minimum="0" Maximum="255" Value="200" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center" Margin="0,0,10,15"/>
                <TextBlock Name="txtBlue" Grid.Row="2" Grid.Column="2" Text="{Binding ElementName=slBlue, Path=Value}" VerticalAlignment="Center" Margin="0,0,0,15"/>
                
                <!-- Transparency -->
                <TextBlock Grid.Row="3" Grid.Column="0" Text="Transp %:" VerticalAlignment="Center"/>
                <Slider Name="slTransp" Grid.Row="3" Grid.Column="1" Minimum="0" Maximum="100" Value="0" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBlock Name="txtTransp" Grid.Row="3" Grid.Column="2" Text="{Binding ElementName=slTransp, Path=Value}" VerticalAlignment="Center"/>
                
                <!-- Basic Colors -->
                <TextBlock Grid.Row="4" Grid.Column="0" Text="Presets:" VerticalAlignment="Center" Margin="0,15,10,0"/>
                <StackPanel Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,15,0,0">
                    <Button Name="btnC1" Width="22" Height="22" Background="#FFF44336" Margin="0,0,5,0" BorderThickness="1" BorderBrush="#DDDDDD" Cursor="Hand" ToolTip="Red"/>
                    <Button Name="btnC2" Width="22" Height="22" Background="#FFFF9800" Margin="0,0,5,0" BorderThickness="1" BorderBrush="#DDDDDD" Cursor="Hand" ToolTip="Orange"/>
                    <Button Name="btnC3" Width="22" Height="22" Background="#FFFFEB3B" Margin="0,0,5,0" BorderThickness="1" BorderBrush="#DDDDDD" Cursor="Hand" ToolTip="Yellow"/>
                    <Button Name="btnC4" Width="22" Height="22" Background="#FF4CAF50" Margin="0,0,5,0" BorderThickness="1" BorderBrush="#DDDDDD" Cursor="Hand" ToolTip="Green"/>
                    <Button Name="btnC5" Width="22" Height="22" Background="#FF2196F3" Margin="0,0,5,0" BorderThickness="1" BorderBrush="#DDDDDD" Cursor="Hand" ToolTip="Blue"/>
                    <Button Name="btnC6" Width="22" Height="22" Background="#FF9C27B0" Margin="0,0,5,0" BorderThickness="1" BorderBrush="#DDDDDD" Cursor="Hand" ToolTip="Purple"/>
                    <Button Name="btnC7" Width="22" Height="22" Background="#FF00BCD4" Margin="0,0,5,0" BorderThickness="1" BorderBrush="#DDDDDD" Cursor="Hand" ToolTip="Cyan"/>
                    <Button Name="btnC8" Width="22" Height="22" Background="#FF9E9E9E" Margin="0,0,5,0" BorderThickness="1" BorderBrush="#DDDDDD" Cursor="Hand" ToolTip="Gray"/>
                </StackPanel>
            </Grid>
        </GroupBox>
        
        <!-- Preview Color -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Top">
            <TextBlock Text="Preview Color:" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <Border BorderBrush="Black" BorderThickness="1" Width="100" Height="30">
                <Rectangle Name="rectColorPreview" Fill="#FF6496C8"/>
            </Border>
        </StackPanel>
        
        <!-- Action Buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnUndo" Content="Undo" Width="80" Height="35" Margin="0,0,10,0" Background="#2196F3" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
            <Button Name="btnReset" Content="Reset" Width="100" Height="35" Margin="0,0,10,0" Background="#FF9800" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
            <Button Name="btnApply" Content="Apply" Width="100" Height="35" Margin="0,0,10,0" Background="#4CAF50" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
            <Button Name="btnCancel" Content="Cancel" Width="80" Height="35" Background="#E0E0E0" BorderThickness="0"/>
        </StackPanel>
    </Grid>
</Window>
"""

# ─── HELPER ────────────────────────────────────────────────────────────────
def get_solid_fill_pattern(doc):
    """Tìm FillPattern dạng SolidFill trong file dự án."""
    collector = DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement)
    for fp in collector:
        if fp.GetFillPattern().IsSolidFill:
            return fp.Id
    return DB.ElementId.InvalidElementId

# ─── MAIN WINDOW CLASS ─────────────────────────────────────────────────────
class ColorStructureWindow(object):
    def __init__(self):
        self.undo_stack = []
        self.ui = XamlReader.Parse(XAML_UI)
        
        # Ánh xạ Controls
        self.chkColumns = self.ui.FindName("chkColumns")
        self.chkFraming = self.ui.FindName("chkFraming")
        self.chkWalls = self.ui.FindName("chkWalls")
        self.chkFloors = self.ui.FindName("chkFloors")
        self.chkRoofs = self.ui.FindName("chkRoofs")
        self.chkFoundations = self.ui.FindName("chkFoundations")
        self.chkGenericModels = self.ui.FindName("chkGenericModels")
        
        self.slRed = self.ui.FindName("slRed")
        self.slGreen = self.ui.FindName("slGreen")
        self.slBlue = self.ui.FindName("slBlue")
        self.slTransp = self.ui.FindName("slTransp")
        self.rectColorPreview = self.ui.FindName("rectColorPreview")
        
        self.btnApply = self.ui.FindName("btnApply")
        self.btnCancel = self.ui.FindName("btnCancel")
        self.btnReset = self.ui.FindName("btnReset")
        self.btnUndo = self.ui.FindName("btnUndo")
        
        # Đăng ký sự kiện
        self.slRed.ValueChanged += self.update_color_preview
        self.slGreen.ValueChanged += self.update_color_preview
        self.slBlue.ValueChanged += self.update_color_preview
        
        for i in range(1, 9):
            btn = self.ui.FindName("btnC" + str(i))
            if btn:
                btn.Click += self.preset_color_click

        self.btnApply.Click += self.btnApply_Click
        self.btnCancel.Click += self.btnCancel_Click
        self.btnReset.Click += self.btnReset_Click
        self.btnUndo.Click += self.btnUndo_Click
        
        # Chạy khởi tạo Preview lần đầu
        self.update_color_preview(None, None)

    def update_color_preview(self, sender, e):
        """Cập nhật Rectangle hiển thị màu sắc trên WPF."""
        r = int(self.slRed.Value)
        g = int(self.slGreen.Value)
        b = int(self.slBlue.Value)
        self.rectColorPreview.Fill = SolidColorBrush(WpfColor.FromRgb(r, g, b))

    def preset_color_click(self, sender, e):
        """Cập nhật thanh trượt khi chọn màu Basic (Preset)."""
        brush = sender.Background
        if brush:
            color = brush.Color
            self.slRed.Value = color.R
            self.slGreen.Value = color.G
            self.slBlue.Value = color.B

    def btnCancel_Click(self, sender, e):
        self.ui.Close()

    def get_selected_categories(self):
        selected_categories = []
        if self.chkColumns.IsChecked:
            selected_categories.append(DB.BuiltInCategory.OST_StructuralColumns)
        if self.chkFraming.IsChecked:
            selected_categories.append(DB.BuiltInCategory.OST_StructuralFraming)
        if self.chkWalls.IsChecked:
            selected_categories.append(DB.BuiltInCategory.OST_Walls)
        if self.chkFloors.IsChecked:
            selected_categories.append(DB.BuiltInCategory.OST_Floors)
        if self.chkRoofs.IsChecked:
            selected_categories.append(DB.BuiltInCategory.OST_Roofs)
        if self.chkFoundations.IsChecked:
            selected_categories.append(DB.BuiltInCategory.OST_StructuralFoundation)
        if self.chkGenericModels.IsChecked:
            selected_categories.append(DB.BuiltInCategory.OST_GenericModel)
        return selected_categories

    def btnReset_Click(self, sender, e):
        selected_categories = self.get_selected_categories()
        if not selected_categories:
            MessageBox.Show("Please select at least 1 Category!", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
            
        elements_to_modify = []
        for bic in selected_categories:
            collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(bic).WhereElementIsNotElementType()
            elements_to_modify.extend(list(collector))
            
        if not elements_to_modify:
            MessageBox.Show("No elements found for the selected categories in the current View.", "Information", MessageBoxButton.OK, MessageBoxImage.Information)
            return

        # Lưu trữ trạng thái màu trước khi thay đổi để phục vụ tính năng Undo
        current_state = [(elem.Id, view.GetElementOverrides(elem.Id)) for elem in elements_to_modify]
        self.undo_stack.append(current_state)

        TransactionManager.Instance.ForceCloseTransaction()
        t = DB.Transaction(doc, "Reset Structural Graphics")
        try:
            t.Start()
            ogs = DB.OverrideGraphicSettings() # Khởi tạo rỗng để reset về mặc định
            count = 0
            for elem in elements_to_modify:
                view.SetElementOverrides(elem.Id, ogs)
                count += 1
            t.Commit()
            MessageBox.Show("Success! Reset color and transparency to default for {} elements in the current View.".format(count), "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            if self.undo_stack: self.undo_stack.pop() # Xóa snapshot lỗi
            error_msg = "Cannot reset overrides.\nThe View might be locked by a View Template.\n\nError details: " + str(ex)
            MessageBox.Show(error_msg, "Execution Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def btnApply_Click(self, sender, e):
        # =========================================================
        # BƯỚC 2: Kiểm tra và thu thập Input (Validation)
        # =========================================================
        selected_categories = self.get_selected_categories()
            
        if not selected_categories:
            MessageBox.Show("Please select at least 1 Category!", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
            
        r = int(self.slRed.Value)
        g = int(self.slGreen.Value)
        b = int(self.slBlue.Value)
        transp = int(self.slTransp.Value)
        
        revit_color = DB.Color(r, g, b)
        
        elements_to_modify = []
        for bic in selected_categories:
            collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(bic).WhereElementIsNotElementType()
            elements_to_modify.extend(list(collector))
            
        if not elements_to_modify:
            MessageBox.Show("No elements found for the selected categories in the current View.", "Information", MessageBoxButton.OK, MessageBoxImage.Information)
            return

        # Lưu trữ trạng thái màu trước khi thay đổi để phục vụ tính năng Undo
        current_state = [(elem.Id, view.GetElementOverrides(elem.Id)) for elem in elements_to_modify]
        self.undo_stack.append(current_state)
        
        # =========================================================
        # BƯỚC 1 & 3: Khởi tạo Transaction & Chạy thử bộ lọc / Gán
        # =========================================================
        TransactionManager.Instance.ForceCloseTransaction() # Đóng Transaction ẩn của Dynamo
        t = DB.Transaction(doc, "Override Structural Graphics")
        try:
            t.Start()
            
            # Lấy Solid Fill Pattern
            solid_pattern_id = get_solid_fill_pattern(doc)
            if solid_pattern_id == DB.ElementId.InvalidElementId:
                raise Exception("Solid Fill Pattern not found in the project.")

            # Khởi tạo đối tượng ghi đè
            ogs = DB.OverrideGraphicSettings()
            
            # Gán màu và pattern bề mặt
            ogs.SetSurfaceForegroundPatternId(solid_pattern_id)
            ogs.SetSurfaceForegroundPatternColor(revit_color)
            
            # Gán màu và pattern mặt cắt
            ogs.SetCutForegroundPatternId(solid_pattern_id)
            ogs.SetCutForegroundPatternColor(revit_color)
            
            # Gán độ trong suốt
            ogs.SetSurfaceTransparency(transp)
            
            count = 0
            for elem in elements_to_modify:
                view.SetElementOverrides(elem.Id, ogs)
                count += 1
                    
            # =========================================================
            # BƯỚC 4: Nếu Không có lỗi -> Commit & Thông báo
            # =========================================================
            t.Commit()
            MessageBox.Show("Success! Applied color to {} structural elements in the current View.".format(count), "Success", MessageBoxButton.OK, MessageBoxImage.Information)
            
        except Exception as ex:
            # =========================================================
            # XỬ LÝ LỖI: Rollback Transaction và quay lại UI cho User
            # =========================================================
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            if self.undo_stack: self.undo_stack.pop() # Xóa snapshot lỗi
            
            error_msg = "Cannot apply overrides.\nThe View might be locked by a View Template.\n\nError details: " + str(ex)
            MessageBox.Show(error_msg, "Execution Error", MessageBoxButton.OK, MessageBoxImage.Error)
            # Không close UI để người dùng có thể tự gỡ lỗi (vd tắt View Template) rồi ấn Apply lại

    def btnUndo_Click(self, sender, e):
        if not self.undo_stack:
            MessageBox.Show("No actions to undo.", "Information", MessageBoxButton.OK, MessageBoxImage.Information)
            return
        
        last_state = self.undo_stack.pop()
        
        TransactionManager.Instance.ForceCloseTransaction()
        t = DB.Transaction(doc, "Undo Graphics Override")
        try:
            t.Start()
            count = 0
            for elem_id, ogs in last_state:
                view.SetElementOverrides(elem_id, ogs)
                count += 1
            t.Commit()
            MessageBox.Show("Undo successful! Reverted {} elements.".format(count), "Success", MessageBoxButton.OK, MessageBoxImage.Information)
        except Exception as ex:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            error_msg = "Cannot undo overrides.\nThe View might be locked by a View Template.\n\nError details: " + str(ex)
            MessageBox.Show(error_msg, "Execution Error", MessageBoxButton.OK, MessageBoxImage.Error)

# ─── RUN SCRIPT ────────────────────────────────────────────────────────────
def main():
    window = ColorStructureWindow()
    window.ui.ShowDialog()

main()
OUT = "Success"