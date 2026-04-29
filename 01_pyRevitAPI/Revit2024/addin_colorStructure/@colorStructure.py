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
        Title="Structural Graphics Override" Height="480" Width="380"
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
                <Slider Name="slTransp" Grid.Row="3" Grid.Column="1" Minimum="0" Maximum="100" Value="30" TickFrequency="1" IsSnapToTickEnabled="True" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBlock Name="txtTransp" Grid.Row="3" Grid.Column="2" Text="{Binding ElementName=slTransp, Path=Value}" VerticalAlignment="Center"/>
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
            <Button Name="btnApply" Content="Apply / Gán màu" Width="120" Height="35" Margin="0,0,10,0" Background="#4CAF50" Foreground="White" FontWeight="Bold" BorderThickness="0"/>
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
        self.ui = XamlReader.Parse(XAML_UI)
        
        # Ánh xạ Controls
        self.chkColumns = self.ui.FindName("chkColumns")
        self.chkFraming = self.ui.FindName("chkFraming")
        self.chkWalls = self.ui.FindName("chkWalls")
        self.chkFloors = self.ui.FindName("chkFloors")
        self.chkRoofs = self.ui.FindName("chkRoofs")
        self.chkFoundations = self.ui.FindName("chkFoundations")
        
        self.slRed = self.ui.FindName("slRed")
        self.slGreen = self.ui.FindName("slGreen")
        self.slBlue = self.ui.FindName("slBlue")
        self.slTransp = self.ui.FindName("slTransp")
        self.rectColorPreview = self.ui.FindName("rectColorPreview")
        
        self.btnApply = self.ui.FindName("btnApply")
        self.btnCancel = self.ui.FindName("btnCancel")
        
        # Đăng ký sự kiện
        self.slRed.ValueChanged += self.update_color_preview
        self.slGreen.ValueChanged += self.update_color_preview
        self.slBlue.ValueChanged += self.update_color_preview
        
        self.btnApply.Click += self.btnApply_Click
        self.btnCancel.Click += self.btnCancel_Click
        
        # Chạy khởi tạo Preview lần đầu
        self.update_color_preview(None, None)

    def update_color_preview(self, sender, e):
        """Cập nhật Rectangle hiển thị màu sắc trên WPF."""
        r = int(self.slRed.Value)
        g = int(self.slGreen.Value)
        b = int(self.slBlue.Value)
        self.rectColorPreview.Fill = SolidColorBrush(WpfColor.FromRgb(r, g, b))

    def btnCancel_Click(self, sender, e):
        self.ui.Close()

    def btnApply_Click(self, sender, e):
        # =========================================================
        # BƯỚC 2: Kiểm tra và thu thập Input (Validation)
        # =========================================================
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
            
        if not selected_categories:
            MessageBox.Show(u"Vui lòng chọn ít nhất 1 Category!", u"Lỗi nhập liệu", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
            
        r = int(self.slRed.Value)
        g = int(self.slGreen.Value)
        b = int(self.slBlue.Value)
        transp = int(self.slTransp.Value)
        
        revit_color = DB.Color(r, g, b)
        
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
                raise Exception(u"Không tìm thấy Solid Fill Pattern trong file dự án.")

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
            for bic in selected_categories:
                # Chỉ lấy các Element thực tế đang nằm trên ActiveView
                collector = DB.FilteredElementCollector(doc, view.Id).OfCategory(bic).WhereElementIsNotElementType()
                for elem in collector:
                    view.SetElementOverrides(elem.Id, ogs)
                    count += 1
                    
            # =========================================================
            # BƯỚC 4: Nếu Không có lỗi -> Commit & Thông báo
            # =========================================================
            t.Commit()
            MessageBox.Show(u"Hoàn thành! Đã gán màu cho {} đối tượng thuộc hệ kết cấu trong View hiện tại.".format(count), u"Thành công", MessageBoxButton.OK, MessageBoxImage.Information)
            self.ui.Close() # Chỉ tắt form khi đã hoàn thành trót lọt
            
        except Exception as ex:
            # =========================================================
            # XỬ LÝ LỖI: Rollback Transaction và quay lại UI cho User
            # =========================================================
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
            
            error_msg = u"Không thể áp dụng Override.\nCó thể View đang bị khóa bởi View Template.\n\nChi tiết lỗi: " + str(ex)
            MessageBox.Show(error_msg, u"Lỗi thực thi", MessageBoxButton.OK, MessageBoxImage.Error)
            # Không close UI để người dùng có thể tự gỡ lỗi (vd tắt View Template) rồi ấn Apply lại

# ─── RUN SCRIPT ────────────────────────────────────────────────────────────
def main():
    window = ColorStructureWindow()
    window.ui.ShowDialog()

main()
OUT = "Success"