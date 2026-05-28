# Cơ sở dữ liệu: pyRevit / Dynamo IronPython Scripts & UI Framework

Tài liệu này chứa các quy chuẩn viết mã (Coding Standard) và thiết kế UI (User Interface) đã hoạt động cực kỳ ổn định trên môi trường Revit / Dynamo (IronPython 2.7).
Khi viết mã Python mới yêu cầu giao diện WPF và COM API, **HÃY ƯU TIÊN** áp dụng định dạng và các kỹ thuật từ file mẫu dưới đây.

## Quy chuẩn cốt lõi:
1. **XAML UI:** Chuỗi XAML phải được khai báo dạng Unicode (`u"""..."""`) và nạp bằng `XamlReader.Parse()`.
2. **Data Binding WPF:** Bắt buộc dùng `System.Data.DataTable` để đổ dữ liệu vào `ListView`/`DataGrid` vì IronPython object thuần không hỗ trợ Binding .NET.
3. **COM API (AutoCAD, Excel, v.v.):** Sử dụng mảng kiểu `Array[Double]` khi truyền dữ liệu tọa độ 3D và dùng `clr.Reference` cho các tham số hàm kiểu `out`.
4. **UI Refresh:** Cập nhật giao diện mượt mà (không đóng băng) bằng `Dispatcher.Invoke` với một `empty_action`.
5. **Entry point:** Khởi tạo qua hàm `main()` và trả về kết quả qua biến `OUT` (đối với Dynamo).

---

## Context: File Mẫu `combineLayout_Template.py`

```python
import os
import glob
import time

# Load System & .NET Data libraries
import clr
clr.AddReference("System")
clr.AddReference("System.Data")
import System
from System.Collections.Generic import List
from System import Array, Double
from System.Data import DataTable

# Load WPF/XAML libraries
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
from System.Windows.Markup import XamlReader
from System.Windows import Window, MessageBox, MessageBoxButton, MessageBoxImage
from System.Windows.Threading import DispatcherPriority

# Connect to COM API (AutoCAD/Excel...)
import System.Runtime.InteropServices

#region XAML UI Definition
# Lưu ý: Luôn dùng tiền tố `u` cho chuỗi XAML để tránh lỗi Parse trong IronPython
WPF_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="App Title Here" Height="500" Width="650" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        
        <ListView Name="lvData" Grid.Row="0">
            <ListView.View>
                <GridView>
                    <!-- Binding vào DataTable cần cặp ngoặc vuông [] -->
                    <GridViewColumn Header="#" DisplayMemberBinding="{Binding [Index]}" Width="50"/>
                    <GridViewColumn Header="Name" DisplayMemberBinding="{Binding [Name]}" Width="300"/>
                </GridView>
            </ListView.View>
        </ListView>

        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button Name="btnRun" Content="Run" Width="100" Height="30" Margin="0,0,10,0" />
            <Button Name="btnCancel" Content="Cancel" Width="100" Height="30" />
        </StackPanel>
    </Grid>
</Window>
"""
#endregion

#region Helpers
def to_double_array(lst):
    """Chuyển đổi list Python thành mảng Double tương thích COM API"""
    return Array[Double]([float(v) for v in lst])
#endregion

#region Main App Class
class AppWindow(object):
    def __init__(self):
        # 1. Khởi tạo UI từ XAML
        self.window = XamlReader.Parse(WPF_XAML)
        self._find_controls()
        self._bind_events()
        self.data_items = []
        
        # 2. Khởi tạo DataTable để Binding với ListView WPF
        self.dt_data = DataTable("Data")
        self.dt_data.Columns.Add("Index", System.Int32)
        self.dt_data.Columns.Add("Name", System.String)
        self.lvData.ItemsSource = self.dt_data.DefaultView

    def _find(self, name):
        return self.window.FindName(name)

    def _find_controls(self):
        self.lvData = self._find("lvData")
        self.btnRun = self._find("btnRun")
        self.btnCancel = self._find("btnCancel")
        
    def _bind_events(self):
        self.btnRun.Click += self.OnRun
        self.btnCancel.Click += self.OnCancel

    def UpdateUI(self):
        """Ép WPF làm mới giao diện (tránh đóng băng khi chạy vòng lặp nặng)"""
        def empty_action(): pass
        self.window.Dispatcher.Invoke(DispatcherPriority.Background, System.Action(empty_action))

    def PopulateData(self):
        """Ví dụ điền dữ liệu vào DataTable"""
        self.dt_data.Clear()
        for i in range(5):
            row = self.dt_data.NewRow()
            row["Index"] = i + 1
            row["Name"] = "Item {}".format(i + 1)
            self.dt_data.Rows.Add(row)
            self.data_items.append({"Index": i + 1, "Name": "Item {}".format(i + 1)})
        self.UpdateUI()

    def OnRun(self, sender, e):
        """Hành động khi bấm Run - Ví dụ gọi AutoCAD COM"""
        try:
            # Lấy AutoCAD đang chạy
            acad_app = System.Runtime.InteropServices.Marshal.GetActiveObject("AutoCAD.Application")
            doc = acad_app.ActiveDocument
            MessageBox.Show("Kết nối AutoCAD thành công: " + doc.Name, "Success")
            
            self.PopulateData()
            
        except Exception as err:
            MessageBox.Show("Lỗi thực thi: " + str(err), "Error", MessageBoxButton.OK, MessageBoxImage.Error)

    def OnCancel(self, sender, e):
        self.window.Close()

    def show(self):
        self.window.ShowDialog()
#endregion

#region Execute
def main():
    app = AppWindow()
    app.show()
    return "Success"

# Hỗ trợ chạy trong cả Dynamo và pyRevit
OUT = main()
```

---