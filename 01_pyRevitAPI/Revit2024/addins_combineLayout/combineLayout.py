"""Copyright by: Senior C# & Revit API Developer"""
#region ___import all Library
import clr
import sys
import math
import os
import glob

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

# Load WinForms for Folder/File Dialogs
clr.AddReference("System.Windows.Forms")
import System.Windows.Forms as WinForms

# Load Revit API (Standard Boilerplate)
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# Load COM Interop for AutoCAD
from System.Runtime.InteropServices import Marshal
#endregion

#region XAML UI Definition
WPF_XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="DWG Layout Merger - Auto CAD Automation" 
        Width="700" Height="600" 
        WindowStartupLocation="CenterScreen" 
        Background="#F5F6F8" FontFamily="Segoe UI" FontSize="13">
    <Window.Resources>
        <!-- Modern Button Style -->
        <Style x:Key="ModernBtn" TargetType="Button">
            <Setter Property="Background" Value="#0078D7"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#106EBE"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#CCCCCC"/>
                    <Setter Property="Foreground" Value="#888888"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- 1. Path Selection -->
        <GroupBox Header="1. File Selection" Grid.Row="0" Padding="10" Margin="0,0,0,10" BorderBrush="#DDDDDD">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="100"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="80"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="Source Folder:" Grid.Row="0" Grid.Column="0" VerticalAlignment="Center"/>
                <TextBox Name="txtSourceDir" Grid.Row="0" Grid.Column="1" Margin="5" IsReadOnly="True" Background="#EEEEEE"/>
                <Button Name="btnBrowseSource" Grid.Row="0" Grid.Column="2" Content="Browse" Margin="5" Style="{StaticResource ModernBtn}"/>

                <TextBlock Text="Output File:" Grid.Row="1" Grid.Column="0" VerticalAlignment="Center"/>
                <TextBox Name="txtOutFile" Grid.Row="1" Grid.Column="1" Margin="5"/>
                <Button Name="btnBrowseOut" Grid.Row="1" Grid.Column="2" Content="Browse" Margin="5" Style="{StaticResource ModernBtn}"/>
            </Grid>
        </GroupBox>

        <!-- 2. Processing Options -->
        <GroupBox Header="2. Options" Grid.Row="1" Padding="10" Margin="0,0,0,10" BorderBrush="#DDDDDD">
            <WrapPanel Orientation="Horizontal">
                <CheckBox Name="chkNumbering" Content="Add Prefix Numbering (01_, 02_)" Margin="0,5,15,5" IsChecked="True"/>
                <CheckBox Name="chkExplode" Content="Explode Blocks" Margin="0,5,15,5"/>
                <CheckBox Name="chkDrawBox" Content="Draw Bounding Box" Margin="0,5,15,5" IsChecked="True"/>
                <CheckBox Name="chkUnlockLayers" Content="Unlock All Layers" Margin="0,5,15,5" IsChecked="True"/>
                <CheckBox Name="chkDeleteSource" Content="Delete Source Files on Success" Margin="0,5,15,5" Foreground="Red"/>
            </WrapPanel>
        </GroupBox>

        <!-- 3. File List -->
        <GroupBox Header="3. Detected DWG Files" Grid.Row="2" Padding="5" Margin="0,0,0,10" BorderBrush="#DDDDDD">
            <ListView Name="lvFiles" BorderThickness="0">
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="#" DisplayMemberBinding="{Binding [Index]}" Width="40"/>
                        <GridViewColumn Header="File Name" DisplayMemberBinding="{Binding [Name]}" Width="350"/>
                        <GridViewColumn Header="Size (KB)" DisplayMemberBinding="{Binding [Size]}" Width="100"/>
                        <GridViewColumn Header="Modified Date" DisplayMemberBinding="{Binding [Date]}" Width="150"/>
                    </GridView>
                </ListView.View>
            </ListView>
        </GroupBox>

        <!-- 4. Execution & Progress -->
        <StackPanel Grid.Row="3" Margin="0,0,0,10">
            <TextBlock Name="lblStatus" Text="Status: Ready..." FontWeight="Bold" Foreground="#555555" Margin="0,0,0,5"/>
            <ProgressBar Name="pbProgress" Height="15" Minimum="0" Maximum="100" Value="0" Foreground="#0078D7"/>
        </StackPanel>

        <!-- 5. Actions -->
        <Grid Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Label Content="@FVC DWG Merger" Foreground="Gray" Grid.Column="0" VerticalAlignment="Center"/>
            <Button Name="btnRun" Grid.Column="1" Content="MERGE LAYOUTS" Width="140" Height="35" Margin="0,0,10,0" Style="{StaticResource ModernBtn}" Background="#107C10"/>
            <Button Name="btnCancel" Grid.Column="2" Content="CANCEL" Width="100" Height="35" Style="{StaticResource ModernBtn}" Background="#A80000"/>
        </Grid>
    </Grid>
</Window>
"""
#endregion

#region Helpers
def to_double_array(lst):
    """Convert Python list to COM compatible double array"""
    return Array[Double]([float(v) for v in lst])

#endregion

#region Main App Class
class DWGMergerWindow(object):
    def __init__(self):
        self.window = XamlReader.Parse(WPF_XAML)
        self._find_controls()
        self._bind_events()
        self.dwg_files = []
        
        # Setup DataTable for WPF Binding (Like autojoin.py)
        self.dt_files = DataTable("Files")
        self.dt_files.Columns.Add("Index", System.Int32)
        self.dt_files.Columns.Add("Name", System.String)
        self.dt_files.Columns.Add("Size", System.Double)
        self.dt_files.Columns.Add("Date", System.String)
        self.lvFiles.ItemsSource = self.dt_files.DefaultView

    def _find(self, name):
        return self.window.FindName(name)

    def _find_controls(self):
        self.txtSourceDir = self._find("txtSourceDir")
        self.txtOutFile = self._find("txtOutFile")
        self.lvFiles = self._find("lvFiles")
        self.lblStatus = self._find("lblStatus")
        self.pbProgress = self._find("pbProgress")
        
        self.chkNumbering = self._find("chkNumbering")
        self.chkExplode = self._find("chkExplode")
        self.chkDrawBox = self._find("chkDrawBox")
        self.chkUnlockLayers = self._find("chkUnlockLayers")
        self.chkDeleteSource = self._find("chkDeleteSource")
        
        self.btnBrowseSource = self._find("btnBrowseSource")
        self.btnBrowseOut = self._find("btnBrowseOut")
        self.btnRun = self._find("btnRun")
        self.btnCancel = self._find("btnCancel")
        
    def _bind_events(self):
        self.btnBrowseSource.Click += self.OnBrowseSource
        self.btnBrowseOut.Click += self.OnBrowseOut
        self.btnRun.Click += self.OnRun
        self.btnCancel.Click += self.OnCancel

    def UpdateUI(self):
        """Force WPF UI to refresh"""
        def empty_action(): pass
        self.window.Dispatcher.Invoke(DispatcherPriority.Background, System.Action(empty_action))

    def LogStatus(self, msg, progress=None):
        self.lblStatus.Text = str("Status: " + msg)
        if progress is not None:
            self.pbProgress.Value = progress
        self.UpdateUI()

    def OnBrowseSource(self, sender, e):
        dlg = WinForms.FolderBrowserDialog()
        dlg.Description = "Select folder containing exported DWGs"
        if dlg.ShowDialog() == WinForms.DialogResult.OK:
            self.txtSourceDir.Text = dlg.SelectedPath
            self.txtOutFile.Text = os.path.join(dlg.SelectedPath, "CaC_Merged.dwg")
            
            # Scan files
            search_pattern = os.path.join(dlg.SelectedPath, "*.dwg")
            files = glob.glob(search_pattern)
            
            # Exclude output file if exists
            files = [f for f in files if f.lower() != self.txtOutFile.Text.lower()]
            files.sort() # Sort alphabetically
            
            self.dwg_files = []
            self.dt_files.Clear()
            import time
            
            for idx, f in enumerate(files):
                self.dwg_files.append({"Path": f, "Name": os.path.basename(f)})
                
                row = self.dt_files.NewRow()
                row["Index"] = idx + 1
                row["Name"] = os.path.basename(f)
                row["Size"] = round(os.path.getsize(f) / 1024.0, 2)
                row["Date"] = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(os.path.getmtime(f)))
                self.dt_files.Rows.Add(row)
                
            self.LogStatus("Detected {} DWG files.".format(len(self.dwg_files)))

    def OnBrowseOut(self, sender, e):
        dlg = WinForms.SaveFileDialog()
        dlg.Filter = "AutoCAD DWG (*.dwg)|*.dwg"
        dlg.FileName = "CaC_Merged.dwg"
        if dlg.ShowDialog() == WinForms.DialogResult.OK:
            self.txtOutFile.Text = dlg.FileName

    def GetAutoCADApp(self):
        """Strict Workflow: Connect to AutoCAD via COM"""
        self.LogStatus("Connecting to AutoCAD Engine...")
        try:
            # Try getting active AutoCAD
            acad = Marshal.GetActiveObject("AutoCAD.Application")
            return acad
        except:
            try:
                # Spawn new AutoCAD process
                self.LogStatus("Starting new AutoCAD process (Please wait)...")
                acadType = System.Type.GetTypeFromProgID("AutoCAD.Application")
                acad = System.Activator.CreateInstance(acadType)
                acad.Visible = True # Show it so user knows it's running
                return acad
            except Exception as ex:
                MessageBox.Show("Failed to connect to AutoCAD. Is it installed?\n\nError: " + str(ex), 
                                "AutoCAD Error", MessageBoxButton.OK, MessageBoxImage.Error)
                return None

    def OnRun(self, sender, e):
        # 1. Input Validation
        if not self.dwg_files:
            MessageBox.Show("No DWG files found in source directory.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return
            
        out_path = self.txtOutFile.Text
        if not out_path:
            MessageBox.Show("Please specify output file path.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning)
            return

        # Disable UI
        self.btnRun.IsEnabled = False
        self.btnBrowseSource.IsEnabled = False
        
        # 2. AutoCAD Connection
        acad = self.GetAutoCADApp()
        if not acad:
            self.btnRun.IsEnabled = True
            self.LogStatus("Operation aborted.")
            return

        total_files = len(self.dwg_files)
        gap_distance = 1000.0 # 1000 unit gap between drawings
        current_x = 0.0

        try:
            # Create a brand new document
            self.LogStatus("Creating new Document...")
            doc = acad.Documents.Add()
            ms = doc.ModelSpace

            # 3. Main Processing Loop
            for i, f_item in enumerate(self.dwg_files):
                pct = int((float(i) / total_files) * 100)
                self.LogStatus("Processing [{}/{}] : {}".format(i+1, total_files, f_item["Name"]), pct)
                
                # Setup insertion point
                ins_pt = to_double_array([current_x, 0.0, 0.0])
                
                # Insert DWG as Block
                try:
                    # InsertBlock(InsertionPoint, Name, Xscale, Yscale, Zscale, Rotation)
                    block_ref = ms.InsertBlock(ins_pt, f_item["Path"], 1.0, 1.0, 1.0, 0.0)
                except Exception as ins_err:
                    MessageBox.Show("Failed to insert file: {}\n\nError: {}".format(f_item["Name"], str(ins_err)), 
                                    "Insertion Error", MessageBoxButton.OK, MessageBoxImage.Error)
                    # Strict workflow: Cancel process if an insertion fails
                    doc.Close(False)
                    raise Exception("Process stopped due to insertion failure.")

                # Calculate Bounding Box of inserted block
                min_ext_ref = clr.Reference[System.Object]()
                max_ext_ref = clr.Reference[System.Object]()
                block_ref.GetBoundingBox(min_ext_ref, max_ext_ref)
                min_ext = min_ext_ref.Value
                max_ext = max_ext_ref.Value
                
                width = max_ext[0] - min_ext[0]

                # Draw Bounding Box if requested
                if self.chkDrawBox.IsChecked == True:
                    # LWPolyline vertices: X1, Y1, X2, Y2, X3, Y3, X4, Y4, X1, Y1
                    box_verts = to_double_array([
                        min_ext[0], min_ext[1],
                        max_ext[0], min_ext[1],
                        max_ext[0], max_ext[1],
                        min_ext[0], max_ext[1],
                        min_ext[0], min_ext[1]
                    ])
                    poly = ms.AddLightWeightPolyline(box_verts)
                    poly.Color = 1 # Red color for bounding box

                # Add Numbering Text if requested
                if self.chkNumbering.IsChecked == True:
                    txt_pt = to_double_array([min_ext[0], max_ext[1] + 500.0, 0.0]) # Put text 500 units above
                    text_obj = ms.AddText("{:02d}_{}".format(i+1, f_item["Name"].replace(".dwg", "")), txt_pt, 250.0) # Height 250
                    text_obj.Color = 3 # Green color

                # Explode if requested
                if self.chkExplode.IsChecked == True:
                    block_ref.Explode()
                    block_ref.Delete() # Remove original block wrapper after exploding

                # Update X for next file
                current_x += width + gap_distance

            # 4. Post Processing
            self.LogStatus("Applying post-processing...", 95)
            
            # Unlock Layers
            if self.chkUnlockLayers.IsChecked == True:
                for layer in doc.Layers:
                    layer.Lock = False

            # Zoom Extents
            acad.ZoomExtents()

            # Save File
            self.LogStatus("Saving merged DWG...", 98)
            if os.path.exists(out_path):
                os.remove(out_path)
            doc.SaveAs(out_path)
            
            # 5. Delete Sources (Dangerous operation, handled carefully)
            if self.chkDeleteSource.IsChecked == True:
                self.LogStatus("Deleting source files...")
                for f_item in self.dwg_files:
                    try:
                        os.remove(f_item["Path"])
                    except: pass

            self.LogStatus("MERGE SUCCESSFUL!", 100)
            MessageBox.Show("All files have been merged successfully into:\n" + out_path, "Success", MessageBoxButton.OK, MessageBoxImage.Information)

        except Exception as general_err:
            self.LogStatus("Error Occurred", 0)
            if "stopped" not in str(general_err):
                MessageBox.Show("An unexpected error occurred:\n\n" + str(general_err), "Error", MessageBoxButton.OK, MessageBoxImage.Error)
        finally:
            self.btnRun.IsEnabled = True
            self.btnBrowseSource.IsEnabled = True

    def OnCancel(self, sender, e):
        self.window.Close()

    def show(self):
        self.window.ShowDialog()

#endregion

def main():
    app = DWGMergerWindow()
    app.show()

# Execute directly in Dynamo Environment
main()
OUT = "Success"
