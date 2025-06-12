"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections
import os

clr.AddReference("ProtoGeometry")

from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB import Line, XYZ, IntersectionResultArray, SetComparisonResult

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.UI import*

clr.AddReference("System") 
from System.Collections.Generic import List

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms.DataVisualization")

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *

clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel
import System.Runtime.InteropServices


'''___'''
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
'''___'''

class TimeTracker:
    def __init__(self):
        self.start_time = None
        self.total_time = TimeSpan.Zero
        self.parameter_name = "TotalWorkingTime"
        
        # Đăng ký sự kiện
        app.DocumentOpened += DocumentOpenedEventHandler(self.OnDocumentOpened)
        app.DocumentClosed += DocumentClosedEventHandler(self.OnDocumentClosed)
    
    def OnDocumentOpened(self, sender, args):
        """Ghi lại thời gian khi mở mô hình."""
        self.start_time = DateTime.Now
        TaskDialog.Show("Time Tracker", "Bắt đầu theo dõi thời gian: {0}".format(self.start_time))
    
    def OnDocumentClosed(self, sender, args):
        """Tính và lưu thời gian khi đóng mô hình."""
        if self.start_time:
            end_time = DateTime.Now
            session_time = end_time - self.start_time
            self.total_time += session_time
            
            # Lưu vào shared parameter
            self.SaveTotalTime(args.Document)
            
            TaskDialog.Show("Time Tracker", "Phiên làm việc: {0}\nTổng thời gian: {1}".format(
                session_time, self.total_time))
            self.start_time = None
    
    def SaveTotalTime(self, document):
        """Lưu tổng thời gian vào shared parameter."""
        try:
            TransactionManager.Instance.EnsureInTransaction(document)
            
            # Lấy Project Information
            proj_info = document.ProjectInformation
            
            # Cập nhật hoặc tạo shared parameter
            param = proj_info.LookupParameter(self.parameter_name)
            if param:
                param.Set(self.total_time.TotalHours.ToString())  # Lưu dưới dạng giờ
            
            TransactionManager.Instance.TransactionTaskDone()
        except Exception as e:
            TransactionManager.Instance.TransactionTaskDone()
            TaskDialog.Show("Lỗi", "Không thể lưu thời gian: {0}".format(str(e)))

# Khởi tạo tracker
tracker = TimeTracker()

OUT = tracker