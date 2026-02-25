import clr
import sys 
import System   
import os
import time

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB import SaveAsOptions
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB.Plumbing import*
clr.AddReference("RevitAPIUI") 
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

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *

"""_______________________________________________________________________________________"""
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app   = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

#region__function
def getAllFamiliesToExport(doc):
    """Thu thập tất cả Family từ 3 category trong 1 lần duyệt, bỏ qua System Family"""
    families = {}
    categories = [
        BuiltInCategory.OST_PipeFitting,
        BuiltInCategory.OST_PipeAccessory,
        BuiltInCategory.OST_GenericModel,
    ]
    for cat in categories:
        collector = (FilteredElementCollector(doc)
                     .OfCategory(cat)
                     .WhereElementIsElementType()
                     .ToElements())
        for symbol in collector:
            try:
                if not isinstance(symbol, FamilySymbol):
                    continue
                family = symbol.Family
                if not family.IsEditable:
                    continue
                fid = family.Id.IntegerValue
                if fid not in families:
                    families[fid] = family
            except:
                pass
    return list(families.values())

def exportFamilies(doc, families, folder, overwrite, progress_callback=None, cancel_check=None):
    """
    Xuất family ra folder.
    - overwrite=False  → bỏ qua file đã tồn tại
    - cancel_check()   → trả về True nếu user yêu cầu dừng
    """
    success  = []
    skipped  = []
    failed   = []
    total    = len(families)

    save_options = SaveAsOptions()
    save_options.OverwriteExistingFile = True

    DOEVENTS_INTERVAL = 0.3
    last_doevents     = time.time()

    for i, family in enumerate(families):
        # ── Kiểm tra cancel trước mỗi family ──────────────────
        if cancel_check and cancel_check():
            failed.append("... export cancelled by user at {}/{}.".format(i, total))
            break

        family_name = ""
        try:
            family_name = family.Name
            file_path   = os.path.join(folder, family_name + ".rfa")

            if not overwrite and os.path.isfile(file_path):
                skipped.append(family_name)
            else:
                family_doc = doc.EditFamily(family)
                family_doc.SaveAs(file_path, save_options)
                family_doc.Close(False)
                success.append(family_name)

        except Exception as ex:
            failed.append("{} → {}".format(family_name or "?", str(ex)))

        # Throttle DoEvents theo interval thời gian
        now = time.time()
        if progress_callback and (now - last_doevents >= DOEVENTS_INTERVAL):
            progress_callback(i + 1, total, family_name)
            System.Windows.Forms.Application.DoEvents()
            last_doevents = now

    # Cập nhật lần cuối
    if progress_callback:
        progress_callback(len(success) + len(skipped), total, "")

    return success, skipped, failed
#endregion

#region__UI
class ExportAllForm(Form):
    def __init__(self):
        self._families        = []
        self._isExporting     = False
        self._cancelRequested = False
        self.InitializeComponent()
        self._loadFamilies()

    def _loadFamilies(self):
        self._lb_status.Text      = "Scanning families..."
        self._lb_status.ForeColor = System.Drawing.Color.DarkOrange
        System.Windows.Forms.Application.DoEvents()

        self._families = getAllFamiliesToExport(doc)
        total = len(self._families)

        self._lb_count.Text        = "Families found: {}  (Pipe Fittings + Pipe Accessories + Generic Models)".format(total)
        self._progressBar.Maximum  = max(total, 1)
        self._lb_progressText.Text = "0 / {}".format(total)
        self._lb_status.Text       = "Ready — {} families to export.".format(total)
        self._lb_status.ForeColor  = System.Drawing.Color.Gray

    def InitializeComponent(self):
        self._grb_folder      = System.Windows.Forms.GroupBox()
        self._txb_folder      = System.Windows.Forms.TextBox()
        self._btt_browse      = System.Windows.Forms.Button()

        self._grb_options     = System.Windows.Forms.GroupBox()
        self._chk_overwrite   = System.Windows.Forms.CheckBox()
        self._chk_skip        = System.Windows.Forms.CheckBox()

        self._grb_summary     = System.Windows.Forms.GroupBox()
        self._lb_count        = System.Windows.Forms.Label()
        self._lb_status       = System.Windows.Forms.Label()

        self._progressBar     = System.Windows.Forms.ProgressBar()
        self._lb_progressText = System.Windows.Forms.Label()

        self._grb_log         = System.Windows.Forms.GroupBox()
        self._txb_log         = System.Windows.Forms.TextBox()

        self._btt_EXPORT      = System.Windows.Forms.Button()
        self._btt_CANCLE      = System.Windows.Forms.Button()
        self._lb_FVC          = System.Windows.Forms.Label()

        self.SuspendLayout()

        # ── grb_folder ────────────────────────────────────────
        self._grb_folder.Text      = "Output Folder"
        self._grb_folder.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._grb_folder.ForeColor = System.Drawing.Color.Red
        self._grb_folder.Location  = System.Drawing.Point(12, 12)
        self._grb_folder.Size      = System.Drawing.Size(576, 55)
        self._grb_folder.Controls.Add(self._btt_browse)
        self._grb_folder.Controls.Add(self._txb_folder)

        self._btt_browse.Text      = "Browse..."
        self._btt_browse.Font      = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Bold)
        self._btt_browse.ForeColor = System.Drawing.Color.Blue
        self._btt_browse.Location  = System.Drawing.Point(6, 22)
        self._btt_browse.Size      = System.Drawing.Size(80, 24)
        self._btt_browse.UseVisualStyleBackColor = True
        self._btt_browse.Click    += self.Btt_browseClick

        self._txb_folder.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._txb_folder.Location  = System.Drawing.Point(92, 22)
        self._txb_folder.Size      = System.Drawing.Size(472, 24)
        self._txb_folder.ReadOnly  = True

        # ── grb_options ───────────────────────────────────────
        self._grb_options.Text      = "Options"
        self._grb_options.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._grb_options.ForeColor = System.Drawing.Color.Red
        self._grb_options.Location  = System.Drawing.Point(12, 75)
        self._grb_options.Size      = System.Drawing.Size(576, 50)
        self._grb_options.Controls.Add(self._chk_overwrite)
        self._grb_options.Controls.Add(self._chk_skip)

        self._chk_overwrite.Text      = "Overwrite existing files"
        self._chk_overwrite.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._chk_overwrite.ForeColor = System.Drawing.Color.Black
        self._chk_overwrite.Location  = System.Drawing.Point(8, 22)
        self._chk_overwrite.Size      = System.Drawing.Size(210, 20)
        self._chk_overwrite.Checked   = True
        self._chk_overwrite.CheckedChanged += self.Chk_overwriteChanged

        self._chk_skip.Text      = "Skip already exported  (faster re-run)"
        self._chk_skip.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._chk_skip.ForeColor = System.Drawing.Color.DarkGreen
        self._chk_skip.Location  = System.Drawing.Point(230, 22)
        self._chk_skip.Size      = System.Drawing.Size(300, 20)
        self._chk_skip.Checked   = False
        self._chk_skip.CheckedChanged += self.Chk_skipChanged

        # ── grb_summary ───────────────────────────────────────
        self._grb_summary.Text      = "Summary"
        self._grb_summary.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._grb_summary.ForeColor = System.Drawing.Color.Red
        self._grb_summary.Location  = System.Drawing.Point(12, 133)
        self._grb_summary.Size      = System.Drawing.Size(576, 55)
        self._grb_summary.Controls.Add(self._lb_count)
        self._grb_summary.Controls.Add(self._lb_status)

        self._lb_count.Text      = "Scanning..."
        self._lb_count.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._lb_count.ForeColor = System.Drawing.Color.Black
        self._lb_count.Location  = System.Drawing.Point(8, 18)
        self._lb_count.Size      = System.Drawing.Size(555, 18)

        self._lb_status.Text      = "Initializing..."
        self._lb_status.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._lb_status.ForeColor = System.Drawing.Color.Gray
        self._lb_status.Location  = System.Drawing.Point(8, 36)
        self._lb_status.Size      = System.Drawing.Size(555, 18)

        # ── progressBar + text ────────────────────────────────
        self._progressBar.Location = System.Drawing.Point(12, 196)
        self._progressBar.Size     = System.Drawing.Size(480, 20)
        self._progressBar.Minimum  = 0
        self._progressBar.Maximum  = 1
        self._progressBar.Value    = 0
        self._progressBar.Style    = System.Windows.Forms.ProgressBarStyle.Continuous

        self._lb_progressText.Text      = "0 / 0"
        self._lb_progressText.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._lb_progressText.ForeColor = System.Drawing.Color.Black
        self._lb_progressText.Location  = System.Drawing.Point(498, 196)
        self._lb_progressText.Size      = System.Drawing.Size(90, 20)
        self._lb_progressText.TextAlign = System.Drawing.ContentAlignment.MiddleRight

        # ── grb_log ───────────────────────────────────────────
        self._grb_log.Text      = "Export Log"
        self._grb_log.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._grb_log.ForeColor = System.Drawing.Color.Red
        self._grb_log.Location  = System.Drawing.Point(12, 224)
        self._grb_log.Size      = System.Drawing.Size(576, 150)
        self._grb_log.Controls.Add(self._txb_log)

        self._txb_log.Multiline   = True
        self._txb_log.ScrollBars  = System.Windows.Forms.ScrollBars.Vertical
        self._txb_log.ReadOnly    = True
        self._txb_log.Font        = System.Drawing.Font("Consolas", 7)
        self._txb_log.ForeColor   = System.Drawing.Color.Black
        self._txb_log.BackColor   = System.Drawing.Color.White
        self._txb_log.Location    = System.Drawing.Point(6, 20)
        self._txb_log.Size        = System.Drawing.Size(560, 122)

        # ── buttons ───────────────────────────────────────────
        self._btt_EXPORT.Text      = "EXPORT\nALL"
        self._btt_EXPORT.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._btt_EXPORT.ForeColor = System.Drawing.Color.Red
        self._btt_EXPORT.Location  = System.Drawing.Point(399, 386)
        self._btt_EXPORT.Size      = System.Drawing.Size(85, 45)
        self._btt_EXPORT.UseVisualStyleBackColor = True
        self._btt_EXPORT.Click    += self.Btt_EXPORTClick

        self._btt_CANCLE.Text      = "CLOSE"
        self._btt_CANCLE.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
        self._btt_CANCLE.Location  = System.Drawing.Point(490, 386)
        self._btt_CANCLE.Size      = System.Drawing.Size(85, 45)
        self._btt_CANCLE.UseVisualStyleBackColor = True
        self._btt_CANCLE.Click    += self.Btt_CANCLEClick

        self._lb_FVC.Text      = "@FVC"
        self._lb_FVC.Font      = System.Drawing.Font("Meiryo UI", 4.8, System.Drawing.FontStyle.Bold)
        self._lb_FVC.ForeColor = System.Drawing.Color.Black
        self._lb_FVC.Location  = System.Drawing.Point(12, 400)
        self._lb_FVC.Size      = System.Drawing.Size(48, 16)

        # ── MainForm ──────────────────────────────────────────
        self.Text            = "Export All Families"
        self.ClientSize      = System.Drawing.Size(600, 440)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.TopMost         = True
        self.Controls.Add(self._grb_folder)
        self.Controls.Add(self._grb_options)
        self.Controls.Add(self._grb_summary)
        self.Controls.Add(self._progressBar)
        self.Controls.Add(self._lb_progressText)
        self.Controls.Add(self._grb_log)
        self.Controls.Add(self._btt_EXPORT)
        self.Controls.Add(self._btt_CANCLE)
        self.Controls.Add(self._lb_FVC)
        # Bắt sự kiện đóng form
        self.FormClosing += self.MainForm_FormClosing
        self.ResumeLayout(False)
#endregion

#region__binding
    def MainForm_FormClosing(self, sender, e):
        """Ngăn đóng form giữa chừng, yêu cầu xác nhận cancel"""
        if self._isExporting:
            result = MessageBox.Show(
                "Export is in progress.\nDo you want to cancel and stop exporting?",
                "Cancel Export",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning)
            if result == System.Windows.Forms.DialogResult.Yes:
                # Ra hiệu dừng vòng lặp, giữ form mở để chờ family hiện tại xong
                self._cancelRequested     = True
                self._lb_status.Text      = "Cancelling... waiting for current family to finish."
                self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
            # Dù Yes hay No đều giữ form, vòng lặp sẽ tự đóng sau khi break
            e.Cancel = True

    def Chk_overwriteChanged(self, sender, e):
        if self._chk_overwrite.Checked:
            self._chk_skip.Checked = False

    def Chk_skipChanged(self, sender, e):
        if self._chk_skip.Checked:
            self._chk_overwrite.Checked = False

    def Btt_browseClick(self, sender, e):
        dlg = FolderBrowserDialog()
        dlg.Description = "Select folder to export families"
        if dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK:
            self._txb_folder.Text = dlg.SelectedPath

    def Btt_EXPORTClick(self, sender, e):
        folder = self._txb_folder.Text.strip()
        if not folder:
            MessageBox.Show("Please select an output folder first.",
                "No Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        if not os.path.isdir(folder):
            MessageBox.Show("The selected folder does not exist.",
                "Invalid Folder", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        if self._isExporting:
            return

        overwrite = self._chk_overwrite.Checked

        # Reset trạng thái
        self._isExporting         = True
        self._cancelRequested     = False
        self._btt_EXPORT.Enabled  = False
        self._btt_CANCLE.Text      = "CANCEL"
        self._btt_CANCLE.ForeColor = System.Drawing.Color.OrangeRed
        self._txb_log.Clear()
        self._progressBar.Value    = 0
        self._progressBar.Maximum  = max(len(self._families), 1)
        self._lb_progressText.Text = "0 / {}".format(len(self._families))
        self._lb_status.ForeColor  = System.Drawing.Color.DarkOrange
        self._lb_status.Text       = "Exporting..."

        start_time = time.time()

        def on_progress(current, total, name):
            self._progressBar.Value    = current
            self._lb_progressText.Text = "{} / {}".format(current, total)
            self._lb_status.Text       = "[{}/{}]  {}".format(current, total, name)

        success, skipped, failed = exportFamilies(
            doc, self._families, folder, overwrite,
            progress_callback = on_progress,
            cancel_check      = lambda: self._cancelRequested)

        elapsed = time.time() - start_time

        # ── Ghi log ──────────────────────────────────────────
        lines = []
        lines.append("=" * 58)
        if self._cancelRequested:
            lines.append("EXPORT CANCELLED  ({:.1f}s)".format(elapsed))
        else:
            lines.append("EXPORT COMPLETE  ({:.1f}s)".format(elapsed))
        lines.append("  Exported : {}".format(len(success)))
        lines.append("  Skipped  : {}  (file already exists)".format(len(skipped)))
        lines.append("  Failed   : {}".format(len(failed)))
        lines.append("=" * 58)
        if success:
            lines.append("[OK]")
            for n in success:
                lines.append("  + " + n)
        if skipped:
            lines.append("[SKIPPED]")
            for n in skipped:
                lines.append("  - " + n)
        if failed:
            lines.append("[FAILED / CANCELLED]")
            for m in failed:
                lines.append("  x " + m)

        self._txb_log.Text = "\r\n".join(lines)

        # ── Trạng thái cuối ───────────────────────────────────
        if self._cancelRequested:
            self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
            self._lb_status.Text = "Cancelled ({:.1f}s) — {} exported, {} skipped, {} failed/cancelled.".format(
                elapsed, len(success), len(skipped), len(failed))
        elif failed:
            self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
            self._lb_status.Text = "Done ({:.1f}s) — {} ok, {} skipped, {} failed.".format(
                elapsed, len(success), len(skipped), len(failed))
        else:
            self._lb_status.ForeColor = System.Drawing.Color.Green
            self._lb_status.Text = "Done ({:.1f}s) — {} exported, {} skipped.".format(
                elapsed, len(success), len(skipped))

        # Khôi phục UI
        self._isExporting          = False
        self._cancelRequested      = False
        self._btt_EXPORT.Enabled   = True
        self._btt_CANCLE.Text      = "CLOSE"
        self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
        self._btt_CANCLE.Enabled   = True

        # Nếu user đã yêu cầu đóng form trong lúc cancel → đóng thật sự
        # Gỡ event FormClosing tạm để đóng không bị chặn
        self.FormClosing -= self.MainForm_FormClosing
        if self._cancelRequested is False and not self.IsDisposed:
            pass  # giữ form để user xem log
        # Re-attach event
        self.FormClosing += self.MainForm_FormClosing

    def Btt_CANCLEClick(self, sender, e):
        if self._isExporting:
            # Nút đổi thành CANCEL khi đang export
            result = MessageBox.Show(
                "Export is in progress.\nDo you want to cancel and stop exporting?",
                "Cancel Export",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning)
            if result == System.Windows.Forms.DialogResult.Yes:
                self._cancelRequested     = True
                self._lb_status.Text      = "Cancelling... waiting for current family to finish."
                self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
                self._btt_CANCLE.Enabled  = False
        else:
            self.Close()
#endregion

f = ExportAllForm()
Application.Run(f)
```

### Tóm tắt luồng Cancel
```
User nhấn CANCEL / X  
         ↓  
Hộp thoại xác nhận Yes/No  
         ↓ Yes  
_cancelRequested = True  
         ↓  
Vòng lặp kiểm tra flag → break sau family hiện tại  
         ↓  
Log hiển thị "CANCELLED", UI khôi phục bình thường  
         ↓  
User nhấn CLOSE để thoátimport clr
import sys 
import System   
import os
import time

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB import SaveAsOptions
from Autodesk.Revit.DB.Structure import*
from Autodesk.Revit.DB.Plumbing import*
clr.AddReference("RevitAPIUI") 
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

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *

"""_______________________________________________________________________________________"""
doc   = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app   = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

#region__function
def getAllFamiliesToExport(doc):
    """Thu thập tất cả Family từ 3 category trong 1 lần duyệt, bỏ qua System Family"""
    families = {}
    categories = [
        BuiltInCategory.OST_PipeFitting,
        BuiltInCategory.OST_PipeAccessory,
        BuiltInCategory.OST_GenericModel,
    ]
    for cat in categories:
        collector = (FilteredElementCollector(doc)
                     .OfCategory(cat)
                     .WhereElementIsElementType()
                     .ToElements())
        for symbol in collector:
            try:
                if not isinstance(symbol, FamilySymbol):
                    continue
                family = symbol.Family
                if not family.IsEditable:
                    continue
                fid = family.Id.IntegerValue
                if fid not in families:
                    families[fid] = family
            except:
                pass
    return list(families.values())

def exportFamilies(doc, families, folder, overwrite, progress_callback=None, cancel_check=None):
    """
    Xuất family ra folder.
    - overwrite=False  → bỏ qua file đã tồn tại
    - cancel_check()   → trả về True nếu user yêu cầu dừng
    """
    success  = []
    skipped  = []
    failed   = []
    total    = len(families)

    save_options = SaveAsOptions()
    save_options.OverwriteExistingFile = True

    DOEVENTS_INTERVAL = 0.3
    last_doevents     = time.time()

    for i, family in enumerate(families):
        # ── Kiểm tra cancel trước mỗi family ──────────────────
        if cancel_check and cancel_check():
            failed.append("... export cancelled by user at {}/{}.".format(i, total))
            break

        family_name = ""
        try:
            family_name = family.Name
            file_path   = os.path.join(folder, family_name + ".rfa")

            if not overwrite and os.path.isfile(file_path):
                skipped.append(family_name)
            else:
                family_doc = doc.EditFamily(family)
                family_doc.SaveAs(file_path, save_options)
                family_doc.Close(False)
                success.append(family_name)

        except Exception as ex:
            failed.append("{} → {}".format(family_name or "?", str(ex)))

        # Throttle DoEvents theo interval thời gian
        now = time.time()
        if progress_callback and (now - last_doevents >= DOEVENTS_INTERVAL):
            progress_callback(i + 1, total, family_name)
            System.Windows.Forms.Application.DoEvents()
            last_doevents = now

    # Cập nhật lần cuối
    if progress_callback:
        progress_callback(len(success) + len(skipped), total, "")

    return success, skipped, failed
#endregion

#region__UI
class ExportAllForm(Form):
    def __init__(self):
        self._families        = []
        self._isExporting     = False
        self._cancelRequested = False
        self.InitializeComponent()
        self._loadFamilies()

    def _loadFamilies(self):
        self._lb_status.Text      = "Scanning families..."
        self._lb_status.ForeColor = System.Drawing.Color.DarkOrange
        System.Windows.Forms.Application.DoEvents()

        self._families = getAllFamiliesToExport(doc)
        total = len(self._families)

        self._lb_count.Text        = "Families found: {}  (Pipe Fittings + Pipe Accessories + Generic Models)".format(total)
        self._progressBar.Maximum  = max(total, 1)
        self._lb_progressText.Text = "0 / {}".format(total)
        self._lb_status.Text       = "Ready — {} families to export.".format(total)
        self._lb_status.ForeColor  = System.Drawing.Color.Gray

    def InitializeComponent(self):
        self._grb_folder      = System.Windows.Forms.GroupBox()
        self._txb_folder      = System.Windows.Forms.TextBox()
        self._btt_browse      = System.Windows.Forms.Button()

        self._grb_options     = System.Windows.Forms.GroupBox()
        self._chk_overwrite   = System.Windows.Forms.CheckBox()
        self._chk_skip        = System.Windows.Forms.CheckBox()

        self._grb_summary     = System.Windows.Forms.GroupBox()
        self._lb_count        = System.Windows.Forms.Label()
        self._lb_status       = System.Windows.Forms.Label()

        self._progressBar     = System.Windows.Forms.ProgressBar()
        self._lb_progressText = System.Windows.Forms.Label()

        self._grb_log         = System.Windows.Forms.GroupBox()
        self._txb_log         = System.Windows.Forms.TextBox()

        self._btt_EXPORT      = System.Windows.Forms.Button()
        self._btt_CANCLE      = System.Windows.Forms.Button()
        self._lb_FVC          = System.Windows.Forms.Label()

        self.SuspendLayout()

        # ── grb_folder ────────────────────────────────────────
        self._grb_folder.Text      = "Output Folder"
        self._grb_folder.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._grb_folder.ForeColor = System.Drawing.Color.Red
        self._grb_folder.Location  = System.Drawing.Point(12, 12)
        self._grb_folder.Size      = System.Drawing.Size(576, 55)
        self._grb_folder.Controls.Add(self._btt_browse)
        self._grb_folder.Controls.Add(self._txb_folder)

        self._btt_browse.Text      = "Browse..."
        self._btt_browse.Font      = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Bold)
        self._btt_browse.ForeColor = System.Drawing.Color.Blue
        self._btt_browse.Location  = System.Drawing.Point(6, 22)
        self._btt_browse.Size      = System.Drawing.Size(80, 24)
        self._btt_browse.UseVisualStyleBackColor = True
        self._btt_browse.Click    += self.Btt_browseClick

        self._txb_folder.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._txb_folder.Location  = System.Drawing.Point(92, 22)
        self._txb_folder.Size      = System.Drawing.Size(472, 24)
        self._txb_folder.ReadOnly  = True

        # ── grb_options ───────────────────────────────────────
        self._grb_options.Text      = "Options"
        self._grb_options.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._grb_options.ForeColor = System.Drawing.Color.Red
        self._grb_options.Location  = System.Drawing.Point(12, 75)
        self._grb_options.Size      = System.Drawing.Size(576, 50)
        self._grb_options.Controls.Add(self._chk_overwrite)
        self._grb_options.Controls.Add(self._chk_skip)

        self._chk_overwrite.Text      = "Overwrite existing files"
        self._chk_overwrite.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._chk_overwrite.ForeColor = System.Drawing.Color.Black
        self._chk_overwrite.Location  = System.Drawing.Point(8, 22)
        self._chk_overwrite.Size      = System.Drawing.Size(210, 20)
        self._chk_overwrite.Checked   = True
        self._chk_overwrite.CheckedChanged += self.Chk_overwriteChanged

        self._chk_skip.Text      = "Skip already exported  (faster re-run)"
        self._chk_skip.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._chk_skip.ForeColor = System.Drawing.Color.DarkGreen
        self._chk_skip.Location  = System.Drawing.Point(230, 22)
        self._chk_skip.Size      = System.Drawing.Size(300, 20)
        self._chk_skip.Checked   = False
        self._chk_skip.CheckedChanged += self.Chk_skipChanged

        # ── grb_summary ───────────────────────────────────────
        self._grb_summary.Text      = "Summary"
        self._grb_summary.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._grb_summary.ForeColor = System.Drawing.Color.Red
        self._grb_summary.Location  = System.Drawing.Point(12, 133)
        self._grb_summary.Size      = System.Drawing.Size(576, 55)
        self._grb_summary.Controls.Add(self._lb_count)
        self._grb_summary.Controls.Add(self._lb_status)

        self._lb_count.Text      = "Scanning..."
        self._lb_count.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._lb_count.ForeColor = System.Drawing.Color.Black
        self._lb_count.Location  = System.Drawing.Point(8, 18)
        self._lb_count.Size      = System.Drawing.Size(555, 18)

        self._lb_status.Text      = "Initializing..."
        self._lb_status.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._lb_status.ForeColor = System.Drawing.Color.Gray
        self._lb_status.Location  = System.Drawing.Point(8, 36)
        self._lb_status.Size      = System.Drawing.Size(555, 18)

        # ── progressBar + text ────────────────────────────────
        self._progressBar.Location = System.Drawing.Point(12, 196)
        self._progressBar.Size     = System.Drawing.Size(480, 20)
        self._progressBar.Minimum  = 0
        self._progressBar.Maximum  = 1
        self._progressBar.Value    = 0
        self._progressBar.Style    = System.Windows.Forms.ProgressBarStyle.Continuous

        self._lb_progressText.Text      = "0 / 0"
        self._lb_progressText.Font      = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Bold)
        self._lb_progressText.ForeColor = System.Drawing.Color.Black
        self._lb_progressText.Location  = System.Drawing.Point(498, 196)
        self._lb_progressText.Size      = System.Drawing.Size(90, 20)
        self._lb_progressText.TextAlign = System.Drawing.ContentAlignment.MiddleRight

        # ── grb_log ───────────────────────────────────────────
        self._grb_log.Text      = "Export Log"
        self._grb_log.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._grb_log.ForeColor = System.Drawing.Color.Red
        self._grb_log.Location  = System.Drawing.Point(12, 224)
        self._grb_log.Size      = System.Drawing.Size(576, 150)
        self._grb_log.Controls.Add(self._txb_log)

        self._txb_log.Multiline   = True
        self._txb_log.ScrollBars  = System.Windows.Forms.ScrollBars.Vertical
        self._txb_log.ReadOnly    = True
        self._txb_log.Font        = System.Drawing.Font("Consolas", 7)
        self._txb_log.ForeColor   = System.Drawing.Color.Black
        self._txb_log.BackColor   = System.Drawing.Color.White
        self._txb_log.Location    = System.Drawing.Point(6, 20)
        self._txb_log.Size        = System.Drawing.Size(560, 122)

        # ── buttons ───────────────────────────────────────────
        self._btt_EXPORT.Text      = "EXPORT\nALL"
        self._btt_EXPORT.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._btt_EXPORT.ForeColor = System.Drawing.Color.Red
        self._btt_EXPORT.Location  = System.Drawing.Point(399, 386)
        self._btt_EXPORT.Size      = System.Drawing.Size(85, 45)
        self._btt_EXPORT.UseVisualStyleBackColor = True
        self._btt_EXPORT.Click    += self.Btt_EXPORTClick

        self._btt_CANCLE.Text      = "CLOSE"
        self._btt_CANCLE.Font      = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold)
        self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
        self._btt_CANCLE.Location  = System.Drawing.Point(490, 386)
        self._btt_CANCLE.Size      = System.Drawing.Size(85, 45)
        self._btt_CANCLE.UseVisualStyleBackColor = True
        self._btt_CANCLE.Click    += self.Btt_CANCLEClick

        self._lb_FVC.Text      = "@FVC"
        self._lb_FVC.Font      = System.Drawing.Font("Meiryo UI", 4.8, System.Drawing.FontStyle.Bold)
        self._lb_FVC.ForeColor = System.Drawing.Color.Black
        self._lb_FVC.Location  = System.Drawing.Point(12, 400)
        self._lb_FVC.Size      = System.Drawing.Size(48, 16)

        # ── MainForm ──────────────────────────────────────────
        self.Text            = "Export All Families"
        self.ClientSize      = System.Drawing.Size(600, 440)
        self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        self.TopMost         = True
        self.Controls.Add(self._grb_folder)
        self.Controls.Add(self._grb_options)
        self.Controls.Add(self._grb_summary)
        self.Controls.Add(self._progressBar)
        self.Controls.Add(self._lb_progressText)
        self.Controls.Add(self._grb_log)
        self.Controls.Add(self._btt_EXPORT)
        self.Controls.Add(self._btt_CANCLE)
        self.Controls.Add(self._lb_FVC)
        # Bắt sự kiện đóng form
        self.FormClosing += self.MainForm_FormClosing
        self.ResumeLayout(False)
#endregion

#region__binding
    def MainForm_FormClosing(self, sender, e):
        """Ngăn đóng form giữa chừng, yêu cầu xác nhận cancel"""
        if self._isExporting:
            result = MessageBox.Show(
                "Export is in progress.\nDo you want to cancel and stop exporting?",
                "Cancel Export",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning)
            if result == System.Windows.Forms.DialogResult.Yes:
                # Ra hiệu dừng vòng lặp, giữ form mở để chờ family hiện tại xong
                self._cancelRequested     = True
                self._lb_status.Text      = "Cancelling... waiting for current family to finish."
                self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
            # Dù Yes hay No đều giữ form, vòng lặp sẽ tự đóng sau khi break
            e.Cancel = True

    def Chk_overwriteChanged(self, sender, e):
        if self._chk_overwrite.Checked:
            self._chk_skip.Checked = False

    def Chk_skipChanged(self, sender, e):
        if self._chk_skip.Checked:
            self._chk_overwrite.Checked = False

    def Btt_browseClick(self, sender, e):
        dlg = FolderBrowserDialog()
        dlg.Description = "Select folder to export families"
        if dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK:
            self._txb_folder.Text = dlg.SelectedPath

    def Btt_EXPORTClick(self, sender, e):
        folder = self._txb_folder.Text.strip()
        if not folder:
            MessageBox.Show("Please select an output folder first.",
                "No Folder", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        if not os.path.isdir(folder):
            MessageBox.Show("The selected folder does not exist.",
                "Invalid Folder", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        if self._isExporting:
            return

        overwrite = self._chk_overwrite.Checked

        # Reset trạng thái
        self._isExporting         = True
        self._cancelRequested     = False
        self._btt_EXPORT.Enabled  = False
        self._btt_CANCLE.Text      = "CANCEL"
        self._btt_CANCLE.ForeColor = System.Drawing.Color.OrangeRed
        self._txb_log.Clear()
        self._progressBar.Value    = 0
        self._progressBar.Maximum  = max(len(self._families), 1)
        self._lb_progressText.Text = "0 / {}".format(len(self._families))
        self._lb_status.ForeColor  = System.Drawing.Color.DarkOrange
        self._lb_status.Text       = "Exporting..."

        start_time = time.time()

        def on_progress(current, total, name):
            self._progressBar.Value    = current
            self._lb_progressText.Text = "{} / {}".format(current, total)
            self._lb_status.Text       = "[{}/{}]  {}".format(current, total, name)

        success, skipped, failed = exportFamilies(
            doc, self._families, folder, overwrite,
            progress_callback = on_progress,
            cancel_check      = lambda: self._cancelRequested)

        elapsed = time.time() - start_time

        # ── Ghi log ──────────────────────────────────────────
        lines = []
        lines.append("=" * 58)
        if self._cancelRequested:
            lines.append("EXPORT CANCELLED  ({:.1f}s)".format(elapsed))
        else:
            lines.append("EXPORT COMPLETE  ({:.1f}s)".format(elapsed))
        lines.append("  Exported : {}".format(len(success)))
        lines.append("  Skipped  : {}  (file already exists)".format(len(skipped)))
        lines.append("  Failed   : {}".format(len(failed)))
        lines.append("=" * 58)
        if success:
            lines.append("[OK]")
            for n in success:
                lines.append("  + " + n)
        if skipped:
            lines.append("[SKIPPED]")
            for n in skipped:
                lines.append("  - " + n)
        if failed:
            lines.append("[FAILED / CANCELLED]")
            for m in failed:
                lines.append("  x " + m)

        self._txb_log.Text = "\r\n".join(lines)

        # ── Trạng thái cuối ───────────────────────────────────
        if self._cancelRequested:
            self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
            self._lb_status.Text = "Cancelled ({:.1f}s) — {} exported, {} skipped, {} failed/cancelled.".format(
                elapsed, len(success), len(skipped), len(failed))
        elif failed:
            self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
            self._lb_status.Text = "Done ({:.1f}s) — {} ok, {} skipped, {} failed.".format(
                elapsed, len(success), len(skipped), len(failed))
        else:
            self._lb_status.ForeColor = System.Drawing.Color.Green
            self._lb_status.Text = "Done ({:.1f}s) — {} exported, {} skipped.".format(
                elapsed, len(success), len(skipped))

        # Khôi phục UI
        self._isExporting          = False
        self._cancelRequested      = False
        self._btt_EXPORT.Enabled   = True
        self._btt_CANCLE.Text      = "CLOSE"
        self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
        self._btt_CANCLE.Enabled   = True

        # Nếu user đã yêu cầu đóng form trong lúc cancel → đóng thật sự
        # Gỡ event FormClosing tạm để đóng không bị chặn
        self.FormClosing -= self.MainForm_FormClosing
        if self._cancelRequested is False and not self.IsDisposed:
            pass  # giữ form để user xem log
        # Re-attach event
        self.FormClosing += self.MainForm_FormClosing

    def Btt_CANCLEClick(self, sender, e):
        if self._isExporting:
            # Nút đổi thành CANCEL khi đang export
            result = MessageBox.Show(
                "Export is in progress.\nDo you want to cancel and stop exporting?",
                "Cancel Export",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning)
            if result == System.Windows.Forms.DialogResult.Yes:
                self._cancelRequested     = True
                self._lb_status.Text      = "Cancelling... waiting for current family to finish."
                self._lb_status.ForeColor = System.Drawing.Color.OrangeRed
                self._btt_CANCLE.Enabled  = False
        else:
            self.Close()
#endregion

f = ExportAllForm()
Application.Run(f)
