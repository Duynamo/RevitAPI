"""Copyright by: vudinhduybm@gmail.com
   Fixed: Group duplicate parameters by Name to ensure UI uniqueness.
"""
import clr
import sys
import collections

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

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

# Khởi tạo các biến Revit cơ bản
doc = DocumentManager.Instance.CurrentDBDocument

class DeleteParamForm(Form):
    def __init__(self, unique_param_names):
        primary_screen = Screen.PrimaryScreen.WorkingArea
        screen_width = primary_screen.Width // 3
        screen_height = primary_screen.Height // 2
        
        self.Text = 'Select Parameters to Delete'
        self.ClientSize = Size(screen_width, screen_height)
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        
        # all_params bây giờ chứa các tên ĐỘC NHẤT đã được sắp xếp
        self.all_params = sorted(unique_param_names)
        
        # Label
        self.label = Label()
        self.label.Text = "Parameters (Duplicates are grouped as one):"
        self.label.Location = Point(20, 20)
        self.label.AutoSize = True
        self.Controls.Add(self.label)
        
        # Search Panel
        self.searchLabel = Label()
        self.searchLabel.Text = "Search:"
        self.searchLabel.Location = Point(20, 50)
        self.searchLabel.AutoSize = True
        self.Controls.Add(self.searchLabel)
        
        self.searchBox = TextBox()
        self.searchBox.Location = Point(80, 48)
        self.searchBox.Size = Size(screen_width - 120, 25)
        self.searchBox.TextChanged += self.OnSearchTextChanged
        self.Controls.Add(self.searchBox)
        
        # Search Mode RadioButtons
        self.radioContains = RadioButton()
        self.radioContains.Text = "Contains (anywhere)"
        self.radioContains.Location = Point(80, 78)
        self.radioContains.AutoSize = True
        self.radioContains.Checked = True
        self.radioContains.CheckedChanged += self.OnSearchModeChanged
        self.Controls.Add(self.radioContains)
        
        self.radioStartsWith = RadioButton()
        self.radioStartsWith.Text = "Starts with"
        self.radioStartsWith.Location = Point(240, 78)
        self.radioStartsWith.AutoSize = True
        self.radioStartsWith.CheckedChanged += self.OnSearchModeChanged
        self.Controls.Add(self.radioStartsWith)
        
        # Result Count Label
        self.resultLabel = Label()
        self.resultLabel.Text = "Unique Parameters: " + str(len(self.all_params))
        self.resultLabel.Location = Point(20, 105)
        self.resultLabel.AutoSize = True
        self.resultLabel.ForeColor = Color.Blue
        self.Controls.Add(self.resultLabel)
        
        # CheckedListBox
        self.checkedListBox = CheckedListBox()
        self.checkedListBox.Location = Point(20, 130)
        self.checkedListBox.Size = Size(screen_width - 60, screen_height - 230)
        self.checkedListBox.CheckOnClick = True
        
        # Load tất cả parameters ban đầu (chỉ load 1 lần)
        for name in self.all_params:
            self.checkedListBox.Items.Add(name, False)
        
        self.Controls.Add(self.checkedListBox)
        
        # Select/Deselect All Buttons
        self.selectAllButton = Button()
        self.selectAllButton.Text = 'Select All'
        self.selectAllButton.Location = Point(screen_width - 260, screen_height - 80)
        self.selectAllButton.Size = Size(100, 40)
        self.selectAllButton.Click += self.OnSelectAll
        self.Controls.Add(self.selectAllButton)
        
        self.deselectAllButton = Button()
        self.deselectAllButton.Text = 'Deselect All'
        self.deselectAllButton.Location = Point(screen_width - 150, screen_height - 80)
        self.deselectAllButton.Size = Size(100, 40)
        self.deselectAllButton.Click += self.OnDeselectAll
        self.Controls.Add(self.deselectAllButton)
        
        # Run Button
        self.runButton = Button()
        self.runButton.Text = 'Delete'
        self.runButton.Location = Point(20, screen_height - 80)
        self.runButton.Size = Size(120, 40)
        self.runButton.DialogResult = DialogResult.OK
        self.Controls.Add(self.runButton)
        
        # Cancel Button
        self.cancelButton = Button()
        self.cancelButton.Text = 'Cancel'
        self.cancelButton.Location = Point(150, screen_height - 80)
        self.cancelButton.Size = Size(120, 40)
        self.cancelButton.DialogResult = DialogResult.Cancel
        self.Controls.Add(self.cancelButton)
        
        self.AcceptButton = self.runButton
        self.CancelButton = self.cancelButton
    
    def OnSelectAll(self, sender, e):
        for i in range(self.checkedListBox.Items.Count):
            self.checkedListBox.SetItemChecked(i, True)
    
    def OnDeselectAll(self, sender, e):
        for i in range(self.checkedListBox.Items.Count):
            self.checkedListBox.SetItemChecked(i, False)
    
    def OnSearchModeChanged(self, sender, e):
        if sender.Checked:
            self.FilterList()
    
    def OnSearchTextChanged(self, sender, e):
        self.FilterList()
    
    def FilterList(self):
        search_text = self.searchBox.Text.lower().strip()
        use_starts_with = self.radioStartsWith.Checked
        
        # Ghi nhớ các item đang được check
        checked_items = set()
        for i in range(self.checkedListBox.Items.Count):
            if self.checkedListBox.GetItemChecked(i):
                checked_items.add(str(self.checkedListBox.Items[i]))
        
        # Clear list và load lại dựa trên filter
        self.checkedListBox.Items.Clear()
        count = 0
        
        for name in self.all_params:
            name_lower = name.lower()
            match = False
            
            if search_text == "":
                match = True
            elif use_starts_with:
                match = name_lower.startswith(search_text)
            else:
                match = search_text in name_lower
            
            if match:
                # Add lại vào UI, và tự động check lại nếu trước đó nó đang được check
                is_checked = name in checked_items
                self.checkedListBox.Items.Add(name, is_checked)
                count += 1
        
        self.resultLabel.Text = "Results: {0} / {1}".format(count, len(self.all_params))


# ----------------- MAIN LOGIC -----------------

# Chỉ lấy các project parameter (loại bỏ built-in)
param_elements = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()

# Sử dụng defaultdict(list) để gom nhóm: Tên -> [Danh sách ID]
param_groups = collections.defaultdict(list)

for pe in param_elements:
    defi = pe.GetDefinition()
    if defi and defi.BuiltInParameter == BuiltInParameter.INVALID:
        name = defi.Name
        param_groups[name].append(pe.Id) # Gom tất cả ID vào cùng 1 Tên

# Lấy danh sách tên độc nhất (các keys của dictionary)
unique_param_names = list(param_groups.keys())

if not unique_param_names:
    OUT = "Không có project parameter nào do người dùng tạo."
else:
    form = DeleteParamForm(unique_param_names)
    if form.ShowDialog() != DialogResult.OK:
        OUT = "Cancelled"
    else:
        # Lấy tên các item đã được check
        selected_names = [str(form.checkedListBox.Items[i]) for i in form.checkedListBox.CheckedIndices]
        
        if not selected_names:
            OUT = "Selected None"
        else:
            deleted_count = 0
            t = Transaction(doc, "Delete Selected Parameters")
            t.Start()
            
            for name in selected_names:
                # Lấy danh sách TẤT CẢ ElementIds thuộc về Tên này
                ids_to_delete = param_groups[name] 
                
                # Duyệt qua từng ID để xóa
                for param_id in ids_to_delete:
                    try:
                        doc.Delete(param_id)
                        deleted_count += 1
                    except Exception:
                        # Bỏ qua nếu parameter bị khóa không cho xóa
                        pass 
            
            t.Commit()
            OUT = "Successfully deleted {} parameter element(s).".format(deleted_count)