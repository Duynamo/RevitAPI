import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		resources = System.Resources.ResourceManager("scheduleExport.MainForm", System.Reflection.Assembly.GetEntryAssembly())
		self._groupBox1 = System.Windows.Forms.GroupBox()
		self._groupBox2 = System.Windows.Forms.GroupBox()
		self._cbsSchedule = System.Windows.Forms.CheckedListBox()
		self._checkBox = System.Windows.Forms.CheckBox()
		self._textBoxTotalItem = System.Windows.Forms.TextBox()
		self._label1 = System.Windows.Forms.Label()
		self._label2 = System.Windows.Forms.Label()
		self._textBoxBrowser = System.Windows.Forms.TextBox()
		self._bttBrowser = System.Windows.Forms.Button()
		self._bttOK = System.Windows.Forms.Button()
		self._bttCANCLE = System.Windows.Forms.Button()
		self._label3 = System.Windows.Forms.Label()
		self._groupBox1.SuspendLayout()
		self._groupBox2.SuspendLayout()
		self.SuspendLayout()
		# 
		# groupBox1
		# 
		self._groupBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		self._groupBox1.Controls.Add(self._cbsSchedule)
		self._groupBox1.Font = System.Drawing.Font("Meiryo UI", 8, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._groupBox1.Location = System.Drawing.Point(17, 12)
		self._groupBox1.Name = "groupBox1"
		self._groupBox1.Size = System.Drawing.Size(350, 200)
		self._groupBox1.TabIndex = 0
		self._groupBox1.TabStop = False
		self._groupBox1.Text = "Schedule Table"
		# 
		# groupBox2
		# 
		self._groupBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
		self._groupBox2.Controls.Add(self._bttBrowser)
		self._groupBox2.Controls.Add(self._textBoxBrowser)
		self._groupBox2.Controls.Add(self._label2)
		self._groupBox2.Font = System.Drawing.Font("Meiryo UI", 8)
		self._groupBox2.Location = System.Drawing.Point(17, 263)
		self._groupBox2.Name = "groupBox2"
		self._groupBox2.Size = System.Drawing.Size(350, 100)
		self._groupBox2.TabIndex = 1
		self._groupBox2.TabStop = False
		self._groupBox2.Text = "Export Options"
		# 
		# cbsSchedule
		# 
		self._cbsSchedule.AllowDrop = True
		self._cbsSchedule.Font = System.Drawing.Font("Meiryo UI", 8)
		self._cbsSchedule.FormattingEnabled = True
		self._cbsSchedule.Location = System.Drawing.Point(10, 33)
		self._cbsSchedule.Name = "cbsSchedule"
		self._cbsSchedule.Size = System.Drawing.Size(330, 156)
		self._cbsSchedule.TabIndex = 0
		self._cbsSchedule.SelectedIndexChanged += self.CbsScheduleSelectedIndexChanged
		# 
		# checkBox
		# 
		self._checkBox.Font = System.Drawing.Font("Meiryo UI", 8)
		self._checkBox.Location = System.Drawing.Point(17, 229)
		self._checkBox.Name = "checkBox"
		self._checkBox.Size = System.Drawing.Size(135, 24)
		self._checkBox.TabIndex = 2
		self._checkBox.Text = "select All/None"
		self._checkBox.UseVisualStyleBackColor = True
		self._checkBox.CheckedChanged += self.CheckBoxCheckedChanged
		# 
		# textBoxTotalItem
		# 
		self._textBoxTotalItem.Font = System.Drawing.Font("Meiryo UI", 8)
		self._textBoxTotalItem.Location = System.Drawing.Point(310, 229)
		self._textBoxTotalItem.Name = "textBoxTotalItem"
		self._textBoxTotalItem.Size = System.Drawing.Size(57, 24)
		self._textBoxTotalItem.TabIndex = 3
		self._textBoxTotalItem.TextChanged += self.TextBoxTotalItemTextChanged
		# 
		# label1
		# 
		self._label1.Font = System.Drawing.Font("Meiryo UI", 8)
		self._label1.Location = System.Drawing.Point(214, 232)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(91, 23)
		self._label1.TabIndex = 4
		self._label1.Text = "Total Items: "
		# 
		# label2
		# 
		self._label2.Font = System.Drawing.Font("Meiryo UI", 7)
		self._label2.Location = System.Drawing.Point(10, 29)
		self._label2.Name = "label2"
		self._label2.Size = System.Drawing.Size(125, 23)
		self._label2.TabIndex = 0
		self._label2.Text = "Output Folder"
		# 
		# textBoxBrowser
		# 
		self._textBoxBrowser.Location = System.Drawing.Point(10, 50)
		self._textBoxBrowser.Name = "textBoxBrowser"
		self._textBoxBrowser.Size = System.Drawing.Size(253, 24)
		self._textBoxBrowser.TabIndex = 1
		self._textBoxBrowser.TextChanged += self.TextBoxBrowserTextChanged
		# 
		# bttBrowser
		# 
		self._bttBrowser.BackColor = System.Drawing.Color.White
		self._bttBrowser.Font = System.Drawing.Font("Meiryo UI", 5)
		self._bttBrowser.Location = System.Drawing.Point(269, 50)
		self._bttBrowser.Name = "bttBrowser"
		self._bttBrowser.Size = System.Drawing.Size(75, 23)
		self._bttBrowser.TabIndex = 2
		self._bttBrowser.Text = "Browse"
		self._bttBrowser.UseVisualStyleBackColor = False
		self._bttBrowser.Click += self.BttBrowserClick
		# 
		# bttOK
		# 
		self._bttOK.Location = System.Drawing.Point(200, 381)
		self._bttOK.Name = "bttOK"
		self._bttOK.Size = System.Drawing.Size(75, 29)
		self._bttOK.TabIndex = 5
		self._bttOK.Text = "OK"
		self._bttOK.UseVisualStyleBackColor = True
		self._bttOK.Click += self.BttOKClick
		# 
		# bttCANCLE
		# 
		self._bttCANCLE.Location = System.Drawing.Point(284, 381)
		self._bttCANCLE.Name = "bttCANCLE"
		self._bttCANCLE.Size = System.Drawing.Size(83, 29)
		self._bttCANCLE.TabIndex = 6
		self._bttCANCLE.Text = "CANCLE"
		self._bttCANCLE.UseVisualStyleBackColor = True
		self._bttCANCLE.Click += self.BttCANCLEClick
		# 
		# label3
		# 
		self._label3.Font = System.Drawing.Font("Meiryo UI", 5, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._label3.Location = System.Drawing.Point(5, 406)
		self._label3.Name = "label3"
		self._label3.Size = System.Drawing.Size(40, 14)
		self._label3.TabIndex = 7
		self._label3.Text = "@FVC"
		# 
		# MainForm
		# 
		self.BackColor = System.Drawing.SystemColors.ButtonHighlight
		self.ClientSize = System.Drawing.Size(382, 428)
		self.Controls.Add(self._label3)
		self.Controls.Add(self._bttCANCLE)
		self.Controls.Add(self._bttOK)
		self.Controls.Add(self._label1)
		self.Controls.Add(self._textBoxTotalItem)
		self.Controls.Add(self._checkBox)
		self.Controls.Add(self._groupBox2)
		self.Controls.Add(self._groupBox1)
		self.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self.Icon = resources.GetObject("$this.Icon")
		self.Name = "MainForm"
		self.Text = "scheduleExport"
		self._groupBox1.ResumeLayout(False)
		self._groupBox2.ResumeLayout(False)
		self._groupBox2.PerformLayout()
		self.ResumeLayout(False)
		self.PerformLayout()


	def CbsScheduleSelectedIndexChanged(self, sender, e):
		pass

	def CheckBoxCheckedChanged(self, sender, e):
		pass

	def TextBoxTotalItemTextChanged(self, sender, e):
		pass

	def TextBoxBrowserTextChanged(self, sender, e):
		pass

	def BttBrowserClick(self, sender, e):
		pass

	def BttOKClick(self, sender, e):
		pass

	def BttCANCLEClick(self, sender, e):
		pass