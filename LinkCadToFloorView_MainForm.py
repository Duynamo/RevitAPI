import System.Drawing
import System.Windows.Forms

from System.Drawing import *
from System.Windows.Forms import *

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._grb_CadFolder = System.Windows.Forms.GroupBox()
		self._txb_link = System.Windows.Forms.TextBox()
		self._lb_CadFile = System.Windows.Forms.Label()
		self._btt_Browser = System.Windows.Forms.Button()
		self._lb_FVC = System.Windows.Forms.Label()
		self._comboBox1 = System.Windows.Forms.ComboBox()
		self._lb_Level = System.Windows.Forms.Label()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._btt_OK = System.Windows.Forms.Button()
		self._grb_CadFolder.SuspendLayout()
		self.SuspendLayout()
		# 
		# grb_CadFolder
		# 
		self._grb_CadFolder.Controls.Add(self._lb_Level)
		self._grb_CadFolder.Controls.Add(self._comboBox1)
		self._grb_CadFolder.Controls.Add(self._btt_Browser)
		self._grb_CadFolder.Controls.Add(self._lb_CadFile)
		self._grb_CadFolder.Controls.Add(self._txb_link)
		self._grb_CadFolder.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._grb_CadFolder.ForeColor = System.Drawing.Color.Black
		self._grb_CadFolder.Location = System.Drawing.Point(13, 23)
		self._grb_CadFolder.Name = "grb_CadFolder"
		self._grb_CadFolder.Size = System.Drawing.Size(412, 169)
		self._grb_CadFolder.TabIndex = 0
		self._grb_CadFolder.TabStop = False
		self._grb_CadFolder.Text = "Import Option"
		# 
		# txb_link
		# 
		self._txb_link.Location = System.Drawing.Point(108, 52)
		self._txb_link.Name = "txb_link"
		self._txb_link.Size = System.Drawing.Size(294, 27)
		self._txb_link.TabIndex = 0
		self._txb_link.TextChanged += self.Txb_linkTextChanged
		# 
		# lb_CadFile
		# 
		self._lb_CadFile.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_CadFile.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
		self._lb_CadFile.Location = System.Drawing.Point(324, 26)
		self._lb_CadFile.Name = "lb_CadFile"
		self._lb_CadFile.Size = System.Drawing.Size(78, 23)
		self._lb_CadFile.TabIndex = 1
		self._lb_CadFile.Text = "Cad File"
		self._lb_CadFile.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# btt_Browser
		# 
		self._btt_Browser.BackColor = System.Drawing.SystemColors.ActiveCaption
		self._btt_Browser.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_Browser.Location = System.Drawing.Point(9, 56)
		self._btt_Browser.Name = "btt_Browser"
		self._btt_Browser.Size = System.Drawing.Size(89, 23)
		self._btt_Browser.TabIndex = 2
		self._btt_Browser.Text = "Browser"
		self._btt_Browser.UseVisualStyleBackColor = False
		self._btt_Browser.Click += self.Btt_BrowserClick
		# 
		# lb_FVC
		# 
		self._lb_FVC.Location = System.Drawing.Point(12, 245)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(61, 21)
		self._lb_FVC.TabIndex = 1
		self._lb_FVC.Text = "@FVC"
		self._lb_FVC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# comboBox1
		# 
		self._comboBox1.FormattingEnabled = True
		self._comboBox1.Location = System.Drawing.Point(108, 104)
		self._comboBox1.Name = "comboBox1"
		self._comboBox1.Size = System.Drawing.Size(294, 27)
		self._comboBox1.TabIndex = 3
		self._comboBox1.SelectedIndexChanged += self.ComboBox1SelectedIndexChanged
		# 
		# lb_Level
		# 
		self._lb_Level.AllowDrop = True
		self._lb_Level.BackColor = System.Drawing.SystemColors.ActiveCaption
		self._lb_Level.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_Level.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._lb_Level.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_Level.Location = System.Drawing.Point(9, 107)
		self._lb_Level.Name = "lb_Level"
		self._lb_Level.Size = System.Drawing.Size(89, 23)
		self._lb_Level.TabIndex = 4
		self._lb_Level.Text = "FloorPlane"
		self._lb_Level.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		self._lb_Level.Click += self.Lb_LevelClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._btt_CANCLE.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_CANCLE.FlatAppearance.BorderSize = 2
		self._btt_CANCLE.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Red
		self._btt_CANCLE.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(128, 128, 255)
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.Location = System.Drawing.Point(331, 221)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(84, 29)
		self._btt_CANCLE.TabIndex = 2
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# btt_OK
		# 
		self._btt_OK.Cursor = System.Windows.Forms.Cursors.WaitCursor
		self._btt_OK.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_OK.FlatAppearance.BorderSize = 2
		self._btt_OK.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Red
		self._btt_OK.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(128, 128, 255)
		self._btt_OK.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_OK.Location = System.Drawing.Point(261, 221)
		self._btt_OK.Name = "btt_OK"
		self._btt_OK.Size = System.Drawing.Size(64, 29)
		self._btt_OK.TabIndex = 3
		self._btt_OK.Text = "OK"
		self._btt_OK.UseVisualStyleBackColor = True
		self._btt_OK.Click += self.Btt_OKClick
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(437, 275)
		self.Controls.Add(self._btt_OK)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._grb_CadFolder)
		self.Name = "MainForm"
		self.Text = "LinkCadToFloorView"
		self.Load += self.MainFormLoad
		self._grb_CadFolder.ResumeLayout(False)
		self._grb_CadFolder.PerformLayout()
		self.ResumeLayout(False)

	def Btt_BrowserClick(self, sender, e):
		pass

	def Txb_linkTextChanged(self, sender, e):
		pass

	def Lb_LevelClick(self, sender, e):
		pass

	def ComboBox1SelectedIndexChanged(self, sender, e):
		pass

	def Btt_OKClick(self, sender, e):
		pass

	def Btt_CANCLEClick(self, sender, e):
		pass