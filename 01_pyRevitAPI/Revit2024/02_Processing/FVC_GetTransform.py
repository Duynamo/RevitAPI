import clr
import sys 
import System   
import math

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*

clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI import*
from Autodesk.Revit.UI.Selection import*
from Autodesk.Revit.UI.Selection import ObjectType

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
"""_______________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
"""_______________________________________________________________"""
# Select objects
def pickObjects():
    refs = uidoc.Selection.PickObjects(ObjectType.Element)
    return  [doc.GetElement(i.ElementId) for i in refs]
"""_______________________________________________________________"""
def calculateRotation(R):
# Calculate the rotation angle
    theta = math.acos((R[0][0] + R[1][1] + R[2][2] - 1) / 2)
    angle = math.degrees(theta)
# Calculate the rotation axis
    if angle == 0:
        # If the angle is 0, default rotation axis is [0, 0, 1]
		axis = [0, 0, 1]
    else:
		# If the angle is not 0, calculate the rotation axis
        sin_theta = math.sin(theta)
        axis = [-(R[2][1] - R[1][2]) / (2 * sin_theta),
                -(R[0][2] - R[2][0]) / (2 * sin_theta),
                -(R[1][0] - R[0][1]) / (2 * sin_theta)]
        
    return axis, angle

"""_______________________________________________________________"""
class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._btt_pickElement = System.Windows.Forms.Button()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._grb_originInformation = System.Windows.Forms.GroupBox()
		self._lb_thetaAngle = System.Windows.Forms.Label()
		self._lb_origin = System.Windows.Forms.Label()
		self._grb_Angle = System.Windows.Forms.GroupBox()
		self._lb_theta = System.Windows.Forms.Label()
		self._txb_originX = System.Windows.Forms.TextBox()
		self._txb_originZ = System.Windows.Forms.TextBox()
		self._txb_theta = System.Windows.Forms.TextBox()
		self._lb_XValue = System.Windows.Forms.Label()
		self._lb_YValue = System.Windows.Forms.Label()
		self._lb_ZValue = System.Windows.Forms.Label()
		self._lb_FVC = System.Windows.Forms.Label()
		self._txb_originY = System.Windows.Forms.TextBox()
		self._txb_Ux = System.Windows.Forms.TextBox()
		self._txb_Uy = System.Windows.Forms.TextBox()
		self._txb_Uz = System.Windows.Forms.TextBox()
		self._label3 = System.Windows.Forms.Label()
		self._label2 = System.Windows.Forms.Label()
		self._label1 = System.Windows.Forms.Label()
		self._grb_originInformation.SuspendLayout()
		self._grb_Angle.SuspendLayout()
		self.SuspendLayout()
		# 
		# btt_pickElement
		# 
		self._btt_pickElement.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_pickElement.FlatAppearance.BorderSize = 5
		self._btt_pickElement.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue
		self._btt_pickElement.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow
		self._btt_pickElement.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_pickElement.ForeColor = System.Drawing.Color.Red
		self._btt_pickElement.Location = System.Drawing.Point(152, 12)
		self._btt_pickElement.Name = "btt_pickElement"
		self._btt_pickElement.Size = System.Drawing.Size(194, 33)
		self._btt_pickElement.TabIndex = 0
		self._btt_pickElement.Text = "Pick Element"
		self._btt_pickElement.UseVisualStyleBackColor = True
		self._btt_pickElement.Click += self.Btt_pickElementClick
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Cursor = System.Windows.Forms.Cursors.AppStarting
		self._btt_CANCLE.FlatAppearance.BorderColor = System.Drawing.Color.Red
		self._btt_CANCLE.FlatAppearance.BorderSize = 5
		self._btt_CANCLE.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Blue
		self._btt_CANCLE.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Yellow
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
		self._btt_CANCLE.Location = System.Drawing.Point(355, 417)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(114, 33)
		self._btt_CANCLE.TabIndex = 1
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# grb_originInformation
		# 
		self._grb_originInformation.Controls.Add(self._lb_ZValue)
		self._grb_originInformation.Controls.Add(self._lb_YValue)
		self._grb_originInformation.Controls.Add(self._lb_XValue)
		self._grb_originInformation.Controls.Add(self._txb_originZ)
		self._grb_originInformation.Controls.Add(self._txb_originY)
		self._grb_originInformation.Controls.Add(self._txb_originX)
		self._grb_originInformation.Controls.Add(self._lb_origin)
		self._grb_originInformation.Location = System.Drawing.Point(24, 72)
		self._grb_originInformation.Name = "grb_originInformation"
		self._grb_originInformation.Size = System.Drawing.Size(445, 129)
		self._grb_originInformation.TabIndex = 2
		self._grb_originInformation.TabStop = False
		self._grb_originInformation.Text = "Origin Information"
		# 
		# lb_thetaAngle
		# 
		self._lb_thetaAngle.Location = System.Drawing.Point(0, 0)
		self._lb_thetaAngle.Name = "lb_thetaAngle"
		self._lb_thetaAngle.Size = System.Drawing.Size(100, 23)
		self._lb_thetaAngle.TabIndex = 6
		# 
		# lb_origin
		# 
		self._lb_origin.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_origin.ForeColor = System.Drawing.Color.Red
		self._lb_origin.Location = System.Drawing.Point(9, 72)
		self._lb_origin.Name = "lb_origin"
		self._lb_origin.Size = System.Drawing.Size(58, 23)
		self._lb_origin.TabIndex = 1
		self._lb_origin.Text = "Origin"
		# 
		# grb_Angle
		# 
		self._grb_Angle.Controls.Add(self._txb_Uz)
		self._grb_Angle.Controls.Add(self._txb_Uy)
		self._grb_Angle.Controls.Add(self._txb_Ux)
		self._grb_Angle.Controls.Add(self._txb_theta)
		self._grb_Angle.Controls.Add(self._label1)
		self._grb_Angle.Controls.Add(self._label2)
		self._grb_Angle.Controls.Add(self._label3)
		self._grb_Angle.Controls.Add(self._lb_theta)
		self._grb_Angle.Location = System.Drawing.Point(24, 221)
		self._grb_Angle.Name = "grb_Angle"
		self._grb_Angle.Size = System.Drawing.Size(445, 184)
		self._grb_Angle.TabIndex = 5
		self._grb_Angle.TabStop = False
		self._grb_Angle.Text = "Angle"
		# 
		# lb_theta
		# 
		self._lb_theta.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_theta.ForeColor = System.Drawing.Color.Red
		self._lb_theta.Location = System.Drawing.Point(9, 23)
		self._lb_theta.Name = "lb_theta"
		self._lb_theta.Size = System.Drawing.Size(58, 23)
		self._lb_theta.TabIndex = 1
		self._lb_theta.Text = "Theta"
		# 
		# txb_originX
		# 
		self._txb_originX.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_originX.ForeColor = System.Drawing.Color.Red
		self._txb_originX.Location = System.Drawing.Point(85, 67)
		self._txb_originX.Name = "txb_originX"
		self._txb_originX.ScrollBars = System.Windows.Forms.ScrollBars.Both
		self._txb_originX.Size = System.Drawing.Size(100, 27)
		self._txb_originX.TabIndex = 5
		self._txb_originX.TextChanged += self.Txb_originXTextChanged
		# 
		# txb_originZ
		# 
		self._txb_originZ.ForeColor = System.Drawing.Color.Red
		self._txb_originZ.Location = System.Drawing.Point(331, 67)
		self._txb_originZ.Name = "txb_originZ"
		self._txb_originZ.Size = System.Drawing.Size(100, 27)
		self._txb_originZ.TabIndex = 5
		self._txb_originZ.TextChanged += self.Txb_originZTextChanged
		# 
		# txb_theta
		# 
		self._txb_theta.ForeColor = System.Drawing.Color.Red
		self._txb_theta.Location = System.Drawing.Point(85, 19)
		self._txb_theta.Name = "txb_theta"
		self._txb_theta.Size = System.Drawing.Size(346, 27)
		self._txb_theta.TabIndex = 15
		self._txb_theta.TextChanged += self.Txb_thetaTextChanged
		# 
		# lb_XValue
		# 
		self._lb_XValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_XValue.ForeColor = System.Drawing.Color.Red
		self._lb_XValue.Location = System.Drawing.Point(102, 31)
		self._lb_XValue.Name = "lb_XValue"
		self._lb_XValue.Size = System.Drawing.Size(66, 23)
		self._lb_XValue.TabIndex = 15
		self._lb_XValue.Text = "X Value"
		# 
		# lb_YValue
		# 
		self._lb_YValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_YValue.ForeColor = System.Drawing.Color.Red
		self._lb_YValue.Location = System.Drawing.Point(224, 31)
		self._lb_YValue.Name = "lb_YValue"
		self._lb_YValue.Size = System.Drawing.Size(66, 23)
		self._lb_YValue.TabIndex = 16
		self._lb_YValue.Text = "Y Value"
		# 
		# lb_ZValue
		# 
		self._lb_ZValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_ZValue.ForeColor = System.Drawing.Color.Red
		self._lb_ZValue.Location = System.Drawing.Point(346, 31)
		self._lb_ZValue.Name = "lb_ZValue"
		self._lb_ZValue.Size = System.Drawing.Size(66, 23)
		self._lb_ZValue.TabIndex = 17
		self._lb_ZValue.Text = "Z Value"
		# 
		# lb_FVC
		# 
		self._lb_FVC.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._lb_FVC.Font = System.Drawing.Font("Meiryo UI", 7, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_FVC.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
		self._lb_FVC.Location = System.Drawing.Point(12, 446)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(58, 23)
		self._lb_FVC.TabIndex = 16
		self._lb_FVC.Text = "@FVC"
		# 
		# txb_originY
		# 
		self._txb_originY.ForeColor = System.Drawing.Color.Red
		self._txb_originY.Location = System.Drawing.Point(207, 67)
		self._txb_originY.Name = "txb_originY"
		self._txb_originY.Size = System.Drawing.Size(100, 27)
		self._txb_originY.TabIndex = 5
		self._txb_originY.TextChanged += self.Txb_originYTextChanged
		# 
		# txb_Ux
		# 
		self._txb_Ux.ForeColor = System.Drawing.Color.Red
		self._txb_Ux.Location = System.Drawing.Point(85, 60)
		self._txb_Ux.Name = "txb_Ux"
		self._txb_Ux.Size = System.Drawing.Size(346, 27)
		self._txb_Ux.TabIndex = 16
		self._txb_Ux.TextChanged += self.Txb_UxTextChanged
		# 
		# txb_Uy
		# 
		self._txb_Uy.ForeColor = System.Drawing.Color.Red
		self._txb_Uy.Location = System.Drawing.Point(85, 100)
		self._txb_Uy.Name = "txb_Uy"
		self._txb_Uy.Size = System.Drawing.Size(346, 27)
		self._txb_Uy.TabIndex = 17
		self._txb_Uy.TextChanged += self.Txb_UyTextChanged
		# 
		# txb_Uz
		# 
		self._txb_Uz.ForeColor = System.Drawing.Color.Red
		self._txb_Uz.Location = System.Drawing.Point(85, 140)
		self._txb_Uz.Name = "txb_Uz"
		self._txb_Uz.Size = System.Drawing.Size(346, 27)
		self._txb_Uz.TabIndex = 18
		self._txb_Uz.TextChanged += self.Txb_UzTextChanged
		# 
		# label3
		# 
		self._label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._label3.ForeColor = System.Drawing.Color.Red
		self._label3.Location = System.Drawing.Point(9, 64)
		self._label3.Name = "label3"
		self._label3.Size = System.Drawing.Size(58, 23)
		self._label3.TabIndex = 2
		self._label3.Text = "Ux"
		# 
		# label2
		# 
		self._label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._label2.ForeColor = System.Drawing.Color.Red
		self._label2.Location = System.Drawing.Point(9, 104)
		self._label2.Name = "label2"
		self._label2.Size = System.Drawing.Size(58, 23)
		self._label2.TabIndex = 3
		self._label2.Text = "Uy"
		# 
		# label1
		# 
		self._label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		self._label1.ForeColor = System.Drawing.Color.Red
		self._label1.Location = System.Drawing.Point(9, 145)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(58, 23)
		self._label1.TabIndex = 4
		self._label1.Text = "Uz"
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(494, 479)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._grb_Angle)
		self.Controls.Add(self._grb_originInformation)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_pickElement)
		self.Controls.Add(self._lb_thetaAngle)
		self.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.Text = "CheckModelOrigin"
		self.TopMost = True
		self._grb_originInformation.ResumeLayout(False)
		self._grb_originInformation.PerformLayout()
		self._grb_Angle.ResumeLayout(False)
		self._grb_Angle.PerformLayout()
		self.ResumeLayout(False)


	def Btt_pickElementClick(self, sender, e):
		# Get selected elements
		desElements = pickObjects()
		for ele in desElements:
			# Get total transform of the element
			transform = ele.GetTotalTransform()
			transScale = transform.Scale
			# Calculate transformed origin coordinates
			transOriginX = transform.Origin.X / transScale
			transOriginY = transform.Origin.Y / transScale
			transOriginZ = transform.Origin.Z / transScale

			# Calculate scaled basis vectors
			asicX_scaled = (transform.BasisX.X / transScale, transform.BasisX.Y / transScale, transform.BasisX.Z / transScale)
			basicY_scaled = (transform.BasisY.X / transScale, transform.BasisY.Y / transScale, transform.BasisY.Z / transScale)
			basicZ_scaled = (transform.BasisZ.X / transScale, transform.BasisZ.Y / transScale, transform.BasisZ.Z / transScale)
			rotation_matrix = [list(basicX_scaled), list(basicY_scaled), list(basicZ_scaled)]
			# Calculate rotation axis and angle
			rotationAxisTheta = calculateRotation(rotation_matrix)
			rotation_angle = rotationAxisTheta[1]
			
			self._txb_originX.Text = str(transOriginX)
			self._txb_originY.Text = str(transOriginY)
			self._txb_originZ.Text = str(transOriginZ)

			self._txb_theta.Text = str(rotation_angle)
			self._txb_Ux.Text = str(rotationAxisTheta[0][0])
			self._txb_Uy.Text = str(rotationAxisTheta[0][1])
			self._txb_Uz.Text = str(rotationAxisTheta[0][2])


		pass

	def Txb_originXTextChanged(self, sender, e):
		pass

	def Txb_originYTextChanged(self, sender, e):
		pass

	def Txb_originZTextChanged(self, sender, e):
		pass

	def Txb_thetaTextChanged(self, sender, e):
		pass

	def Txb_UxTextChanged(self, sender, e):
		pass

	def Txb_UyTextChanged(self, sender, e):
		pass

	def Txb_UzTextChanged(self, sender, e):
		pass

	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass

f = MainForm()
Application.Run(f)