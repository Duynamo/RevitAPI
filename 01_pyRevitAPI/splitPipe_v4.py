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
from Autodesk.Revit.DB.Plumbing import*
clr.AddReference("RevitAPIUI") 
from Autodesk.Revit.UI import*
from Autodesk.Revit.UI.Selection import ISelectionFilter
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
"""_______________________________________________________________________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion
"""_____"""
#region ___def pickPipe
class selectionFilter(ISelectionFilter):
	def __init__(self, category):
		self.category = category
	def AllowElement(self, element):
		if element.Category.Name == self.category:
			return True
		else:
			return False
def pickPipe():
	pipeFilter = selectionFilter("Pipes")
	ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, pipeFilter,"pick Pipe")
	pipe = doc.GetElement(ref.ElementId)
	return pipe
#endregion
#region ___def getPipeLocationCurve
def getPipeLocationCurve(pipe):
    lCurve = pipe.Location.Curve
    return lCurve
#endregion
#region ___def flatten
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList
#endregion
#region ___def devideLineSegment
def divideLineSegment(line, length, startPoint, endPoint):
    points = []
    total_length = line.Length
    direction = (endPoint - startPoint).Normalize()
    current_point = startPoint
    points.append(current_point.ToPoint())
    while (current_point.DistanceTo(startPoint) + length) <= total_length:
        current_point = current_point + direction * length
        points.append(current_point.ToPoint())
    return points
#endregion
def splitPipeAtPoints(doc, pipe, points):
    newPipes = []
    currentPipe = pipe
    originalPipe = pipe
    TransactionManager.Instance.EnsureInTransaction(doc)
    for point in points:
        pipeLocation = currentPipe.Location
        if isinstance(pipeLocation, LocationCurve):
            pipeCurve = pipeLocation.Curve
            if pipeCurve is not None:
                if is_point_on_curve(pipeCurve, point):
                    newPipeIds = PlumbingUtils.BreakCurve(doc, currentPipe.Id, point)
                    if isinstance(newPipeIds, list):
                        newPipe = doc.GetElement(newPipeIds[0])  # Assuming the first element is the new pipe
                        newPipes.append(newPipe)
                        currentPipe = newPipe
                    else:
                        currentPipe = doc.GetElement(newPipeIds)
                else:
                    # Use the original pipe for splitting
                    newPipeIds = PlumbingUtils.BreakCurve(doc, originalPipe.Id, point)
                    if isinstance(newPipeIds, list):
                        newPipe = doc.GetElement(newPipeIds[0])  # Assuming the first element is the new pipe
                        newPipes.append(newPipe)
                        currentPipe = newPipe
                    else:
                        currentPipe = doc.GetElement(newPipeIds)
    TransactionManager.Instance.TransactionTaskDone()
    return newPipes

def splitElementAtPoints(doc, element, points):
    newElements = []
    currentElement = element
    originalElement = element
    TransactionManager.Instance.EnsureInTransaction(doc)

    for point in points:
        location = currentElement.Location
        if isinstance(location, LocationCurve):
            curve = location.Curve
            if curve is not None and is_point_on_curve(curve, point):
                if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeCurves):
                    # Split the pipe at the point
                    newElementIds = PlumbingUtils.BreakCurve(doc, currentElement.Id, point)
                    if isinstance(newElementIds, list):
                        newElement = doc.GetElement(newElementIds[0])
                        newElements.append(newElement)
                        currentElement = newElement
                    else:
                        currentElement = doc.GetElement(newElementIds)
                elif element.Category.Id.IntegerValue == int(BuiltInCategory.OST_GenericModel):
                    # Split the generic model at the point
                    curveStart = curve.GetEndPoint(0)
                    curveEnd = curve.GetEndPoint(1)
                    
                    if curveStart.DistanceTo(point) > 1e-6 and curveEnd.DistanceTo(point) > 1e-6:
                        # Create two new curves for the generic model
                        curve1 = Line.CreateBound(curveStart, point)
                        curve2 = Line.CreateBound(point, curveEnd)
                        
                        # Set the original generic model to the first curve
                        location.Curve = curve1
                        
                        # Duplicate the generic model for the second curve
                        newElementId = ElementTransformUtils.CopyElement(doc, currentElement.Id, XYZ.Zero).ElementAt(0)
                        newElement = doc.GetElement(newElementId)
                        newElement.Location.Curve = curve2
                        newElements.append(newElement)
                        currentElement = newElement
                    else:
                        # Handle the case where the split point is at or near an endpoint
                        newElements.append(currentElement)
        else:
            # Handle elements without a LocationCurve (e.g., FamilyInstance)
            pass


def is_point_on_curve(curve, point):
    projected_point = curve.Project(point)
    return projected_point.Distance < 1e-6
def pickObject():
	ref = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element)
	ele = doc.GetElement(ref.ElementId)
	return  ele

def getElementLocationCurve(element):
    # Check if the element is a pipe
    if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_PipeCurves):
        locationCurve = element.Location.Curve
        return locationCurve

    # Check if the element is a generic model
    elif element.Category.Id.IntegerValue == int(BuiltInCategory.OST_GenericModel):
        if isinstance(element, FamilyInstance):
            location = element.Location
            if hasattr(location, 'Curve'):
                locationCurve = location.Curve
                return locationCurve
        else:
            return None
    else:
        return None
#endregion

class MainForm(Form):
	def __init__(self):
		self.InitializeComponent()
	
	def InitializeComponent(self):
		self._btt_pickPipe = System.Windows.Forms.Button()
		self._grb_sortConn = System.Windows.Forms.GroupBox()
		self._btt_CANCLE = System.Windows.Forms.Button()
		self._btt_SPLIT = System.Windows.Forms.Button()
		self._lb_FVC = System.Windows.Forms.Label()
		self._txb_Length = System.Windows.Forms.TextBox()
		self._lb_Length = System.Windows.Forms.Label()
		self._txb_K = System.Windows.Forms.TextBox()
		self._lb_splitNumber = System.Windows.Forms.Label()
		self._grb_inputData = System.Windows.Forms.GroupBox()
		self._rbt_pX = System.Windows.Forms.RadioButton()
		self._rbt_pY = System.Windows.Forms.RadioButton()
		self._rbt_pZ = System.Windows.Forms.RadioButton()
		self._rbt_sortByMax = System.Windows.Forms.RadioButton()
		self._rbt_sortByMin = System.Windows.Forms.RadioButton()
		self._grb_MinMax = System.Windows.Forms.GroupBox()
		self._grb_sortConn.SuspendLayout()
		self._grb_inputData.SuspendLayout()
		self._grb_MinMax.SuspendLayout()
		self.SuspendLayout()
		# 
		# btt_pickPipe
		# 
		self._btt_pickPipe.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_pickPipe.ForeColor = System.Drawing.Color.Red
		self._btt_pickPipe.Location = System.Drawing.Point(112, 19)
		self._btt_pickPipe.Name = "btt_pickPipe"
		self._btt_pickPipe.Size = System.Drawing.Size(133, 41)
		self._btt_pickPipe.TabIndex = 0
		self._btt_pickPipe.Text = "PICK PIPE"
		self._btt_pickPipe.UseVisualStyleBackColor = True
		self._btt_pickPipe.Click += self.Btt_pickPipeClick
		# 
		# grb_sortConn
		# 
		self._grb_sortConn.Controls.Add(self._rbt_pZ)
		self._grb_sortConn.Controls.Add(self._rbt_pY)
		self._grb_sortConn.Controls.Add(self._rbt_pX)
		self._grb_sortConn.Cursor = System.Windows.Forms.Cursors.Default
		self._grb_sortConn.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._grb_sortConn.Location = System.Drawing.Point(11, 161)
		self._grb_sortConn.Name = "grb_sortConn"
		self._grb_sortConn.Size = System.Drawing.Size(160, 109)
		self._grb_sortConn.TabIndex = 1
		self._grb_sortConn.TabStop = False
		self._grb_sortConn.Text = "Sort Connector by"
		self._grb_sortConn.UseCompatibleTextRendering = True
		# 
		# btt_CANCLE
		# 
		self._btt_CANCLE.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_CANCLE.ForeColor = System.Drawing.Color.Red
		self._btt_CANCLE.Location = System.Drawing.Point(265, 293)
		self._btt_CANCLE.Name = "btt_CANCLE"
		self._btt_CANCLE.Size = System.Drawing.Size(101, 37)
		self._btt_CANCLE.TabIndex = 0
		self._btt_CANCLE.Text = "CANCLE"
		self._btt_CANCLE.UseVisualStyleBackColor = True
		self._btt_CANCLE.Click += self.Btt_CANCLEClick
		# 
		# btt_SPLIT
		# 
		self._btt_SPLIT.Font = System.Drawing.Font("Meiryo UI", 10.2, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._btt_SPLIT.ForeColor = System.Drawing.Color.Red
		self._btt_SPLIT.Location = System.Drawing.Point(158, 293)
		self._btt_SPLIT.Name = "btt_SPLIT"
		self._btt_SPLIT.Size = System.Drawing.Size(101, 37)
		self._btt_SPLIT.TabIndex = 0
		self._btt_SPLIT.Text = "SPLIT"
		self._btt_SPLIT.UseVisualStyleBackColor = True
		self._btt_SPLIT.Click += self.Btt_SPLITClick
		# 
		# lb_FVC
		# 
		self._lb_FVC.Font = System.Drawing.Font("Meiryo UI", 4.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_FVC.ForeColor = System.Drawing.Color.Blue
		self._lb_FVC.Location = System.Drawing.Point(18, 313)
		self._lb_FVC.Name = "lb_FVC"
		self._lb_FVC.Size = System.Drawing.Size(56, 20)
		self._lb_FVC.TabIndex = 2
		self._lb_FVC.Text = "@FVC"
		self._lb_FVC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# txb_Length
		# 
		self._txb_Length.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_Length.Location = System.Drawing.Point(20, 49)
		self._txb_Length.Name = "txb_Length"
		self._txb_Length.Size = System.Drawing.Size(133, 23)
		self._txb_Length.TabIndex = 3
		# self._txb_Length.TextChanged += self.Txb_LengthTextChanged
		self._txb_Length.Text = '3000'
		# 
		# lb_Length
		# 
		self._lb_Length.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_Length.Location = System.Drawing.Point(14, 26)
		self._lb_Length.Name = "lb_Length"
		self._lb_Length.Size = System.Drawing.Size(117, 20)
		self._lb_Length.TabIndex = 2
		self._lb_Length.Text = "Length(mm):"
		self._lb_Length.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# txb_K
		# 
		self._txb_K.Font = System.Drawing.Font("Meiryo UI", 7.20000029, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._txb_K.Location = System.Drawing.Point(199, 49)
		self._txb_K.Name = "txb_K"
		self._txb_K.Size = System.Drawing.Size(133, 23)
		self._txb_K.TabIndex = 5
		self._txb_K.TextChanged += self.Txb_KTextChanged
		self._txb_K.Text = '1000'
		# 
		# lb_splitNumber
		# 
		self._lb_splitNumber.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._lb_splitNumber.Location = System.Drawing.Point(192, 26)
		self._lb_splitNumber.Name = "lb_splitNumber"
		self._lb_splitNumber.Size = System.Drawing.Size(36, 20)
		self._lb_splitNumber.TabIndex = 4
		self._lb_splitNumber.Text = "K:"
		self._lb_splitNumber.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# grb_inputData
		# 
		self._grb_inputData.Controls.Add(self._lb_Length)
		self._grb_inputData.Controls.Add(self._txb_K)
		self._grb_inputData.Controls.Add(self._txb_Length)
		self._grb_inputData.Controls.Add(self._lb_splitNumber)
		self._grb_inputData.Cursor = System.Windows.Forms.Cursors.Default
		self._grb_inputData.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._grb_inputData.Location = System.Drawing.Point(6, 66)
		self._grb_inputData.Name = "grb_inputData"
		self._grb_inputData.RightToLeft = System.Windows.Forms.RightToLeft.No
		self._grb_inputData.Size = System.Drawing.Size(351, 82)
		self._grb_inputData.TabIndex = 6
		self._grb_inputData.TabStop = False
		self._grb_inputData.Text = "Input"
		# 
		# rbt_pX
		# 
		self._rbt_pX.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._rbt_pX.ForeColor = System.Drawing.Color.Red
		self._rbt_pX.Location = System.Drawing.Point(25, 24)
		self._rbt_pX.Name = "rbt_pX"
		self._rbt_pX.Size = System.Drawing.Size(61, 24)
		self._rbt_pX.TabIndex = 7
		self._rbt_pX.TabStop = True
		self._rbt_pX.Text = "pX"
		self._rbt_pX.UseVisualStyleBackColor = True
		self._rbt_pX.CheckedChanged += self.Rbt_pXCheckedChanged
		self._rbt_pX.Checked = True
		# 
		# rbt_pY
		# 
		self._rbt_pY.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._rbt_pY.ForeColor = System.Drawing.Color.Red
		self._rbt_pY.Location = System.Drawing.Point(25, 53)
		self._rbt_pY.Name = "rbt_pY"
		self._rbt_pY.Size = System.Drawing.Size(61, 24)
		self._rbt_pY.TabIndex = 7
		self._rbt_pY.TabStop = True
		self._rbt_pY.Text = "pY"
		self._rbt_pY.UseVisualStyleBackColor = True
		self._rbt_pY.CheckedChanged += self.Rbt_pYCheckedChanged
		# 
		# rbt_pZ
		# 
		self._rbt_pZ.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._rbt_pZ.ForeColor = System.Drawing.Color.Red
		self._rbt_pZ.Location = System.Drawing.Point(25, 81)
		self._rbt_pZ.Name = "rbt_pZ"
		self._rbt_pZ.Size = System.Drawing.Size(61, 24)
		self._rbt_pZ.TabIndex = 7
		self._rbt_pZ.TabStop = True
		self._rbt_pZ.Text = "pZ"
		self._rbt_pZ.UseVisualStyleBackColor = True
		self._rbt_pZ.CheckedChanged += self.Rbt_pZCheckedChanged
		# 
		# rbt_sortByMax
		# 
		self._rbt_sortByMax.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._rbt_sortByMax.ForeColor = System.Drawing.Color.Fuchsia
		self._rbt_sortByMax.Location = System.Drawing.Point(6, 73)
		self._rbt_sortByMax.Name = "rbt_sortByMax"
		self._rbt_sortByMax.Size = System.Drawing.Size(132, 24)
		self._rbt_sortByMax.TabIndex = 7
		self._rbt_sortByMax.TabStop = True
		self._rbt_sortByMax.Text = "Sort By Max?"
		self._rbt_sortByMax.UseVisualStyleBackColor = True
		self._rbt_sortByMax.CheckedChanged += self.Rbt_sortByMaxCheckedChanged
		# 
		# rbt_sortByMin
		# 
		self._rbt_sortByMin.Font = System.Drawing.Font("Meiryo UI", 7.8, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 128)
		self._rbt_sortByMin.ForeColor = System.Drawing.Color.Fuchsia
		self._rbt_sortByMin.Location = System.Drawing.Point(7, 43)
		self._rbt_sortByMin.Name = "rbt_sortByMin"
		self._rbt_sortByMin.Size = System.Drawing.Size(132, 24)
		self._rbt_sortByMin.TabIndex = 7
		self._rbt_sortByMin.TabStop = True
		self._rbt_sortByMin.Text = "Sort By Min?"
		self._rbt_sortByMin.UseVisualStyleBackColor = True
		self._rbt_sortByMin.CheckedChanged += self.Rbt_sortByMaxCheckedChanged
		self._rbt_sortByMin.Checked = True
		# 
		# grb_MinMax
		# 
		self._grb_MinMax.Controls.Add(self._rbt_sortByMin)
		self._grb_MinMax.Controls.Add(self._rbt_sortByMax)
		self._grb_MinMax.Font = System.Drawing.Font("Meiryo UI", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 128)
		self._grb_MinMax.Location = System.Drawing.Point(198, 161)
		self._grb_MinMax.Name = "grb_MinMax"
		self._grb_MinMax.Size = System.Drawing.Size(159, 109)
		self._grb_MinMax.TabIndex = 8
		self._grb_MinMax.TabStop = False
		self._grb_MinMax.Text = "MinMax"
		# 
		# MainForm
		# 
		self.ClientSize = System.Drawing.Size(369, 342)
		self.Controls.Add(self._grb_MinMax)
		self.Controls.Add(self._grb_inputData)
		self.Controls.Add(self._lb_FVC)
		self.Controls.Add(self._grb_sortConn)
		self.Controls.Add(self._btt_SPLIT)
		self.Controls.Add(self._btt_CANCLE)
		self.Controls.Add(self._btt_pickPipe)
		self.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
		self.Name = "MainForm"
		self.Text = "Split Pipe"
		self.TopMost = True
		self.Load += self.MainFormLoad
		self._grb_sortConn.ResumeLayout(False)
		self._grb_inputData.ResumeLayout(False)
		self._grb_inputData.PerformLayout()
		self._grb_MinMax.ResumeLayout(False)
		self.ResumeLayout(False)


	def Btt_pickPipeClick(self, sender, e):
		_pipe = pickObject()
		self.selPipe = _pipe

	def Btt_SPLITClick(self, sender, e):
		'''__________'''
		pipe = self.selPipe
		new_dynPoints = []
		try:
			if pipe is not None:
				splitNumber1 = self._txb_K.Text
				if splitNumber1.strip():
					try:
						splitNumber = int(splitNumber1)
					except ValueError:
						splitNumber = None
				else:
					splitNumber = None
				'''__________'''
				splitLength1 = self._txb_Length.Text
				if splitLength1.strip():
					try:
						splitLength = int(splitLength1)/304.8
					except ValueError:
						splitLength = None
				else:
					splitLength = None
				'''___'''
				if hasattr(pipe, 'Pipes'):
					locationCurve = pipe.Location.Curve
				if hasattr(pipe, 'GenericModel'):
					locationCurve = pipe.GenericModel.Location



				pipeCurve  = getElementLocationCurve(pipe)
				originConns = []
				originConns.append(pipeCurve.GetEndPoint(0))
				originConns.append(pipeCurve.GetEndPoint(1))

				# conns = list(pipe.ConnectorManager.Connectors.GetEnumerator())
				# originConns = list(c.Origin for c in conns)
				sortByCor_case1 = self._rbt_pX.Checked
				sortByCor_case2 = self._rbt_pY.Checked
				sortByCor_case3 = self._rbt_pZ.Checked
				'''___'''
				minCase = self._rbt_sortByMin.Checked
				maxCase = self._rbt_sortByMax.Checked
				'''___'''
				if sortByCor_case1 == True and minCase == True :
					sortConns = sorted(originConns, key=lambda c : c.X)			
				elif sortByCor_case1 == True and maxCase == True :
					sortConns = sorted(originConns, key=lambda c : c.X)
					sortConns.reverse()
				elif sortByCor_case2 == True and minCase == True :
					sortConns = sorted(originConns, key=lambda c : c.Y)			
				elif sortByCor_case2 == True and maxCase == True :
					sortConns = sorted(originConns, key=lambda c : c.Y)
					sortConns.reverse()
				elif sortByCor_case3 == True and minCase == True :
					sortConns = sorted(originConns, key=lambda c : c.Z)			
				elif sortByCor_case3 == True and maxCase == True :
					sortConns = sorted(originConns, key=lambda c : c.Z)
					sortConns.reverse()		
				points = divideLineSegment(pipeCurve, splitLength, sortConns[0], sortConns[1])
				
				dynPoints = list(c.ToRevitType() for c in points)
				if splitNumber <= len(dynPoints):
					new_dynPoints = dynPoints[1:splitNumber+1]
				else:
					new_dynPoints = dynPoints[1:]
		except Exception as e:
			TransactionManager.Instance.ForceCloseTransaction()
			pass
		TransactionManager.Instance.EnsureInTransaction(doc)
		newPipes = splitElementAtPoints(doc, pipe, new_dynPoints)
		TransactionManager.Instance.TransactionTaskDone
		pass

	def Btt_CANCLEClick(self, sender, e):
		self.Close()
		pass

	def MainFormLoad(self, sender, e):
		pass

	def Txb_KTextChanged(self, sender, e):
		self.K = float(self._txb_K.Text)
		pass
	def Txb_LengthTextChanged(self, sender, e):
		self.length = float(self._txb_Length.Text)
		pass	

	def Cbb_sortConnectorBySelectedIndexChanged(self, sender, e):
		pass

	def Rbt_pXCheckedChanged(self, sender, e):
		pass

	def Rbt_pYCheckedChanged(self, sender, e):
		pass

	def Rbt_pZCheckedChanged(self, sender, e):
		pass

	def Rbt_sortByMinCheckedChanged(self, sender, e):
		pass	
	def Rbt_sortByMaxCheckedChanged(self, sender, e):
		pass
f = MainForm()
Application.Run(f)