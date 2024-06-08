#Copyright (c) Data Shapes,  2020
#Data-Shapes www.data-shapes.io , elayoubi.mostafa@data-shapes.io @data_shapes
	
import clr
import sys
pyt_path = r'C:\Program Files (x86)\IronPython 2.7\Lib'
sys.path.append(pyt_path)
import os
import webbrowser
import unicodedata
import io
import tempfile
import System
	
try:
	clr.AddReference('System.Windows.Forms')
	clr.AddReference('System.Drawing')
	clr.AddReference('System.Windows.Forms.DataVisualization')
	
	from System.Drawing import Point , Size , Graphics, Bitmap, Image, Font, FontStyle, Icon, Color, Region , Rectangle , ContentAlignment , SystemFonts, FontFamily
	from System.Windows.Forms import Application, DockStyle, Button, Form, Label, TrackBar , ToolTip, ColumnHeader, TextBox, CheckBox, FolderBrowserDialog, OpenFileDialog, DialogResult, ComboBox, FormBorderStyle, FormStartPosition, ListView, ListViewItem , SortOrder, Panel, ImageLayout, GroupBox, RadioButton, BorderStyle, PictureBox, PictureBoxSizeMode, LinkLabel, CheckState, ColumnHeaderStyle , ImageList, VScrollBar, DataGridView, DataGridViewSelectionMode, DataGridViewAutoSizeColumnsMode , DataGridViewClipboardCopyMode , TreeView , TreeNode , TreeNodeCollection , AutoScaleMode , Screen, Padding
	from System.Windows.Forms.DataVisualization.Charting import *#Chart , SeriesChartType
	from System.Collections.Generic import *
	from System.Collections.Generic import List as iList
	from System.Windows.Forms import View as vi
	clr.AddReference('System')
	from System import IntPtr , Char
	from System import Type as SType, IO
	from System import Array
	from System.ComponentModel import Container
	clr.AddReference('System.Data')
	from System.Data import DataTable , DataView

	try: #try to import All Revit dependencies
		clr.AddReference('RevitAPIUI')
		from  Autodesk.Revit.UI import Selection , TaskDialog 
		from  Autodesk.Revit.UI.Selection import ISelectionFilter
		clr.AddReference('RevitNodes')
		import Revit
		clr.ImportExtensions(Revit.Elements)
		clr.ImportExtensions(Revit.GeometryConversion)
		
		clr.AddReference('RevitServices')
		from RevitServices.Persistence import DocumentManager
		doc = DocumentManager.Instance.CurrentDBDocument
		uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
	
		clr.AddReference('RevitAPI')
		try:
			from Autodesk.Revit.DB import ImageImportOptions	
		except:
			from Autodesk.Revit.DB import ImageTypeOptions , ImageType, ImagePlacementOptions , ImageInstance
		from Autodesk.Revit.DB import FilteredElementCollector , Transaction, View , ViewType , ViewFamily, ViewDrafting, ViewFamilyType, Element, ElementId , FamilyInstance , Document , XYZ, BoxPlacement, UnitUtils
		
		try:
			from Autodesk.Revit.DB import  UnitType
		except:
			from Autodesk.Revit.DB import SpecTypeId

		dbviews = [v for v in FilteredElementCollector(doc).OfClass(View).ToElements() if (v.ViewType == ViewType.FloorPlan or v.ViewType == ViewType.CeilingPlan or v.ViewType == ViewType.Section or v.ViewType == ViewType.Elevation or v.ViewType == ViewType.ThreeD)]
		viewindex = 0
		try:
			UIunit = Document.GetUnits(doc).GetFormatOptions(UnitType.UT_Length).DisplayUnits
		except:
			UIunit = Document.GetUnits(doc).GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()
			
		class selectionfilter(ISelectionFilter):
			def __init__(self,category):
				self.category = category
			def AllowElement(self,element):
				if element.Category.Name in [c.Name for c in self.category]:
					return True
				else:
					return False
			def AllowReference(reference,point):
				return False


			
	except: #in case we are in the Sandbox, Formit or Civil 3D environment
		pass
	
	importcolorselection = 0
	
	try:
		from  Autodesk.Revit.UI import ColorSelectionDialog
	except:
		importcolorselection = 1
	

	
	clr.AddReference('ProtoGeometry')
	from Autodesk.DesignScript.Geometry import Point as dsPoint

	from System.Reflection import Assembly
	import xml.etree.ElementTree as et
	

	
	import re	
	def regexEndNum(input):
		try:
			return 	re.search('(\d+)$', input).group(0)
		except:
			return ""

	def iterateThroughNodes(collection,li):
		if hasattr(collection,'Nodes'):
			ntest = collection.Nodes
			if len(ntest) > 0:
				for i in ntest:
					iterateThroughNodes(i,li)
			else:
				if collection.Checked:
					li.append(collection.Tag)
		return li
		

	
	class MultiTextBoxForm(Form):
		
	    def __init__(self):
	        self.Text = 'Data-Shapes | Multi Input UI ++'
	        self.output = []
	        self.values = []
	        self.cancelled = False
	
	    def setclose(self, sender, event):
	    	cbindexread = 0
	    	if sender.Name != "Cancel":
	    		for f in self.output:
	    			if f.GetType() == myTextBox:
	    				if f._isNum :
	    					val = float(f.Text)
	    				else:
	    					val = f.Text
	    				self.values.append(val)
	    			if f.GetType() == CheckBox:
	    				self.values.append(f.Checked)
	    			if f.GetType() == Button:
	    				if isinstance(f.Tag ,list):
	    					try:
	    						self.values.append([e for e in f.Tag if e.__class__.__name__ != "Category"])	    					
	    					except:
	    						self.values.append(f.Tag)	    					
	    				else:
	    					try:	    				
	    						if f.Tag.__class__.__name__ != "Category":
	    							self.values.append(f.Tag)
	    						else:
    								self.values.append([])
	    					except:
	    						self.values.append(f.Tag)	
	    			if f.GetType() == ComboBox:
	    				try:
	    					key = f.Text
	    					self.values.append(f.Tag[key])
	    				except:
	    					self.values.append(None)
	    			if f.GetType() == mylistview:
	    				self.values.append([f.Values[i.Text] for i in f.CheckedItems])
	    			if f.GetType() == mytrackbar:
	    				self.values.append(f.startval+f.Value*f.step)
	    			if f.GetType() == mygroupbox:
	    				try:
	    					key = [j.Text for j in f.Controls if j.Checked == True][0]
	    					self.values.append(f.Tag[key])
	    				except:
	    					self.values.append(None)
	    			if f.GetType() == myDataGridView:
	    				f.EndEdit()
	    				dsrc = f.DataSource
	    				out = []
	    				colcount = f.ColumnCount
	    				rowcount = f.RowCount - 1
	    				if f.Tag:
		    				l = []
	    					for i in range(colcount):	    						
	    						l.append(dsrc.Columns[i].ColumnName)
	    					out.append(l)	    						
		    				for r in range(rowcount):
		    					l = []
		    					for i in range(colcount):
		    						l.append(dsrc.DefaultView[r].Row[i])
		    					out.append(l)
	    				else: 
		    				for r in range(rowcount):
		    					l = []
		    					for i in range(colcount):
		    						l.append(dsrc.DefaultView[r].Row[i])
		    					out.append(l)
	    				self.values.append(out)
	    			if f.GetType() == TreeView:
	    				ls = []
	    				nds = f.Nodes[0]
	    				iterateThroughNodes(nds,ls)
	    				self.values.append(ls)
	    			if f.GetType() == GroupBox:
	    				rb = [c for c in f.Controls if c.GetType() == RadioButton and c.Checked][0]
	    				self.values.append(rb.Text)
	    				f.Controls.Remove(rb)
	    	else:
	    		self.values = None
	    		self.cancelled = True
	    	self.Close()
	
	    def reset(self, sender, event):
			pass
	
	    def openfile(self, sender, event):
			ofd = OpenFileDialog()
			dr = ofd.ShowDialog()
			if dr == DialogResult.OK:
				sender.Text = ofd.FileName
				sender.Tag = ofd.FileName

	    def exportToExcel(self, sender, event):
	    	#importing Excel IronPython libraries
	    	clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
	    	from Microsoft.Office.Interop import Excel
	    	ex = Excel.ApplicationClass()
	    	ex.Visible = sender.Tag[1]
	    	ex.DisplayAlerts = False
	    	fbd = FolderBrowserDialog()
	    	fbd.SelectedPath = sender.Text
	    	parent = sender.Parent
	    	fptextbox = parent.GetChildAtPoint(Point(parent.Location.X,sender.Location.Y+5*yRatio))
	    	dataGrid = parent.GetChildAtPoint(Point(parent.Location.X,parent.Location.Y+23*xRatio))
	    	dataTable = dataGrid.DataSource
	    	fptext = fptextbox.Text
	    	titletext = parent.GetChildAtPoint(Point(0,0)).Text
	    	dr = fbd.ShowDialog()
	    	frstRwTtle = sender.Tag[0]
	    	if frstRwTtle:
	    		_header = Excel.XlYesNoGuess.xlYes
	    	else:
	    		_header = Excel.XlYesNoGuess.xlNo
	    	if dr == DialogResult.OK:
	    		workbk = ex.Workbooks.Add()
	    		worksheet = workbk.Worksheets.Add()
	    		#Writing title and doc info
	    		if sender.Tag[2]:
		    		titleCell = worksheet.Cells[1,1]
		    		worksheet.Cells[2,1].Value2 = sender.Tag[3]
		    		titleCell.Value2 = titletext
		    		titleCell.Font.Size = 18
		    		titleCell.Font.Bold = True
		    		startR = 3
		    		endR = 3
		    	else:
		    		startR = 1
		    		endR = 0	
		    	if frstRwTtle:
		    		for j in range(0,dataTable.Columns.Count):
		    			worksheet.Cells[startR,j+1] = dataTable.Columns[j].ColumnName		    	
		    		for i in range(0,dataTable.Rows.Count):
		    			for j in range(0,dataTable.Columns.Count):
			    			worksheet.Cells[i+startR+1,j+1] = dataTable.DefaultView[i].Row[j].ToString()
		    		xlrange = ex.get_Range(worksheet.Cells[startR,1],worksheet.Cells[dataTable.Rows.Count+endR+1,dataTable.Columns.Count])			    			
		    	else :
		    		for i in range(0,dataTable.Rows.Count):
		    			for j in range(0,dataTable.Columns.Count):
			    			worksheet.Cells[i+startR,j+1] = dataTable.DefaultView[i].Row[j].ToString()		    	
		    		xlrange = ex.get_Range(worksheet.Cells[startR,1],worksheet.Cells[dataTable.Rows.Count+endR,dataTable.Columns.Count])
		    	xlrange.Columns.AutoFit()
		    	worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, xlrange, SType.Missing, _header, SType.Missing).Name = "DataShapesTable"
		    	worksheet.ListObjects["DataShapesTable"].TableStyle =  "TableStyleMedium16"
	    		workbk.SaveAs(fbd.SelectedPath + "\\" + fptext)
	    		if not sender.Tag[1]:
	    			workbk.Close()
	    			ex.Quit()

	    def startCell(self, sender, event ):
			sender.startcell["X"] = event.ColumnIndex
			sender.startcell["Y"] = event.RowIndex	

	    def endCell(self, sender, event ):
	    	try:
				sender.endcell["X"] = event.ColumnIndex
				sender.endcell["Y"] = event.RowIndex 
				startval = sender.Rows[sender.startcell["Y"]].Cells[sender.startcell["X"]].Value
				endNum = regexEndNum(startval)
				if endNum != "":
					if sender.endcell["Y"] == sender.startcell["Y"]:
						for e,i in enumerate(range(sender.startcell["X"],sender.endcell["X"] + 1)):
							sender.Rows[sender.startcell["Y"]].Cells[i].Value = startval[:-len(endNum)] + str(int(endNum) + e)
					elif sender.endcell["X"] == sender.startcell["X"]:
						for e,i in enumerate(range(sender.startcell["Y"],sender.endcell["Y"] + 1)):
							sender.Rows[i].Cells[sender.endcell["X"]].Value = startval[:-len(endNum)] + str(int(endNum) + e)
				else:				
					if sender.endcell["Y"] == sender.startcell["Y"]:
						for i in range(sender.startcell["X"],sender.endcell["X"] + 1):
							sender.Rows[sender.startcell["Y"]].Cells[i].Value = startval
					elif sender.endcell["X"] == sender.startcell["X"]:
						for i in range(sender.startcell["Y"],sender.endcell["Y"] + 1):
							sender.Rows[i].Cells[sender.endcell["X"]].Value = startval
	    	except:
	    		pass
	    		    	
	    def startRowDrag(self, sender, event ):
	    	shmak
			
	    def opendirectory(self, sender, event):
			fbd = FolderBrowserDialog()
			fbd.SelectedPath = sender.Text
			dr = fbd.ShowDialog()
			if dr == DialogResult.OK:
				sender.Text = fbd.SelectedPath
				sender.Tag = fbd.SelectedPath
	
	    def pickobjects(self, sender, event):   
			for c in self.Controls:
				c.Enabled = False
			try:
				sel = uidoc.Selection.PickObjects(Selection.ObjectType.Element,'')
				selelem = [doc.GetElement(s.ElementId) for s in sel]
				sender.Tag = (selelem)
			except:
				pass
			for c in self.Controls:
				c.Enabled = True
		#THIS METHOD IS FOR CIVIL 3D EVIRONMENT
	    def pickautocadobjects(self, sender, event):   
			selelem = []	
			for c in self.Controls:
				c.Enabled = False
			try:
				acadDoc = System.Runtime.InteropServices.Marshal.GetActiveObject("Autocad.Application").ActiveDocument
				acadDoc.Activate()
				acadUser = acadDoc.GetVariable("users5")	
				acadDoc.SendCommand("(and(princ\042"+ sender.Text + "\042)(setq ss(ssget))(setvar\042users5\042\042LinkDWGUIOK\042)(command\042_.Select\042ss\042\042)) ")
				selection_ = acadDoc.ActiveSelectionSet
				acadDoc.SendCommand("(setq ss nil) ")
				if acadDoc.GetVariable("users5") == "LinkDWGUIOK" and selection_ != None:
					for sel in selection_:				
						selelem.append(sel)		
					acadDoc.SetVariable("users5", acadUser)
				sender.Tag = list(selelem)		
			except:
				pass
			for c in self.Controls:
				c.Enabled = True	

	    def pickautocadobject(self, sender, event):   
			selelem = None	
			for c in self.Controls:
				c.Enabled = False
			try:
				acadDoc = System.Runtime.InteropServices.Marshal.GetActiveObject("Autocad.Application").ActiveDocument
				acadUser = acadDoc.GetVariable("users5")
				acadPickBox = acadDoc.GetVariable("pickbox")
				acadDoc.SetVariable("pickbox", 5)
				acadDoc.Activate()
				acadDoc.SendCommand("(setq obj(car(entsel\042" + sender.Text + "\042))) ")
				acadDoc.SendCommand("(and obj(setvar\042users5\042(cdr(assoc 5(entget obj))))(setq obj nil)) ")		
				selection_ = acadDoc.GetVariable("users5")
				acadDoc.SetVariable("pickbox", acadPickBox)
				acadDoc.SetVariable("users5", acadUser)
				selelem = acadDoc.HandleToObject(selection_)
				sender.Tag = selelem		
			except:
				pass
			for c in self.Controls:
				c.Enabled = True	

	    def pickobjectsordered(self, sender, event):
			for c in self.Controls:
				c.Enabled = False
			output = []
			test = True
			TaskDialog.Show("Data|Shapes", 'Pick elements in order, then hit ESC to exit.')
			while test:
				try:
					sel = doc.GetElement(uidoc.Selection.PickObject(Selection.ObjectType.Element, 'Pick elements in order').ElementId)
					output.append(sel.ToDSType(True))
				except : 
					test = False
				sender.Tag = output
			for c in self.Controls:
				c.Enabled = True
	    
	    def pickobjectsofcatordered(self, sender, event):
			for c in self.Controls:
				c.Enabled = False
			output = []
			test = True
			if isinstance(sender.Tag,list):			
				category = UnwrapElement(sender.Tag)
			else:
				category = [UnwrapElement(sender.Tag)]
			TaskDialog.Show("Data|Shapes", 'Select %s in order, then press ESC to exit.' %(', '.join([c.Name for c in category])))
			while test:
				try:
					selfilt = selectionfilter(category)
					sel = doc.GetElement(uidoc.Selection.PickObject(Selection.ObjectType.Element,selfilt, 'Select %s' %(', '.join([c.Name for c in category]))).ElementId)				
					output.append(sel.ToDSType(True))
				except : 
					test = False
				sender.Tag = (output)
			for c in self.Controls:
				c.Enabled = True
			
	    def picklinkedobjects(self, sender, event):
			#This part was made easier by Dimitar Venkov's work
			for c in self.Controls:
				c.Enabled = False
			try:
				linkref = uidoc.Selection.PickObject(Selection.ObjectType.Element,'Select the link instance.')
				link = doc.GetElement(linkref.ElementId).GetLinkDocument()
				td = TaskDialog.Show('Data-Shapes','Select the linked elements and press Finish.')
				sel = uidoc.Selection.PickObjects(Selection.ObjectType.LinkedElement,'Select the linked elements and press Finish.')
				selelem = [link.GetElement(s.LinkedElementId) for s in sel]
				sender.Tag = (selelem)
			except:
				pass
			for c in self.Controls:
				c.Enabled = True		

	    def pickobject(self, sender, event):
			for c in self.Controls:
				c.Enabled = False
			try:
				sel = uidoc.Selection.PickObject(Selection.ObjectType.Element,'')
				selelem = doc.GetElement(sel.ElementId) 
				sender.Tag = (selelem)
			except:
				pass
			for c in self.Controls:
				c.Enabled = True
			
	    def picklinkedobject(self, sender, event):
			#This part was made easier by Dimitar Venkov's work
			for c in self.Controls:
				c.Enabled = False
			try:
				linkref = uidoc.Selection.PickObject(Selection.ObjectType.Element,'Select the link instance.')
				link = doc.GetElement(linkref.ElementId).GetLinkDocument()
				td = TaskDialog.Show('Data-Shapes','Select the linked element.')
				sel = uidoc.Selection.PickObject(Selection.ObjectType.LinkedElement,'Select the linked element.')
				selelem = link.GetElement(sel.LinkedElementId)
				sender.Tag = (selelem)
			except:
				pass				
			for c in self.Controls:
				c.Enabled = True
			
	    def pickobjectsofcat(self, sender, event):
			for c in self.Controls:
				c.Enabled = False
			if isinstance(sender.Tag,list):	    
				category = UnwrapElement(sender.Tag)
			else:
				category = [UnwrapElement(sender.Tag)]
			try:
				selfilt = selectionfilter(category)
				sel = uidoc.Selection.PickObjects(Selection.ObjectType.Element,selfilt,'Select %s' %(', '.join([c.Name for c in category])))
				selelem = [doc.GetElement(s.ElementId) for s in sel]
				sender.Tag = (selelem)
			except:
				pass
			for c in self.Controls:
				c.Enabled = True

	    def pickobjectofcat(self, sender, event):
			for c in self.Controls:
				c.Enabled = False
			if isinstance(sender.Tag,list):	    
				category = UnwrapElement(sender.Tag)
			else:
				category = [UnwrapElement(sender.Tag)]
			try:
				selfilt = selectionfilter(category)
				sel = uidoc.Selection.PickObject(Selection.ObjectType.Element,selfilt,'Select %s' %(', '.join([c.Name for c in category])))
				selelem = doc.GetElement(sel.ElementId) 
				sender.Tag = (selelem)
			except:
				pass
			for c in self.Controls:
				c.Enabled = True
			
	    def pickfaces(self, sender, event):
			faces = []	    	
			for c in self.Controls:
				c.Enabled = False
			try:
				selface = uidoc.Selection.PickObjects(Selection.ObjectType.Face,'')
				for s in selface:
					f = uidoc.Document.GetElement(s).GetGeometryObjectFromReference(s).ToProtoType(True)
					[i.Tags.AddTag("RevitFaceReference", s) for i in f]
					faces.append(f)
				sender.Tag = [i for j in faces for i in j]
			except:
				pass
			for c in self.Controls:
				c.Enabled = True
				
	    def pickpointsonface(self, sender, event):
			faces = []	    	
			for c in self.Controls:
				c.Enabled = False
			selpoints = uidoc.Selection.PickObjects(Selection.ObjectType.PointOnElement,'')
			points = []
			for s in selpoints:
				pt = s.GlobalPoint
				points.append(dsPoint.ByCoordinates(UnitUtils.ConvertFromInternalUnits(pt.X,UIunit),UnitUtils.ConvertFromInternalUnits(pt.Y,UIunit),UnitUtils.ConvertFromInternalUnits(pt.Z,UIunit)))
			sender.Tag = points
			for c in self.Controls:
				c.Enabled = True
				
	    def pickedges(self, sender, event):
			edges = []
			for c in self.Controls:
				c.Enabled = False	
			try:				
				seledge = uidoc.Selection.PickObjects(Selection.ObjectType.Edge,'')
				for s in seledge:
					e = uidoc.Document.GetElement(s).GetGeometryObjectFromReference(s).AsCurve().ToProtoType(True)
					e.Tags.AddTag("RevitFaceReference", s)
					edges.append(e)
				sender.Tag = edges
			except:
				pass
			for c in self.Controls:
				c.Enabled = True

	    def colorpicker(self, sender, event):
			dialog = ColorSelectionDialog()
			selection = ColorSelectionDialog.Show(dialog)
			selected = dialog.SelectedColor
			sender.Tag = selected
			sender.BackColor = Color.FromArgb(selected.Red,selected.Green,selected.Blue)
			sender.ForeColor = Color.FromArgb(selected.Red,selected.Green,selected.Blue)
	
	    def topmost(self):
			self.TopMost = True
	
	    def lvadd(self, sender, event):
			sender.Tag = [i for i in sender.CheckedItems]
			
	    def scroll(self, sender, event):
			parent = sender.Parent
			child = parent.GetChildAtPoint(Point(0,5*yRatio))
			child.Text = str(sender.startval+sender.Value*sender.step)

	    def openurl(self, sender, event):
			webbrowser.open(sender.Tag)
	
	    def selectall(self, sender, event):
			if sender.Checked:
				parent = sender.Parent
				listview = parent.GetChildAtPoint(Point(0,0))
				for i in listview.Items:
					i.Checked = True
			else:
				pass
				
	    def selectnone(self, sender, event):
			if sender.Checked:
				parent = sender.Parent
				listview = parent.GetChildAtPoint(Point(0,0))
				for i in listview.Items:
					i.Checked = False
			else:
				pass		

	    def updateallnone(self, sender, event):
	    	try:
	    		parent = sender.Parent
	    		rball = parent.GetChildAtPoint(Point(0,sender.Height + 5*yRatio))
	    		rbnone = parent.GetChildAtPoint(Point(80 * xRatio,sender.Height + 5*yRatio))
	    		if sender.CheckedItems.Count == 0 and event.NewValue == CheckState.Unchecked:
	    			rbnone.Checked = False
	    			rball.Checked = False
	    		elif sender.CheckedItems.Count == sender.Items.Count and event.NewValue == CheckState.Unchecked:
	    			rball.Checked = False
	    			rbnone.Checked = False
	    		elif sender.CheckedItems.Count == sender.Items.Count-1 and event.NewValue == CheckState.Checked:
	    			rball.Checked = True
	    			rbnone.Checked = False
	    		elif sender.CheckedItems.Count == 1 and event.NewValue == CheckState.Unchecked:
	    			rball.Checked = False
	    			rbnone.Checked = True	    	
	    		else :
	    			rball.Checked = False
	    			rbnone.Checked = False
	    	except:
	    		pass

	    def zoomcenter(self, sender, event ):
			if event.X > 15:	    
				try:
					element = doc.GetElement(uidoc.Selection.GetElementIds()[0])
					uidoc.ShowElements(element)
				except:
					pass
			else:
				pass
				
			
	    def setviewforelement(self, sender, event ):    
			if event.X > 15*xRatio:	    		
				try:
					item = sender.GetItemAt(event.X,event.Y).Text
					element = UnwrapElement(sender.Values[item])
					try:
						viewsforelement = [v for v in dbviews if (not v.IsTemplate) and (element.Id in [e.Id for e in FilteredElementCollector(doc,v.Id).OfClass(element.__class__).ToElements()])]
					except:
						viewsforelement = [v for v in dbviews if (not v.IsTemplate) and (element.Id in [e.Id for e in FilteredElementCollector(doc,v.Id).OfClass(FamilyInstance).ToElements()])]
					global viewindex
					dbView = viewsforelement[viewindex]
					id = [element.Id]
					icollection = iList[ElementId](id)
					uidoc.Selection.SetElementIds(icollection)
				except:
					pass
			else:	    
				pass


	    def CheckChildren(self, sender, event ):
			evNode = event.Node	    
			checkState = evNode.Checked	
			for n in event.Node.Nodes:	   	
				n.Checked = checkState			
				
	    def ActivateOption(self, sender, event ):
	    	parent = sender.Parent
	    	associatedControls = [p for p in parent.Controls if p.Name == sender.Text and p.GetType() == Panel][0]
	    	restofcontrols = [p for p in parent.Controls if p.Name != sender.Text and p.GetType() == Panel]
	    	if sender.Checked:
	    		associatedControls.Enabled = True
	    		for c in restofcontrols:
	    			c.Enabled = False
	    		parent.Tag = sender.Text
	    		
	    def showtooltip(self, sender, event ):
	    	ttp = ToolTip()
	    	ttp.AutoPopDelay = 10000
	    	ttp.SetToolTip(sender , sender.Tag)	

	    def numsOnly(self, sender, event ):
	    	if Char.IsDigit(event.KeyChar)==False and event.KeyChar != "." and Char.IsControl(event.KeyChar)==False:
	    		event.Handled = True
	    
	    def chart_showLabels(self, sender, event):
			cb = sender
			panelcht = sender.Parent
			chart1 = panelcht.GetChildAtPoint(Point(0,0))
			for s in chart1.Series:
				if s.ChartType == SeriesChartType.Pie:
					if cb.Checked:
						s["PieLabelStyle"] = "Inside"
					else:
						s["PieLabelStyle"] = "Disabled"
				else:
					if cb.Checked:
						s.IsValueShownAsLabel = True
					else:
						s.IsValueShownAsLabel = False
	    		
	    def imageexport(self, sender, event):
	    	import datetime
	    	from datetime import datetime
	    	from RevitServices.Transactions import TransactionManager
	    	#Modify resolution before the render
	    	fontFam = FontFamily("Segoe UI Symbol")
	    	originalFont = Font(fontFam,8)
	    	panelcht = sender.Parent
	    	chart1 = panelcht.GetChildAtPoint(Point(0,0))
	    	originalTitleFont = chart1.Titles[0].Font
	    	originalWidth = chart1.Width
	    	originalHeight = chart1.Height
	    	chart1.Visible = False
	    	chart1.Dock = DockStyle.None
	    	chart1.Width = 2100 * 0.8
	    	chart1.Height = 1500 * 0.8
	    	chart1.ChartAreas[0].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.None
	    	chart1.ChartAreas[0].AxisY.LabelAutoFitStyle = LabelAutoFitStyles.None
	    	chart1.ChartAreas[0].AxisX.LabelStyle.Font = Font(fontFam, 30)
	    	chart1.ChartAreas[0].AxisY.LabelStyle.Font = Font(fontFam, 30)
	    	chart1.ChartAreas[0].AxisX.TitleFont = Font(fontFam, 30)
	    	chart1.ChartAreas[0].AxisY.TitleFont = Font(fontFam, 30)
	    	chart1.TextAntiAliasingQuality = TextAntiAliasingQuality.High
	    	chart1.BackColor = Color.White
	    	chart1.Titles[0].Font = Font(fontFam, 32, FontStyle.Bold)
	    	chart1.ChartAreas[0].BackColor = Color.White
	    	for serie in chart1.Series:
	    		serie.Font = Font(fontFam, 30)
		    	for p in serie.Points:
		    		p.Font = Font(fontFam, 30)
		    		p.MarkerSize = 15
	    	for legend in chart1.Legends:
	    		legend.Font = Font(fontFam, 30)
	    		legend.BackColor = Color.White
	    	chart1.Invalidate()
	    	chart1.SaveImage(tempfile.gettempdir() + "\\chartImage.bmp", ChartImageFormat.Bmp)
	    	#Get back to original settings
	    	chart1.Width = originalWidth
	    	chart1.Height = originalHeight
	    	chart1.BackColor = Color.Transparent
	    	chart1.ChartAreas[0].BackColor = Color.Transparent
	    	chart1.ChartAreas[0].AxisX.LabelStyle.Font = originalFont
	    	chart1.ChartAreas[0].AxisY.LabelStyle.Font = originalFont
	    	chart1.ChartAreas[0].AxisX.TitleFont = originalFont
	    	chart1.ChartAreas[0].AxisY.TitleFont = originalFont
	    	chart1.Titles[0].Font = originalTitleFont
	    	for serie in chart1.Series:
	    		serie.Font = originalFont
	    		for p in serie.Points:
	    			p.Font = originalFont
		    		p.MarkerSize = 8
	    	for legend in chart1.Legends:
	    		legend.Font = originalFont
	    		legend.BackColor = Color.Transparent
	    	chart1.Invalidate()
	    	chart1.Visible = True
	    	#Import the picture in a Drafting View
	    	#Import the picture in a Drafting View // The try catch if for handling the fact that ImageImportOptions was deprecated in 2020 and is obsolete in 2021	    	
	    	try:
	    		importOptions = ImageImportOptions()	
	    		importOptions.Resolution = 72
	    		importOptions.Placement = BoxPlacement.TopLeft	    		
	    	except:
	    		imageTypeOption = ImageTypeOptions()	
	    		imageTypeOption.Resolution = 72

	    	collector = FilteredElementCollector(doc).OfClass(ViewFamilyType)
	    	viewFamilyTypes = []
	    	for c in collector:
	    		if c.ViewFamily == ViewFamily.Drafting:
	    			viewFamilyTypes.append(c)
	    	viewFamilyType = viewFamilyTypes[0]
	    	TransactionManager.Instance.EnsureInTransaction(doc)
	    	draftView = ViewDrafting.Create(doc,viewFamilyType.Id)
	    	draftView.Name = chart1.Titles[0].Text + datetime.now().strftime(" (%m/%d/%Y, %H.%M.%S)")
	    	imagePath = tempfile.gettempdir() + "\\chartImage.bmp"
	    	newElement = clr.StrongBox[Element]()
	    	try:
	    		doc.Import(imagePath,importOptions,draftView,newElement)	    	
	    	except:
	    		imageTypeOption.SetPath(imagePath)
	    		imageType = ImageType.Create(doc,imageTypeOption)
	    		placementOptions = ImagePlacementOptions(XYZ(0,0,0),BoxPlacement.TopLeft)
	    		ImageInstance.Create(doc,draftView,imageType.Id,placementOptions)	    		
	    	TransactionManager.Instance.TransactionTaskDone()
	    		    	
	    def chart_showLegend(self, sender, event ):
	    	cb = sender
	    	panelcht = sender.Parent
	    	chart1 = panelcht.GetChildAtPoint(Point(0,0))
	    	if len(chart1.Legends) <= 1:
	    		for legend in chart1.Legends:
	    			if cb.Checked:
	    				legend.Enabled = True
	    			else:
	    				legend.Enabled = False
	    	else:
	    		if cb.Checked:
	    			chart1.Legends[1].Enabled = True
	    		else:
	    			chart1.Legends[1].Enabled = False
		    		

	class mylistview(ListView):
	
		def __init__(self):
			self.Values = []

	class mytrackbar(TrackBar):
	
		def __init__(self,startval,step):
			self.startval = startval
			self.step = step

	class myDataGridView(DataGridView):
	
		def __init__(self):
			self.startcell = {}
			self.endcell = {}
			
	class mygroupbox(GroupBox):
	
		def __init__(self):
			self.Values = []
			
	class myTextBox(TextBox):
	
		def __init__(self):
			self._isNum = False
	
			
	#Form initialization
	
	form = MultiTextBoxForm()
	xRatio = Screen.PrimaryScreen.Bounds.Width/1920
	if xRatio == 0:
		xRatio = 1
	yRatio = Screen.PrimaryScreen.Bounds.Height/1080
	if yRatio == 0:
		yRatio = 1
	form.topmost()	
	form.ControlBox = True
	xlabel = 25 * xRatio
	xinput = 150 * xRatio
	formy = 10 * yRatio
	if IN[8] * xRatio > (350 * xRatio):	formwidth = IN[8] * xRatio
	else: formwidth = 350 * xRatio
	fields = []
	error = 0
	
	#Description 
	
	if IN[3] != "":
		des = Label()
		des.Location = Point(xlabel,formy)
		des.Font = Font("Arial", 15,FontStyle.Bold)		
		des.AutoSize = True
		des.MaximumSize = Size(formwidth - (2 * xlabel)*xRatio,0)
		des.Text = IN[3]
		form.Controls.Add(des)
		formy = des.Bottom + (15*xRatio)
	formheaderheight = formy
	
	#Input form
	
	# Create a container panel for all inputs
	body = Panel()
	body.Location = Point(0,formy)
	body.Width = formwidth - 15*xRatio
	
	
	# Process form inputs
	if isinstance(IN[0],list):
		inputtypes = IN[0]
	else:
		inputtypes = [IN[0]]
	# This definition is to handle the sorting of special characters
	def remove_accents(input_str):
	    nfkd_form = unicodedata.normalize('NFKD', input_str)
	    only_ascii = nfkd_form.encode('ASCII', 'ignore')
	    return only_ascii	

	#Adding Logo 
	#default logo in case no input	
	def getImageByName(ImName):
		dynamo = Assembly.Load('DynamoCore')
		version = str(dynamo.GetName().Version)[:3]		
		dynPath = os.getenv('APPDATA')+"\\Dynamo\Dynamo Revit\\" + version 
		root = et.parse(dynPath + "\DynamoSettings.xml").getroot()
		logopaths = []
		for child in root:
			if child.tag == "CustomPackageFolders":
				for path in child:
					logopaths.append(path.text + "\packages\Data-Shapes\extra\\" + ImName)
					logopaths.append(path.text + "\Data-Shapes\extra\\" + ImName)
		deflogopath = ""
		for path in logopaths:
			if deflogopath == "":
				if os.path.isfile(path):
					deflogopath = path
					try:
						ima = Image.FromFile(deflogopath)
						bmp = Bitmap.FromFile(deflogopath)
						return ima,bmp
					except: 
						pass
						
	try:
		if IN[4] != '':
			try:
				ima = Image.FromFile(IN[4])
				bmp = Bitmap.FromFile(IN[4])
			except:
				ima = IN[4]
				bmp = IN[4]

		else :
			_ims = getImageByName("a.png")
			ima = _ims[0]
			bmp = _ims[1]
			
		logo = Panel()
		if IN[6] == None:
			xlogo = 20 * xRatio
			ylogo = formy+ 10*yRatio
		else:
			xlogo = 30 * xRatio
			ylogo = formy
		size = 110 * xRatio
		logo = PictureBox()
		ratio = (ima.Height)/(ima.Width)
		h = float(ima.Height)
		w = float(ima.Width)
		ratio = h/w
		logo.Size = Size(size,size*ratio)
		scaledimage = Bitmap(size,size*ratio)
		gr = Graphics.FromImage(scaledimage)
		gr.DrawImage(ima,0,0,size,size*ratio)
		logo.Image = ima
		logo.SizeMode = PictureBoxSizeMode.Zoom
		form.Controls.Add(logo)
		logo.Location = Point(xlogo,ylogo)
		
		
		#Setting icon 
		thumb = bmp.GetThumbnailImage(64, 64, bmp.GetThumbnailImageAbort,IntPtr.Zero)
		thumb.MakeTransparent();
		icon = Icon.FromHandle(thumb.GetHicon())
		form.Icon = icon

	except :
		logo = Panel()
		logo.Width = 110 * xRatio
		logo.Height = 110 * yRatio
	
	def addinput(formbody,inputs,starty,xinput,rightmargin,labelsize,importcolorselection):
		y = starty
		for j in inputs:
			label = Label()
			label.Location = Point(xlabel,y+4*yRatio)
			label.AutoSize = True
			label.MaximumSize = Size(labelsize,0)
			if j.__class__.__name__ == 'listview' and j.element_count > 0:
				label.Text = j.inputname + '\n(' + str(j.element_count) + ' element' + ("s" if j.element_count > 1 else "") + ')'
			else:
				try:
					label.Text = j.inputname
				except:
					pass
			formbody.Controls.Add(label)
	
			if j.__class__.__name__ == 'dropdown':
				cb = ComboBox()
				if j.inputname != "":
					cb.Width = formwidth - rightmargin - xinput
					cb.Location = Point(xinput,y)
				else:
					cb.Width = formwidth - rightmargin - xlabel
					cb.Location = Point(xlabel,y)
				cb.Sorted = j._sorted
				[cb.Items.Add(i) for i in j.keys() if not (i == 'inputname' or i == 'height' or i == 'defaultvalue' or i == 'highlight' or i == '_sorted' )]
				cb.Tag = j
				if j.defaultvalue != None:
					defindex = [i for i in cb.Items].index([i for i in j.keys() if not (i == 'inputname' or i == 'height' or i == 'defaultvalue' or i == 'highlight' or i == '_sorted' )][j.defaultvalue])
					cb.SelectedIndex = defindex
				formbody.Controls.Add(cb)
				form.output.append(cb)
				y = label.Bottom + 25 * yRatio
			#CHARTS 
			if j.__class__.__name__ == 'uipiechart':					
				labels = j.xseries
				inputcolors = j.inputcolors
				pyList = []
				xseries = [j.xseries]
				yseries = [j.yseries]
				nbOfSeries = 1
				if inputcolors is not None:
					colorBool = True
					for color in inputcolors:
						pyList.append(Color.FromArgb(color.Red,color.Green,color.Blue))
					wfColor = Array[Color](pyList)
				else:
					colorBool = False
				#Set a global Font code
				fontFam = FontFamily("Segoe UI Symbol")
				originalFont = Font(fontFam,8)
				#Chart created and Panel to host the chart
				chart1 = Chart()
				panelChart = Panel()
				panelChart.Name = "panelChart"
				#Panels colors
				chart1.BackColor = formbody.BackColor
				panelChart.BackColor = Color.Transparent
				#PanelChart location
				panelChart.Width = formwidth  - rightmargin - xlabel
				panelChart.Location = Point(xlabel,y + 30)
				#Chart colors palette
				chart1.Palette =  ChartColorPalette.None
				if colorBool:
					chart1.PaletteCustomColors = wfColor
				else:
					chart1.Palette =  ChartColorPalette.None
				chart1.Series.Clear()
				#Series created
				label = SmartLabelStyle()
				def createSeries(seriesName,_keys,_values):
					srs = Series()
					srs["PieLabelStyle"] = "Inside"
					srs.Name = seriesName
					srs.ChartType = SeriesChartType.Pie
					srs.ToolTip = "Percent: #PERCENT"
					srs.IsValueShownAsLabel = True
					for i,j in zip(_keys,_values):
						srs.Points.AddXY(i,j)
					return srs
				#Create series
				series = []
				for i in range(0,len(xseries)):
					series.append(createSeries(labels[i],xseries[i],yseries[i]))
				#Add series to chart
				for s in series:
					chart1.Series.Add(s)
				#Legend text of the points
				for s in series:
					for p in s.Points: 
						p.Font = originalFont
				#Refresh the Chart
				chart1.Invalidate()
				#Create a ChartArea and add it to the chart
				chartArea1 = ChartArea()
				chartArea1.BackColor = formbody.BackColor
				chartArea1.Name = "ChartArea1"
				chart1.ChartAreas.Add(chartArea1)				
				#Create a Legend and add it to the chart
				legend1 = Legend()
				legend1.BackColor = formbody.BackColor
				legend1.Font = originalFont
				legend1.Name = "Legend1"
				legend1.IsTextAutoFit = True
				legend1.LegendStyle = LegendStyle.Column
				chart1.Legends.Add(legend1)
				#Initialize the chart and its properties
				chart1.BeginInit()
				chart1.Size = Size(panelChart.Width,panelChart.Width)
				chart1.AntiAliasing = AntiAliasingStyles.All
				fontFam = FontFamily("Segoe UI Symbol")
				titleFont = Font(fontFam,10,FontStyle.Bold)
				title = Title()
				title.Text = j.chartname
				title.Alignment = ContentAlignment.TopLeft
				title.Font = titleFont
				chart1.Titles.Add(title)
				#Add a button to push on view
				bt = Button()
				bt.Tag = "Push the chart on a view in Revit"
				bt.Text = "Push on a draft view"
				bt.Height = 20 * yRatio
				bt.Width = 140 * yRatio				
				panelChart.Controls.Add(bt)
				bt.BringToFront()
				bt.Font = originalFont
				bt.Click += form.imageexport
				#Add a checkBox to show/hide legend
				checkBox1 = CheckBox()
				checkBox1.AutoCheck = True
				checkBox1.Name = "checkBox1"
				checkBox1.Text = "Legend"
				panelChart.Controls.Add(checkBox1)
				checkBox1.BringToFront()
				checkBox1.Checked = True
				checkBox1.Font = originalFont
				checkBox1.AutoSize = True
				checkBox1.Click += form.chart_showLegend
				#Add a checkBox to show/hide labels
				checkBox2 = CheckBox()
				checkBox2.AutoCheck = True
				checkBox2.Name = "checkBox2"
				checkBox2.Text = "Labels"
				panelChart.Controls.Add(checkBox2)
				checkBox2.BringToFront()
				checkBox2.Checked = True
				checkBox2.Font = originalFont
				checkBox2.AutoSize = True
				checkBox2.Click += form.chart_showLabels
				
				panelChart.Size = Size(panelChart.Width, panelChart.Width + checkBox2.Height + 30)
				#Add chart to panel
				panelChart.Controls.Add(chart1)
				chart1.Location = Point(0,0)
				bt.Location = Point(panelChart.Width - bt.Width, panelChart.Width + 15)
				checkBox1.Location = Point(panelChart.Width - bt.Width - 80, panelChart.Width + 15)
				checkBox2.Location = Point(panelChart.Width - bt.Width - 160, panelChart.Width + 15)
				formbody.Controls.Add(panelChart)
				
				y = panelChart.Bottom + 25 * yRatio
				
			elif j.__class__.__name__ == 'uibarchart':
				if isinstance(j.labels,list):
					labels = j.labels
				else:
					labels = [j.labels]
				inputcolors = j.inputcolors
				horizont = j.ishorizontal
				pyList = []
				if isinstance(j.xseries[0],list) and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.xseries)
					xseries = j.xseries
					yseries = j.yseries
				elif isinstance(j.xseries[0],list) == False and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.yseries)
					xseries = []
					i = 1
					while i <= nbOfSeries:
						xseries.append(j.xseries)
						i = i + 1
					yseries = j.yseries
				elif isinstance(j.yseries[0],list) == False and isinstance(j.xseries[0],list):
					nbOfSeries = len(j.xseries)
					yseries = []
					i = 1
					while i <= nbOfSeries:
						yseries.append(j.yseries)
						i = i + 1
					xseries = j.xseries
				else:
					xseries = [j.xseries]
					yseries = [j.yseries]
					nbOfSeries = 1
				if inputcolors is not None:
					colorBool = True
					for color in inputcolors:
						pyList.append(Color.FromArgb(color.Red,color.Green,color.Blue))
					wfColor = Array[Color](pyList)
				else:
					colorBool = False
				#Set a global Font code
				fontFam = FontFamily("Segoe UI Symbol")
				originalFont = Font(fontFam,8)
				#Chart created and Panel to host the chart
				chart1 = Chart()
				panelChart = Panel()
				panelChart.Name = "panelChart"
				#Panels colors
				chart1.BackColor = formbody.BackColor
				panelChart.BackColor = Color.Transparent
				#PanelChart location
				panelChart.Width = formwidth  - rightmargin - xlabel
				panelChart.Location = Point(xlabel,y + 30)				
				#Chart colors palette
				chart1.Palette =  ChartColorPalette.None
				if colorBool:
					chart1.PaletteCustomColors = wfColor
				else:
					chart1.Palette =  ChartColorPalette.None
				chart1.Series.Clear()
				#Series created
				def createSeries(seriesName,_keys,_values,_horizont):
					srs = Series()
					srs.Name = seriesName
					if _horizont:
						srs.ChartType = SeriesChartType.Bar
					else:
						srs.ChartType = SeriesChartType.Column
					srs.ToolTip = "Percent: #PERCENT"
					srs.IsValueShownAsLabel = True
					srs.Font = originalFont
					for i,j in zip(_keys,_values):
						srs.Points.AddXY(i,j)
					return srs
				#Create series
				series = []
				for i in range(0,nbOfSeries):
					series.append(createSeries(labels[i],xseries[i],yseries[i],horizont))
				#Add series to chart
				for s in series:
					chart1.Series.Add(s)
				#Refresh the Chart
				chart1.Invalidate()
				#Create a ChartArea and add it to the chart
				chartArea1 = ChartArea()
				chartArea1.BackColor = formbody.BackColor
				chartArea1.Name = "ChartArea1"
				chartArea1.AxisX.Title = j.xaxislabel
				chartArea1.AxisY.Title = j.yaxislabel
				chart1.ChartAreas.Add(chartArea1)
				#Create a Legend and add it to the chart
				legend1 = Legend()
				legend1.BackColor = formbody.BackColor
				legend1.Font = originalFont
				legend1.Name = "Legend1"
				legend1.Docking = Docking.Right
				chart1.Legends.Add(legend1) 
				#Initialize the chart and its properties
				chart1.BeginInit()
				chart1.Size = Size(panelChart.Width,panelChart.Width)
				chart1.AntiAliasing = AntiAliasingStyles.All
				fontFam = FontFamily("Segoe UI Symbol")
				titleFont = Font(fontFam,10,FontStyle.Bold)
				title = Title()
				title.Text = j.chartname
				title.Alignment = ContentAlignment.TopLeft
				title.Font = titleFont
				chart1.Titles.Add(title)
				#Add a button to push on view
				bt = Button()
				bt.Tag = "Push the chart on a view in Revit"
				bt.Text = "Push on a draft view"
				bt.Height = 20 * yRatio
				bt.Width = 140 * yRatio			
				panelChart.Controls.Add(bt)
				bt.BringToFront()
				bt.Click += form.imageexport
				#Add a checkBox to show/hide legend
				checkBox1 = CheckBox()
				checkBox1.AutoCheck = True
				checkBox1.Name = "checkBox1"
				checkBox1.Text = "Legend"
				panelChart.Controls.Add(checkBox1)
				checkBox1.BringToFront()
				checkBox1.Checked = True
				checkBox1.Font = originalFont
				checkBox1.AutoSize = True
				checkBox1.Click += form.chart_showLegend
				#Add a checkBox to show/hide labels
				checkBox2 = CheckBox()
				checkBox2.AutoCheck = True
				checkBox2.Name = "checkBox2"
				checkBox2.Text = "Labels"
				panelChart.Controls.Add(checkBox2)
				checkBox2.BringToFront()
				checkBox2.Checked = True
				checkBox2.Font = originalFont
				checkBox2.AutoSize = True
				checkBox2.Click += form.chart_showLabels
				
				panelChart.Size = Size(panelChart.Width, panelChart.Width + checkBox2.Height + 20)
				#Add chart to panel
				panelChart.Controls.Add(chart1)
				chart1.Location = Point(0,0)
				bt.Location = Point(panelChart.Width - bt.Width, panelChart.Width + 15)
				checkBox1.Location = Point(panelChart.Width - bt.Width - 80, panelChart.Width + 15)
				checkBox2.Location = Point(panelChart.Width - bt.Width - 160, panelChart.Width + 15)
				formbody.Controls.Add(panelChart)
				
				y = panelChart.Bottom + 25 * yRatio
				
			elif j.__class__.__name__ == 'uiradarchart':
				if isinstance(j.labels,list):
					labels = j.labels
				else:
					labels = [j.labels]
				inputcolors = j.inputcolors
				pyList = []
				if isinstance(j.xseries[0],list) and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.xseries)
					xseries = j.xseries
					yseries = j.yseries
				elif isinstance(j.xseries[0],list) == False and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.yseries)
					xseries = []
					i = 1
					while i <= nbOfSeries:
						xseries.append(j.xseries)
						i = i + 1
					yseries = j.yseries
				elif isinstance(j.yseries[0],list) == False and isinstance(j.xseries[0],list):
					nbOfSeries = len(j.xseries)
					yseries = []
					i = 1
					while i <= nbOfSeries:
						yseries.append(j.yseries)
						i = i + 1
					xseries = j.xseries
				else:
					xseries = [j.xseries]
					yseries = [j.yseries]
					nbOfSeries = 1
				if inputcolors is not None:
					colorBool = True
					for color in inputcolors:
						pyList.append(Color.FromArgb(color.Red,color.Green,color.Blue))
					wfColor = Array[Color](pyList)
				else:
					colorBool = False
				#Set a global Font code
				fontFam = FontFamily("Segoe UI Symbol")
				originalFont = Font(fontFam,8)
				#Chart created and Panel to host the chart
				chart1 = Chart()
				panelChart = Panel()
				panelChart.Name = "panelChart"
				#Create a panel for buttons
				panelCtrls = Panel()
				#Panels colors
				chart1.BackColor = formbody.BackColor
				panelChart.BackColor = Color.Transparent
				panelCtrls.BackColor = Color.Transparent
				#PanelChart location
				panelChart.Width = formwidth  - rightmargin - xlabel
				panelChart.Location = Point(xlabel,y + 30)				
				autoheight = 250 * yRatio			
				panelChart.Height = autoheight + 73 * yRatio
				ratio = (panelChart.Height)/(panelChart.Width)
				h = float(panelChart.Height)
				w = float(panelChart.Width)
				ratio = h/w
				chart_maxsize = formwidth - xlabel - rightmargin
				panelChart.Size = Size(chart_maxsize,chart_maxsize*ratio)				
				#Chart colors palette
				chart1.Palette =  ChartColorPalette.None
				if colorBool:
					chart1.PaletteCustomColors = wfColor
				else:
					chart1.Palette =  ChartColorPalette.None
				chart1.Series.Clear()
				#Series created
				def createSeries(seriesName,_keys,_values):
					srs = Series()
					srs.Name = seriesName
					srs.ChartType = SeriesChartType.Radar
					srs.ToolTip = "Percent: #PERCENT"
					srs.IsValueShownAsLabel = True
					srs.Font = originalFont
					srs.BorderWidth = 4
					for i,j in zip(_keys,_values):
						srs.Points.AddXY(i,j)
					return srs
				#Create series
				series = []
				for i in range(0,nbOfSeries):
					series.append(createSeries(labels[i],xseries[i],yseries[i]))
				#Add series to chart
				for s in series:
					chart1.Series.Add(s)
				#Refresh the Chart
				chart1.Invalidate()
				#Create a ChartArea and add it to the chart
				chartArea1 = ChartArea()
				chartArea1.BackColor = formbody.BackColor
				chartArea1.Name = "ChartArea1"
				chartArea1.AxisX.Title = j.xaxislabel
				chartArea1.AxisY.Title = j.yaxislabel
				chart1.ChartAreas.Add(chartArea1)
				#Create a Legend and add it to the chart
				legend1 = Legend()
				legend1.BackColor = formbody.BackColor
				legend1.Font = originalFont
				legend1.Name = "Legend1"
				chart1.Legends.Add(legend1) 
				#Initialize the chart and its properties
				chart1.BeginInit()
				chart1.Dock = DockStyle.Fill
				chart1.AntiAliasing = AntiAliasingStyles.All
				fontFam = FontFamily("Segoe UI Symbol")
				titleFont = Font(fontFam,10,FontStyle.Bold)
				title = Title()
				title.Text = j.chartname
				title.Alignment = ContentAlignment.TopLeft
				title.Font = titleFont
				chart1.Titles.Add(title)
				#Add a button to push on view
				bt = Button()
				bt.Tag = "Push the chart on a view in Revit"
				bt.Text = "Push on a draft view"
				bt.Height = 20 * yRatio
				bt.Width = 140 * yRatio				
				panelCtrls.Controls.Add(bt)
				bt.BringToFront()
				bt.Click += form.imageexport
				#Add a checkBox to show/hide legend
				checkBox1 = CheckBox()
				checkBox1.AutoCheck = True
				checkBox1.Name = "checkBox1"
				checkBox1.Text = "Show/Hide legend"
				panelCtrls.Controls.Add(checkBox1)
				checkBox1.BringToFront()
				checkBox1.Checked = True
				checkBox1.AutoSize = True
				checkBox1.Click += form.chart_showLegend
				
				#Add a checkBox to show/hide labels
				checkBox2 = CheckBox()
				checkBox2.AutoCheck = True
				checkBox2.Name = "checkBox2"
				checkBox2.Text = "Show/Hide labels"
				panelCtrls.Controls.Add(checkBox2)
				checkBox2.BringToFront()
				checkBox2.Checked = True
				checkBox2.Font = originalFont
				checkBox2.AutoSize = True
				checkBox2.Click += form.chart_showLabels
				
				panelChart.Size = Size(panelChart.Width, panelChart.Width + checkBox2.Height + 20)
				#Add chart to panel
				panelChart.Controls.Add(chart1)
				chart1.Location = Point(0,0)
				bt.Location = Point(panelChart.Width - bt.Width, panelChart.Width + 15)
				checkBox1.Location = Point(panelChart.Width - bt.Width - 80, panelChart.Width + 15)
				checkBox2.Location = Point(panelChart.Width - bt.Width - 160, panelChart.Width + 15)
				formbody.Controls.Add(panelChart)
				
				y = panelChart.Bottom + 25 * yRatio
				
			elif j.__class__.__name__ == 'uipointchart':
				if isinstance(j.labels,list):
					labels = j.labels
				else:
					labels = [j.labels]
				inputcolors = j.inputcolors
				pyList = []
				if isinstance(j.xseries[0],list) and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.xseries)
					xseries = j.xseries
					yseries = j.yseries
				elif isinstance(j.xseries[0],list) == False and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.yseries)
					xseries = []
					i = 1
					while i <= nbOfSeries:
						xseries.append(j.xseries)
						i = i + 1
					yseries = j.yseries
				elif isinstance(j.yseries[0],list) == False and isinstance(j.xseries[0],list):
					nbOfSeries = len(j.xseries)
					yseries = []
					i = 1
					while i <= nbOfSeries:
						yseries.append(j.yseries)
						i = i + 1
					xseries = j.xseries
				else:
					xseries = [j.xseries]
					yseries = [j.yseries]
					nbOfSeries = 1
				if inputcolors is not None:
					colorBool = True
					for color in inputcolors:
						pyList.append(Color.FromArgb(color.Red,color.Green,color.Blue))
					wfColor = Array[Color](pyList)
				else:
					colorBool = False
				#Set a global Font code
				fontFam = FontFamily("Segoe UI Symbol")
				originalFont = Font(fontFam,8)
				#Chart created and Panel to host the chart
				chart1 = Chart()
				panelChart = Panel()
				panelChart.Name = "panelChart"
				#Panels colors
				chart1.BackColor = formbody.BackColor
				panelChart.BackColor = Color.Transparent
				#PanelChart location
				panelChart.Width = formwidth  - rightmargin - xlabel
				panelChart.Location = Point(xlabel,y + 30)				
				#Chart colors palette
				chart1.Palette =  ChartColorPalette.None
				if colorBool:
					chart1.PaletteCustomColors = wfColor
				else:
					chart1.Palette =  ChartColorPalette.None
				chart1.Series.Clear()
				#Series created
				def createSeries(seriesName,_keys,_values):
					srs = Series()
					srs.Name = seriesName
					srs.ChartType = SeriesChartType.Point
					srs.ToolTip = "Percent: #PERCENT"
					srs.IsValueShownAsLabel = True
					srs.Font = originalFont
					for i,j in zip(_keys,_values):
						srs.Points.AddXY(i,j)
					return srs
				#Create series
				series = []
				for i in range(0,nbOfSeries):
					series.append(createSeries(labels[i],xseries[i],yseries[i]))
				#Add series to chart
				for s in series:
					chart1.Series.Add(s)
				#Change marker size
				originalMarkerSize = 8
				for s in series:
					for p in s.Points:
						p.MarkerSize = originalMarkerSize
				#Refresh the Chart
				chart1.Invalidate()
				#Create a ChartArea and add it to the chart
				chartArea1 = ChartArea()
				chartArea1.BackColor = formbody.BackColor
				chartArea1.Name = "ChartArea1"
				chartArea1.AxisX.Title = j.xaxislabel
				chartArea1.AxisY.Title = j.yaxislabel
				chart1.ChartAreas.Add(chartArea1)
				chartArea1 = chart1.ChartAreas[0]
				#Create a Legend and add it to the chart
				legend1 = Legend()
				legend2 = Legend()
				#Add legends to charts
				chart1.Legends.Add(legend1)
				chart1.Legends.Add(legend2)
				#Create Custom legend
				customevent = CustomizeLegendEventArgs(legend2.CustomItems)
				for s in series:
					legend2.CustomItems.Add(LegendItem(s.Name.ToString(),s.Color,""))					
				def chart_CustomizeLegend(sender, customevent ):
					chart1 = sender
					#legend1.Enabled = False
					legend2.Name = "Legend2"
					for s,it in zip(series,legend2.CustomItems):
						it.ImageStyle = LegendImageStyle.Marker
						it.MarkerStyle = s.MarkerStyle
						it.MarkerColor = s.Color
						it.BorderColor = Color.Transparent
						it.MarkerSize *= 1.3
				legend1.Enabled = False
				legend2.Font = originalFont
				legend2.IsTextAutoFit = True
				legend2.BackColor = formbody.BackColor			
				chart1.CustomizeLegend += chart_CustomizeLegend
				#Initialize the chart and its properties
				chart1.BeginInit()
				chart1.Size = Size(panelChart.Width,panelChart.Width)
				chart1.AntiAliasing = AntiAliasingStyles.All
				fontFam = FontFamily("Segoe UI Symbol")
				titleFont = Font(fontFam,10,FontStyle.Bold)
				title = Title()
				title.Text = j.chartname
				title.Alignment = ContentAlignment.TopLeft
				title.Font = titleFont
				chart1.Titles.Add(title)
				#Add a button to push on view
				bt = Button()
				bt.Tag = "Push the chart on a view in Revit"
				bt.Text = "Push on a draft view"
				bt.Height = 20 * yRatio
				bt.Width = 140 * yRatio				
				panelChart.Controls.Add(bt)
				bt.BringToFront()
				bt.Click += form.imageexport
				#Add a checkBox to show/hide legend
				checkBox1 = CheckBox()
				checkBox1.AutoCheck = True
				checkBox1.Name = "checkBox1"
				checkBox1.Text = "Legend"
				panelChart.Controls.Add(checkBox1)
				checkBox1.BringToFront()
				checkBox1.Checked = True
				checkBox1.AutoSize = True
				checkBox1.Click += form.chart_showLegend
				
				#Add a checkBox to show/hide labels
				checkBox2 = CheckBox()
				checkBox2.AutoCheck = True
				checkBox2.Name = "checkBox2"
				checkBox2.Text = "Labels"
				panelChart.Controls.Add(checkBox2)
				checkBox2.BringToFront()
				checkBox2.Checked = True
				checkBox2.Font = originalFont
				checkBox2.AutoSize = True
				checkBox2.Click += form.chart_showLabels
				
				panelChart.Size = Size(panelChart.Width, panelChart.Width + checkBox2.Height + 20)
				#Add chart to panel
				panelChart.Controls.Add(chart1)
				chart1.Location = Point(0,0)
				bt.Location = Point(panelChart.Width - bt.Width, panelChart.Width + 15)
				checkBox1.Location = Point(panelChart.Width - bt.Width - 80, panelChart.Width + 15)
				checkBox2.Location = Point(panelChart.Width - bt.Width - 160, panelChart.Width + 15)
				formbody.Controls.Add(panelChart)
				
				y = panelChart.Bottom + 25 * yRatio
								
			elif j.__class__.__name__ == 'uilinechart':
				if isinstance(j.labels,list):
					labels = j.labels
				else:
					labels = [j.labels]
				inputcolors = j.inputcolors
				pyList = []
				if isinstance(j.xseries[0],list) and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.xseries)
					xseries = j.xseries
					yseries = j.yseries
				elif isinstance(j.xseries[0],list) == False and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.yseries)
					xseries = []
					i = 1
					while i <= nbOfSeries:
						xseries.append(j.xseries)
						i = i + 1
					yseries = j.yseries
				elif isinstance(j.yseries[0],list) == False and isinstance(j.xseries[0],list):
					nbOfSeries = len(j.xseries)
					yseries = []
					i = 1
					while i <= nbOfSeries:
						yseries.append(j.yseries)
						i = i + 1
					xseries = j.xseries
				else:
					xseries = [j.xseries]
					yseries = [j.yseries]
					nbOfSeries = 1
				if inputcolors is not None:
					colorBool = True
					for color in inputcolors:
						pyList.append(Color.FromArgb(color.Red,color.Green,color.Blue))
					wfColor = Array[Color](pyList)
				else:
					colorBool = False
				#Set a global Font code
				fontFam = FontFamily("Segoe UI Symbol")
				originalFont = Font(fontFam,8)
				#Chart created and Panel to host the chart
				chart1 = Chart()
				panelChart = Panel()
				panelChart.Name = "panelChart"
				#Panels colors
				chart1.BackColor = formbody.BackColor
				panelChart.BackColor = Color.Transparent
				#PanelChart location
				panelChart.Width = formwidth  - rightmargin - xlabel
				panelChart.Location = Point(xlabel,y + 30)				
				#Chart colors palette
				chart1.Palette =  ChartColorPalette.None
				if colorBool:
					chart1.PaletteCustomColors = wfColor
				else:
					chart1.Palette =  ChartColorPalette.None
				chart1.Series.Clear()
				#Series created
				def createSeries(seriesName,_keys,_values):
					srs = Series()
					srs.Name = seriesName
					srs.ChartType = SeriesChartType.Line
					srs.ToolTip = "Percent: #PERCENT"
					srs.IsValueShownAsLabel = True
					srs.Font = originalFont
					srs.BorderWidth = 4
					for i,j in zip(_keys,_values):
						srs.Points.AddXY(i,j)
					return srs
				#Create series
				series = []
				for i in range(0,nbOfSeries):
					series.append(createSeries(labels[i],xseries[i],yseries[i]))
				#Add series to chart
				for s in series:
					chart1.Series.Add(s)
				#Refresh the Chart
				chart1.Invalidate()
				#Create a ChartArea and add it to the chart
				chartArea1 = ChartArea()
				chartArea1.BackColor = formbody.BackColor
				chartArea1.Name = "ChartArea1"
				chartArea1.AxisX.Title = j.xaxislabel
				chartArea1.AxisY.Title = j.yaxislabel
				chart1.ChartAreas.Add(chartArea1)
				#Create a Legend and add it to the chart
				legend1 = Legend()
				legend1.BackColor = formbody.BackColor
				legend1.Font = originalFont
				legend1.Name = "Legend1"
				chart1.Legends.Add(legend1) 
				#Initialize the chart and its properties
				chart1.BeginInit()
				chart1.Size = Size(panelChart.Width,panelChart.Width)
				chart1.AntiAliasing = AntiAliasingStyles.All
				fontFam = FontFamily("Segoe UI Symbol")
				titleFont = Font(fontFam,10,FontStyle.Bold)
				title = Title()
				title.Text = j.chartname
				title.Alignment = ContentAlignment.TopLeft
				title.Font = titleFont
				chart1.Titles.Add(title)
				#Add a button to push on view
				bt = Button()
				bt.Tag = "Push the chart on a view in Revit"
				bt.Text = "Push on a draft view"
				bt.Height = 20 * yRatio
				bt.Width = 140 * yRatio				
				panelChart.Controls.Add(bt)
				bt.BringToFront()
				bt.Click += form.imageexport
				#Add a checkBox to show/hide legend
				checkBox1 = CheckBox()
				checkBox1.AutoCheck = True
				checkBox1.Name = "checkBox1"
				checkBox1.Text = "Legend"
				panelChart.Controls.Add(checkBox1)
				checkBox1.BringToFront()
				checkBox1.Checked = True
				checkBox1.AutoSize = True
				checkBox1.Click += form.chart_showLegend
				
				#Add a checkBox to show/hide labels
				checkBox2 = CheckBox()
				checkBox2.AutoCheck = True
				checkBox2.Name = "checkBox2"
				checkBox2.Text = "Labels"
				panelChart.Controls.Add(checkBox2)
				checkBox2.BringToFront()
				checkBox2.Checked = True
				checkBox2.Font = originalFont
				checkBox2.AutoSize = True
				checkBox2.Click += form.chart_showLabels
				
				panelChart.Size = Size(panelChart.Width, panelChart.Width + checkBox2.Height + 20)
				#Add chart to panel
				panelChart.Controls.Add(chart1)
				chart1.Location = Point(0,0)
				bt.Location = Point(panelChart.Width - bt.Width, panelChart.Width + 15)
				checkBox1.Location = Point(panelChart.Width - bt.Width - 80, panelChart.Width + 15)
				checkBox2.Location = Point(panelChart.Width - bt.Width - 160, panelChart.Width + 15)
				formbody.Controls.Add(panelChart)
				
				y = panelChart.Bottom + 25 * yRatio
			
			elif j.__class__.__name__ == 'uisplinechart':
				if isinstance(j.labels,list):
					labels = j.labels
				else:
					labels = [j.labels]
				inputcolors = j.inputcolors
				pyList = []
				if isinstance(j.xseries[0],list) and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.xseries)
					xseries = j.xseries
					yseries = j.yseries
				elif isinstance(j.xseries[0],list) == False and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.yseries)
					xseries = []
					i = 1
					while i <= nbOfSeries:
						xseries.append(j.xseries)
						i = i + 1
					yseries = j.yseries
				elif isinstance(j.yseries[0],list) == False and isinstance(j.xseries[0],list):
					nbOfSeries = len(j.xseries)
					yseries = []
					i = 1
					while i <= nbOfSeries:
						yseries.append(j.yseries)
						i = i + 1
					xseries = j.xseries
				else:
					xseries = [j.xseries]
					yseries = [j.yseries]
					nbOfSeries = 1
				if inputcolors is not None:
					colorBool = True
					for color in inputcolors:
						pyList.append(Color.FromArgb(color.Red,color.Green,color.Blue))
					wfColor = Array[Color](pyList)
				else:
					colorBool = False
				#Set a global Font code
				fontFam = FontFamily("Segoe UI Symbol")
				originalFont = Font(fontFam,8)
				#Chart created and Panel to host the chart
				chart1 = Chart()
				panelChart = Panel()
				panelChart.Name = "panelChart"
				#Panels colors
				chart1.BackColor = formbody.BackColor
				panelChart.BackColor = Color.Transparent
				#PanelChart location
				panelChart.Width = formwidth  - rightmargin - xlabel
				panelChart.Location = Point(xlabel,y + 30)				
				#Chart colors palette
				chart1.Palette =  ChartColorPalette.None
				if colorBool:
					chart1.PaletteCustomColors = wfColor
				else:
					chart1.Palette =  ChartColorPalette.None
				chart1.Series.Clear()
				#Series created
				def createSeries(seriesName,_keys,_values):
					srs = Series()
					srs.Name = seriesName
					srs.ChartType = SeriesChartType.Spline
					srs.ToolTip = "Percent: #PERCENT"
					srs.IsValueShownAsLabel = True
					srs.Font = originalFont
					srs.BorderWidth = 4
					for i,j in zip(_keys,_values):
						srs.Points.AddXY(i,j)
					return srs
				#Create series
				series = []
				for i in range(0,nbOfSeries):
					series.append(createSeries(labels[i],xseries[i],yseries[i]))
				#Add series to chart
				for s in series:
					chart1.Series.Add(s)
				#Refresh the Chart
				chart1.Invalidate()
				#Create a ChartArea and add it to the chart
				chartArea1 = ChartArea()
				chartArea1.BackColor = formbody.BackColor
				chartArea1.Name = "ChartArea1"
				chartArea1.AxisX.Title = j.xaxislabel
				chartArea1.AxisY.Title = j.yaxislabel
				chart1.ChartAreas.Add(chartArea1)
				#Create a Legend and add it to the chart
				legend1 = Legend()
				legend1.BackColor = formbody.BackColor
				legend1.Font = originalFont
				legend1.Name = "Legend1"
				chart1.Legends.Add(legend1) 
				#Initialize the chart and its properties
				chart1.BeginInit()
				chart1.Size = Size(panelChart.Width,panelChart.Width)
				chart1.AntiAliasing = AntiAliasingStyles.All
				fontFam = FontFamily("Segoe UI Symbol")
				titleFont = Font(fontFam,10,FontStyle.Bold)
				title = Title()
				title.Text = j.chartname
				title.Alignment = ContentAlignment.TopLeft
				title.Font = titleFont
				chart1.Titles.Add(title)
				#Add a button to push on view
				bt = Button()
				bt.Tag = "Push the chart on a view in Revit"
				bt.Text = "Push on a draft view"
				bt.Height = 20 * yRatio
				bt.Width = 140 * yRatio			
				panelChart.Controls.Add(bt)
				bt.BringToFront()
				bt.Click += form.imageexport
				#Add a checkBox to show/hide legend
				checkBox1 = CheckBox()
				checkBox1.AutoCheck = True
				checkBox1.Name = "checkBox1"
				checkBox1.Text = "Legend"
				panelChart.Controls.Add(checkBox1)
				checkBox1.BringToFront()
				checkBox1.Checked = True
				checkBox1.AutoSize = True
				checkBox1.Click += form.chart_showLegend
				
				#Add a checkBox to show/hide labels
				checkBox2 = CheckBox()
				checkBox2.AutoCheck = True
				checkBox2.Name = "checkBox2"
				checkBox2.Text = "Labels"
				panelChart.Controls.Add(checkBox2)
				checkBox2.BringToFront()
				checkBox2.Checked = True
				checkBox2.Font = originalFont
				checkBox2.AutoSize = True
				checkBox2.Click += form.chart_showLabels
				
				panelChart.Size = Size(panelChart.Width, panelChart.Width + checkBox2.Height + 20)
				#Add chart to panel
				panelChart.Controls.Add(chart1)
				chart1.Location = Point(0,0)
				bt.Location = Point(panelChart.Width - bt.Width, panelChart.Width + 15)
				checkBox1.Location = Point(panelChart.Width - bt.Width - 80, panelChart.Width + 15)
				checkBox2.Location = Point(panelChart.Width - bt.Width - 160, panelChart.Width + 15)
				formbody.Controls.Add(panelChart)
				
				y = panelChart.Bottom + 25 * yRatio
				
			elif j.__class__.__name__ == 'uibubblechart':
				if isinstance(j.labels,list):
					labels = j.labels
				else:
					labels = [j.labels]
				inputcolors = j.inputcolors
				pyList = []
				if isinstance(j.xseries[0],list) and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.xseries)
					xseries = j.xseries
					yseries = j.yseries
				elif isinstance(j.xseries[0],list) == False and isinstance(j.yseries[0],list):
					nbOfSeries = len(j.yseries)
					xseries = []
					i = 1
					while i <= nbOfSeries:
						xseries.append(j.xseries)
						i = i + 1
					yseries = j.yseries
				elif isinstance(j.yseries[0],list) == False and isinstance(j.xseries[0],list):
					nbOfSeries = len(j.xseries)
					yseries = []
					i = 1
					while i <= nbOfSeries:
						yseries.append(j.yseries)
						i = i + 1
					xseries = j.xseries
				else:
					xseries = [j.xseries]
					yseries = [j.yseries]
					nbOfSeries = 1
				if inputcolors is not None:
					colorBool = True
					for color in inputcolors:
						pyList.append(Color.FromArgb(color.Red,color.Green,color.Blue))
					wfColor = Array[Color](pyList)
				else:
					colorBool = False
				#Set a global Font code
				fontFam = FontFamily("Segoe UI Symbol")
				originalFont = Font(fontFam,8)
				#Chart created and Panel to host the chart
				chart1 = Chart()
				panelChart = Panel()
				#panelChart.Padding = Padding(1) To create a black border
				chart1.BackColor = formbody.BackColor
				panelChart.BackColor = Color.Transparent
				#PanelChart location
				panelChart.Width = formwidth  - rightmargin - xlabel
				panelChart.Location = Point(xlabel,y + 30)				
				autoheight = 250 * yRatio			
				panelChart.Height = autoheight + 73 * yRatio
				ratio = (panelChart.Height)/(panelChart.Width)
				h = float(panelChart.Height)
				w = float(panelChart.Width)
				ratio = h/w
				chart_maxsize = formwidth - xlabel - rightmargin
				panelChart.Size = Size(chart_maxsize,chart_maxsize*ratio)				
				#Chart colors palette
				chart1.Palette =  ChartColorPalette.None
				if colorBool:
					chart1.PaletteCustomColors = wfColor
				else:
					chart1.Palette =  ChartColorPalette.None
				chart1.Series.Clear()
				#Series created
				def createSeries(seriesName,_keys,_values):
					srs = Series()
					srs.Name = seriesName
					srs.ChartType = SeriesChartType.Bubble
					srs.ToolTip = "Percent: #PERCENT"
					srs.Font = originalFont
					for i,j in zip(_keys,_values):
						srs.Points.AddXY(i,j)
					return srs
				#Create series
				series = []
				for i in range(0,nbOfSeries):
					series.append(createSeries(labels[i],xseries[i],yseries[i]))
				#Add series to chart
				for s in series:
					chart1.Series.Add(s)
				#Refresh the Chart
				chart1.Invalidate()
				#Create a ChartArea and add it to the chart
				chartArea1 = ChartArea()
				chartArea1.BackColor = formbody.BackColor
				chartArea1.Name = "ChartArea1"
				chartArea1.AxisX.Title = j.xaxislabel
				chartArea1.AxisY.Title = j.yaxislabel
				chart1.ChartAreas.Add(chartArea1)
				#Create a Legend and add it to the chart
				legend1 = Legend()
				legend1.BackColor = formbody.BackColor
				legend1.Font = originalFont
				legend1.Name = "Legend1"
				chart1.Legends.Add(legend1) 
				#Initialize the chart and its properties
				chart1.BeginInit()
				chart1.Dock = DockStyle.Fill
				chart1.AntiAliasing = AntiAliasingStyles.All
				fontFam = FontFamily("Segoe UI Symbol")
				titleFont = Font(fontFam,10,FontStyle.Bold)
				title = Title()
				title.Text = j.chartname
				title.Alignment = ContentAlignment.TopLeft
				title.Font = titleFont
				chart1.Titles.Add(title)
				#Add a button to push on view
				bt = Button()
				bt.Tag = "Push the chart on a view in Revit"
				bt.Text = "Push on a draft view"
				bt.Height = 20 * yRatio
				bt.Width = 140 * yRatio			
				panelChart.Controls.Add(bt)
				bt.BringToFront()
				bt.Click += form.imageexport
				#Add a checkBox to show/hide legend
				checkBox1 = CheckBox()
				checkBox1.AutoCheck = True
				checkBox1.Name = "checkBox1"
				checkBox1.Text = "Show/Hide legend"
				panelChart.Controls.Add(checkBox1)
				checkBox1.BringToFront()
				checkBox1.Checked = True
				checkBox1.Click += form.chart_showLegend
				#Add chart to panel
				panelChart.Size = Size(panelChart.Width, panelChart.Width + checkBox2.Height + 20)
				#Add chart to panel
				panelChart.Controls.Add(chart1)
				chart1.Location = Point(0,0)
				bt.Location = Point(panelChart.Width - bt.Width, panelChart.Width + 15)
				checkBox1.Location = Point(panelChart.Width - bt.Width - 80, panelChart.Width + 15)
				checkBox2.Location = Point(panelChart.Width - bt.Width - 160, panelChart.Width + 15)
				formbody.Controls.Add(panelChart)
				
				y = panelChart.Bottom + 25 * yRatio

			elif j.__class__.__name__ == 'uitreeview':
				tv = TreeView()
				tv.CheckBoxes = True
				titles = j._hastitles
				if j.inputname != "":
					tv.Width = formwidth - rightmargin - xinput
					tv.Location = Point(xinput,y)
				else:
					tv.Width = formwidth - rightmargin - xlabel
					tv.Location = Point(xlabel,y)
				tv.Height = j._height
				def treeIterationTitle(control, input):
					if isinstance(input,list):
						try:
							currentNode = TreeNode(remove_accents(input[0].ToString()))
							currentNode.Tag = ""
							input.pop(0)
						except:
							currentNode = TreeNode(control.Text +"."+ str(len(control.Nodes)+1))
							currentNode.Tag = ""
						control.Nodes.Add(currentNode)
						for i in input:
							treeIterationTitle(currentNode, i)
					else:
						item = TreeNode(remove_accents(input.ToString()))
						item.Tag = input
						control.Nodes.Add(item)	
				def treeIteration(control, input):
					if isinstance(input,list):
						currentNode = TreeNode(control.Text +"."+ str(len(control.Nodes)+1))
						currentNode.Tag = input
						control.Nodes.Add(currentNode)
						for i in input:
							treeIteration(currentNode, i)
					else:
						item = TreeNode(remove_accents(input.ToString()))
						item.Tag = input
						control.Nodes.Add(item)
				currentNode = TreeNode("List1")
				currentNode.Expand()				
				if titles:
					tv.Nodes.Add(currentNode)								
					for dl in j.datalist:
						treeIterationTitle(currentNode,dl)	
				else:
					tv.Nodes.Add(currentNode)								
					for dl in j.datalist:
						treeIteration(currentNode,dl)
				tv.AfterCheck += form.CheckChildren						
				formbody.Controls.Add(tv)
				form.output.append(tv)
				y = tv.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'listview':
				lvp = Panel()
				if j.inputname != "":
					lvp.Location = Point(xinput,y)
					lvp.Width = formwidth - rightmargin - xinput				
				elif j.inputname == "" and j.element_count != 0:
					lvp.Location = Point(xlabel,y+35*yRatio)
					lvp.Width = formwidth - rightmargin - xlabel
				else : 
					lvp.Location = Point(xlabel,y)
					lvp.Width = formwidth - rightmargin - xlabel				
				lvp.Height = j.height * yRatio + 25 * yRatio
				#Creating the listview box
				lv = mylistview()
				lv.Scrollable = True
				dummyheader = ColumnHeader()
				dummyheader.Text = ""
				dummyheader.Name = ""
				dummyheader.Width = -2 * xRatio
				lv.HeaderStyle = ColumnHeaderStyle.None
				lv.Columns.Add(dummyheader)
				lv.Values = j
				if not j.display_mode:
					lv.CheckBoxes = True
				lv.View = vi.Details
				lvItems =  [i for i in j.keys() if not (i == 'inputname' or i == 'height' or i == 'highlight' or i == 'display_mode' or i == 'element_count' or i == 'default_values' or i == 'sorted' or i == 'showId')]
				if j.sorted :
					lvItems = sorted(lvItems,key = remove_accents)
				else:
					pass
				[lv.Items.Add(i) for i in lvItems]
				lv.Location = Point(0,0)
				if j.inputname != "" and j.element_count != 0:
					lv.Width = formwidth - rightmargin - xinput
				elif j.inputname != "" and j.element_count == 0:
					lv.Width = formwidth - rightmargin - xinput
				else:
					lv.Width = formwidth - rightmargin - xlabel
				lv.Height = j.height * yRatio
				lv.Scrollable = True
				lv.ItemCheck += form.updateallnone
				for i in j.default_values:
					defInd = lvItems.index([x for x in j.keys() if not (x == 'inputname' or x == 'height' or x == 'highlight' or x == 'display_mode' or x == 'element_count' or x == 'default_values' or x == 'sorted' or x == 'showId')][i])
					lv.Items[defInd].Checked = True
				#Click listview items to hilight and center in appropriate view
				if j.highlight :
					lv.MouseDown += form.setviewforelement
					lv.MouseUp += form.zoomcenter
					items = list(lv.Items)
					lv.FullRowSelect = True
					del items[::2]
					for item in items:
						item.BackColor = Color.FromArgb(230,243,255)
				#Creating select all and select none radiobuttons
				if not j.display_mode:
					rball = RadioButton()
					rball.Location = Point(0,(j.height+5)*yRatio)
					rball.Width = 100 * xRatio
					rball.Height = 20 * yRatio
					rball.Font = SystemFonts.DefaultFont
					rball.Text = "Select all"
					rball.Click += form.selectall
					rbnone = RadioButton()
					rbnone.Location = Point(105 * xRatio,(j.height+5)*yRatio)
					rbnone.Width = 100 * xRatio
					rbnone.Height = 20 * yRatio
					rbnone.Font = SystemFonts.DefaultFont
					rbnone.Text = "Select none"			
					rbnone.Click += form.selectnone			
				#Adding controls to panel
				lvp.Controls.Add(lv)
				if not j.display_mode:
					lvp.Controls.Add(rball)
					lvp.Controls.Add(rbnone)
				formbody.Controls.Add(lvp)
				if not j.display_mode:
					form.output.append(lv)
				y = lvp.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uitext':
				tb = myTextBox()
				tb.Text = j.defaultvalue
				if j.inputname != "":
					tb.Width = formwidth  - rightmargin - xinput
					tb.Location = Point(xinput,y)
				else:
					tb.Width = formwidth - rightmargin - xlabel
					tb.Location = Point(xlabel,y)
				if j._isnum:
					tb._isNum = True
					tb.KeyPress += form.numsOnly
				formbody.Controls.Add(tb)
				formbody.Controls.Add(label)
				form.output.append(tb)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uimultilinetext':
				tb = myTextBox()
				tb.Text = j.defaultvalue
				tb.Multiline  = True
				tb.Height = j._height * yRatio
				if j.inputname != "":
					tb.Width = formwidth  - rightmargin - xinput
					tb.Location = Point(xinput,y)
				else:
					tb.Width = formwidth - rightmargin - xlabel
					tb.Location = Point(xlabel,y)
				if j._isnum:
					tb._isNum = True
					tb.KeyPress += form.numsOnly
				formbody.Controls.Add(tb)
				formbody.Controls.Add(label)
				form.output.append(tb)
				y = tb.Bottom + 25 * yRatio								
			elif j.__class__.__name__ == 'tableview':
				#Creating grouping panel
				tvp = Panel()
				tvp.Location = Point(xlabel,y)
				tvp.Width = formwidth - rightmargin - xlabel
				if (50 + len(j.dataList) * 15) * yRatio > 800 * yRatio:
					autoheight = 800 * yRatio
				else:
					autoheight = (50 + len(j.dataList) * 15	) * yRatio			
				tvp.Height = autoheight + 73 * yRatio
				#Creating title
				titlep = Label()
				titlep.Text = j._tabletitle
				titlep.Width = formwidth - rightmargin - xlabel
				titlep.BackColor = Color.FromArgb(153,180,209)
				titlep.Font = Font("Arial", 11, FontStyle.Bold)
				titlep.TextAlign = ContentAlignment.MiddleLeft
				titlep.BorderStyle = BorderStyle.FixedSingle
				titlep.Location = Point(0,0)
				tvp.Controls.Add(titlep)
				#Creating data structure
				dg = myDataGridView()
				#dg.SelectionMode = DataGridViewSelectionMode.CellSelect
				dg.EnableHeadersVisualStyles = False				
				dt = DataTable()
				dl = j.dataList
				for i in range(len(dl[0])):
					if j._hasTitle:
						dt.Columns.Add(remove_accents(dl[0][i].ToString()))
						rngstart = 1
					else:
						dt.Columns.Add("Column " + str(i))
						rngstart = 0						
				for rindex in range(rngstart,len(dl)):
					row = dt.NewRow()
					for i in range(len(dl[rindex])):
						row[i] = dl[rindex][i]
					dt.Rows.Add(row)						
				dg.Tag = j._hasTitle				
				dg.DataSource = dt
				dg.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText
				dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
				dg.Width = formwidth - rightmargin - xlabel
				dg.Location = Point(0,23 * yRatio)
				dg.Height = autoheight
				#Creatin Excel like drag copy functionalities
				dg.CellMouseDown += form.startCell
				dg.CellMouseUp += form.endCell
				#dg.MouseDown += form.startRowDrag
				#dg.MouseUp = 			
				tvp.Controls.Add(dg)			
				y = tvp.Bottom + 15 * yRatio
				#Adding export to excel button
				ex = Button()
				ex.Tag = [j._hasTitle , j._openExcel , j._showinfo , j._fileInfo]
				try:
					expImage = getImageByName("exp.png")[0]
					ex.BackgroundImage = expImage
				except:
					ex.Text = "Export"					
				ex.Width = 60 * xRatio
				ex.Height = 30 * yRatio
				ex.Location = Point(formwidth - 125 * xRatio ,dg.Bottom + 15 * yRatio)					
				tvp.Controls.Add(ex)
				ex.Click += form.exportToExcel
				#Adding filepath export textbox
				filepathtb = TextBox()
				filepathtb.Text = "ExportFileName"
				filepathtb.Location = Point(0 ,dg.Bottom + 20 * yRatio)
				filepathtb.Width = formwidth - xlabel - 125 * xRatio
				tvp.Controls.Add(filepathtb)
				#Adding copy to clipboard button
				cb = Button()
				#Adding panel to form
				formbody.Controls.Add(tvp)
				form.output.append(dg)
				y = tvp.Bottom + 25 * yRatio
				form.topmost()
			elif j.__class__.__name__ == 'uitextnote':
				gb = GroupBox()
				gb.Text = j.title
				gb.Parent = form
				gb.SendToBack()
				gb.BackColor = Color.Transparent
				gb.Location = Point(xlabel, y)
				tn = Label()
				tn.Location = Point(xlabel,18 * yRatio)
				tn.AutoSize = True
				tn.MaximumSize = Size(formwidth - rightmargin - 50 * xRatio,0)
				tn.Text = j.textnote
				tn.BringToFront()
				gb.Controls.Add(tn)
				gb.Size = Size(formwidth  - rightmargin - 25 * xRatio, tn.Bottom-tn.Top+25 * yRatio)  
				y = gb.Bottom + 25 * xRatio
				formbody.Controls.Add(gb)
			elif j.__class__.__name__ == 'uibool':
				yn = CheckBox()
				yn.Width = formwidth - xinput - rightmargin + 10 * xRatio
				yn.Location = Point(xinput,y)
				yn.Text = j.booltext
				g = yn.CreateGraphics()
				sizetext = g.MeasureString(yn.Text,yn.Font, formwidth - xinput - rightmargin -20 * xRatio)
				heighttext = sizetext.Height
				yn.Height = heighttext + 15 * yRatio
				yn.Checked = j.defaultvalue
				formbody.Controls.Add(yn)
				form.output.append(yn)
				y = yn.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiradio':
				yrb = 20 * yRatio
				rbs = []
				rbgroup = mygroupbox()
				if j.inputname != "":
					rbgroup.Width = formwidth - xinput - rightmargin
					rbgroup.Location = Point(xinput,y)
				else:
					rbgroup.Width = formwidth - xlabel - rightmargin
					rbgroup.Location = Point(xlabel,y)					
				rbgroup.Tag = j
				rbcounter = 0
				for key in j.keys():
					if key != 'inputname' and key != 'defaultvalue':
						rb = RadioButton()
						rb.Text = key 
						if j.inputname != "":
							rb.Width = formwidth - xinput - rightmargin - 35 * xRatio
						else:
							rb.Width = formwidth - xlabel - rightmargin - 35 * xRatio
						rb.Location = Point(20 * xRatio,yrb)
						if rbcounter == j.defaultvalue:
							rb.Checked = True
						rbgroup.Controls.Add(rb)
						g = rb.CreateGraphics()
						sizetext = g.MeasureString(key,rb.Font, formwidth - xinput - 90*xRatio)
						heighttext = sizetext.Height
						rb.Height = heighttext + 15 * yRatio
						ybot = rb.Bottom
						yrb += heighttext + 12 * yRatio
						rbcounter += 1
					else:
						pass
				rbgroup.Height = ybot + 20 * yRatio
				[rbgroup.Controls.Add(rb) for rb in rbs]
				formbody.Controls.Add(rbgroup)
				form.output.append(rbgroup)
				y = rbgroup.Bottom + 25 * yRatio
			elif j.__class__.__name__  == 'uifilepath':
				fp = Button()
				if j.inputname != "":
					fp.Width = formwidth - xinput - rightmargin
					fp.Location = Point(xinput,y)
				else:
					fp.Width = formwidth - xlabel - rightmargin		
					fp.Location = Point(xlabel,y)					
				fp.Tag = j.defaultvalue
				if not j.defaultvalue == "FilePath":
					fp.Tag = j.defaultvalue
					fp.MouseHover += form.showtooltip	
				fp.Text = j.buttontext
				fp.Height = 20 * yRatio				
				formbody.Controls.Add(fp)
				fp.Click += form.openfile
				form.output.append(fp)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uidirectorypath':
				dp = Button()
				if j.inputname != "":
					dp.Width = formwidth - xinput - rightmargin
					dp.Location = Point(xinput,y)
				else:
					dp.Width = formwidth - xlabel - rightmargin		
					dp.Location = Point(xlabel,y)				
				dp.Tag = j.defaultvalue
				dp.Text = j.buttontext
				if not j.defaultvalue == "DirectoryPath":
					dp.Tag = j.defaultvalue
					dp.MouseHover += form.showtooltip
				dp.Text = j.buttontext	
				dp.Height = 20 * yRatio				
				formbody.Controls.Add(dp)
				dp.Click += form.opendirectory
				form.output.append(dp)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiselectelements':
				se = Button()
				if j.inputname != "":
					se.Width = formwidth - xinput - rightmargin
					se.Location = Point(xinput,y)
				else:
					se.Width = formwidth - xlabel - rightmargin		
					se.Location = Point(xlabel,y)
				se.Text = j.buttontext
				se.Height = 20 * yRatio
				formbody.Controls.Add(se)
				if j.multi == False:
					se.Click += form.pickobjects
				else:
					se.Click  += form.pickobject
				form.output.append(se)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiselectautocadelements':
				se = Button()
				if j.inputname != "":
					se.Width = formwidth - xinput - rightmargin
					se.Location = Point(xinput,y)
				else:
					se.Width = formwidth - xlabel - rightmargin		
					se.Location = Point(xlabel,y)
				se.Text = j.buttontext
				se.Height = 20 * yRatio
				formbody.Controls.Add(se)
				if j.multi == False:
					se.Click += form.pickautocadobjects
				else:
					se.Click += form.pickautocadobject								
				form.output.append(se)
				y = label.Bottom + 25 * yRatio				
			elif j.__class__.__name__ == 'uiselectOrderedelements':
				se = Button()
				if j.inputname != "":
					se.Width = formwidth - xinput - rightmargin
					se.Location = Point(xinput,y)
				else:
					se.Width = formwidth - xlabel - rightmargin		
					se.Location = Point(xlabel,y)
				se.Text = j.buttontext
				se.Height = 20 * yRatio				
				formbody.Controls.Add(se)
				se.Click  += form.pickobjectsordered
				form.output.append(se)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiselectlinkedelements':
				se = Button()
				if j.inputname != "":
					se.Width = formwidth - xinput - rightmargin
					se.Location = Point(xinput,y)
				else:
					se.Width = formwidth - xlabel - rightmargin		
					se.Location = Point(xlabel,y)
				se.Text = j.buttontext
				se.Height = 20 * yRatio				
				formbody.Controls.Add(se)
				if j.multi == False:
					se.Click += form.picklinkedobjects
				else:
					se.Click  += form.picklinkedobject
				form.output.append(se)

				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiselectlinkedelementsofcat':
				sec = Button()
				if j.inputname != "":
					sec.Width = formwidth - xinput - rightmargin
					sec.Location = Point(xinput,y)
				else:
					sec.Width = formwidth - xlabel - rightmargin		
					sec.Location = Point(xlabel,y)
				sec.Text = j.buttontext
				sec.Tag = j.category
				sec.Height = 20 * yRatio
				formbody.Controls.Add(sec)
				if j.multi == False:
					sec.Click += form.picklinkedobjectsofcat
				else:
					sec.Click += form.picklinkedobjectofcat
				form.output.append(sec)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiselectelementsofcat':
				sec = Button()
				if j.inputname != "":
					sec.Width = formwidth - xinput - rightmargin
					sec.Location = Point(xinput,y)
				else:
					sec.Width = formwidth - xlabel - rightmargin		
					sec.Location = Point(xlabel,y)
				sec.Text = j.buttontext
				sec.Tag = j.category
				sec.Height = 20 * yRatio
				formbody.Controls.Add(sec)
				if j.multi == False:
					sec.Click += form.pickobjectsofcat
				else:
					sec.Click += form.pickobjectofcat
				form.output.append(sec)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiselectOrderedelementsofcat':
				sec = Button()
				if j.inputname != "":
					sec.Width = formwidth - xinput - rightmargin
					sec.Location = Point(xinput,y)
				else:
					sec.Width = formwidth - xlabel - rightmargin		
					sec.Location = Point(xlabel,y)
				sec.Text = j.buttontext
				sec.Tag = j.category
				sec.Height = 20 * yRatio
				formbody.Controls.Add(sec)
				sec.Click += form.pickobjectsofcatordered
				form.output.append(sec)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiselectfaces':
				sf = Button()
				if j.inputname != "":
					sf.Width = formwidth - xinput - rightmargin
					sf.Location = Point(xinput,y)
				else:
					sf.Width = formwidth - xlabel - rightmargin		
					sf.Location = Point(xlabel,y)
				sf.Text = j.buttontext
				sf.Height = 20 * yRatio
				formbody.Controls.Add(sf)
				sf.Click += form.pickfaces
				form.output.append(sf)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiselectpointsonface':
				spf = Button()
				if j.inputname != "":
					spf.Width = formwidth - xinput - rightmargin
					spf.Location = Point(xinput,y)
				else:
					spf.Width = formwidth - xlabel - rightmargin		
					spf.Location = Point(xlabel,y)
				spf.Text = j.buttontext
				spf.Height = 20 * yRatio
				formbody.Controls.Add(spf)
				spf.Click += form.pickpointsonface
				form.output.append(spf)
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiselectedges':
				sed = Button()
				if j.inputname != "":
					sed.Width = formwidth - xinput - rightmargin
					sed.Location = Point(xinput,y)
				else:
					sed.Width = formwidth - xlabel - rightmargin		
					sed.Location = Point(xlabel,y)
				sed.Text = j.buttontext
				sed.Height = 20 * yRatio
				formbody.Controls.Add(sed)
				sed.Click += form.pickedges
				form.output.append(sed)
				y = label.Bottom + 25 * yRatio
				
			elif j.__class__.__name__ == 'uislider':
				gb = Panel()
				if j.defaultvalue == '':
					defval = j.minimum
				else:
					defval = j.defaultvalue
				sl = mytrackbar(j.minimum,j.step)
				if j.inputname != "":
					gb.Width = formwidth - xinput - rightmargin
					gb.Location = Point(xinput,y)
				else:
					gb.Width = formwidth - xlabel - rightmargin		
					gb.Location = Point(xlabel,y)
				gb.Height = 40 * yRatio
				sl.Minimum = 0
				sl.Maximum = (j.maximum-j.minimum)/j.step
				sl.Value = (defval - j.minimum) / j.step
				sl.TickFrequency = (j.maximum-j.minimum)/j.step/10
				sl.Location = Point(40 * xRatio,0)
				if j.inputname != "":
					sl.Width = formwidth - xinput - rightmargin - 35 * xRatio
				else:
					sl.Width = formwidth - xlabel -rightmargin - 35 * xRatio
				gb.Controls.Add(sl)
				form.output.append(sl)
				displabel = Label()
				displabel.Width = 50 * xRatio
				displabel.Location = Point(0,5 * yRatio)
				displabel.Text = str(defval)
				displabel.Height = 30 * yRatio
				displabel.BringToFront()
				gb.Controls.Add(displabel)	
				formbody.Controls.Add(gb)			
				sl.Scroll += form.scroll
				y = label.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiimage':
				pic = Image.FromFile(j.image)
				im = PictureBox()
				if j.showborder:
					im.BorderStyle = BorderStyle.Fixed3D
				else:
					im.BorderStyle = BorderStyle.None
				ratio = (pic.Height)/(pic.Width)
				h = float(pic.Height)
				w = float(pic.Width)
				ratio = h/w
				image_maxsize = formwidth - xlabel - rightmargin
				im.Size = Size(image_maxsize,image_maxsize*ratio)
				scaledimage = Bitmap(image_maxsize,image_maxsize*ratio)
				gr = Graphics.FromImage(scaledimage)
				gr.DrawImage(pic,0,0,image_maxsize,image_maxsize*ratio)
				im.Image = pic
				im.SizeMode = PictureBoxSizeMode.Zoom
				formbody.Controls.Add(im)
				im.Location = Point(25 * xRatio,y+25*yRatio)
				y = im.Bottom + 25 * yRatio
		
			elif j.__class__.__name__ == 'uicolorpick' and importcolorselection == 0:
				cp = Button()
				if j.inputname != "":
					cp.Width = formwidth - xinput - rightmargin
					cp.Location = Point(xinput,y)
				else:
					cp.Width = formwidth - xlabel - rightmargin		
					cp.Location = Point(xlabel,y)
				cp.Text = j.buttontext
				cp.Height = 30 * yRatio
				formbody.Controls.Add(cp)
				cp.Click += form.colorpicker
				form.output.append(cp)
				y = label.Bottom + 25 * yRatio	
			elif j.__class__.__name__ == 'uicolorpick' and importcolorselection == 1:
				importcolorselection = 2
			elif j.__class__.__name__ == 'uigroup':
				grouppanel = GroupBox()
				grouppanel.Text = j.groupname
				grouppanel.Location = Point(xlabel,y)
				grouppanel.Width = formwidth - 65 * xRatio
				#recursive implementation of the definition to process grouped inputs
				addinput(grouppanel,j.inputgroup,25*yRatio,xinput-25*xRatio,80 * xRatio,90 * xRatio,importcolorselection)
				formbody.Controls.Add(grouppanel)
				y = grouppanel.Bottom + 25 * yRatio
			elif j.__class__.__name__ == 'uiconditional':
				grouppanelcond = GroupBox()
				grouppanelcond.Location = Point(xlabel,y)
				grouppanelcond.Width = formwidth - 65 * xRatio	
				form.output.append(grouppanelcond)				
				#recursive implementation of the definition to process conditional groups of inputs
				panlist = []
				rblist = []
				yp = 25 * yRatio
				xrb = 25 * xRatio
				for i,d in zip(j._Conditions,j._GroupedInputs):	
					rb = RadioButton()
					rb.CheckedChanged += form.ActivateOption
					rb.Text = i
					rb.Location	= Point(xrb,10*yRatio)
					rb.Width = 95 * xRatio
					rblist.append(rb)					
					grouppanelcond.Controls.Add(rb)				
					condition_pannel = Panel()
					condition_pannel.Name = i
					condition_pannel.Top = yp
					condition_pannel.Width = formwidth - 65 * xRatio
					condition_pannel.BackColor = Color.Transparent
					condition_pannel.BringToFront()					
					addinput(condition_pannel,d,25*yRatio,xinput-25*xRatio,80*xRatio,90*xRatio,importcolorselection)
					panlist.append(condition_pannel)
					yp = condition_pannel.Bottom - 25*yRatio
					xrb += 100 * xRatio
				for pan in panlist:
					grouppanelcond.Controls.Add(pan)
				for e,pan in enumerate(panlist):
					if e != j._defaultOptionIndex:
						pan.Enabled = False
					else:
						continue
				rblist[j._defaultOptionIndex].Checked = True
				grouppanelcond.Height = sum([p.Height for p in panlist]) - (len(panlist)-2)*25 * yRatio
				grouppanelcond.BackColor = Color.Transparent				
				formbody.Controls.Add(grouppanelcond)
				form.output.append([grouppanelcond.Tag])				
				y = grouppanelcond.Bottom + 25*yRatio
			formbody.Height = y


	# process input lists 
	addinput(body,inputtypes,0,IN[9],40 * xRatio ,IN[9] * xRatio,importcolorselection)
	
	#add the formbody panel to the form
	form.Controls.Add(body)		
	

	if IN[6] != None:
		if IN[8] > 400 * yRatio:
			formy += 50 * yRatio
			xinput = 270 * yRatio
		else:
			formy = logo.Bottom + 20 * yRatio
	else:
		formy += 50 * yRatio



	#Adding validation button
	
	button = Button()
	button.Text = IN[1]
	button.Width = formwidth - xinput - 40 * xRatio
	button.Height = 20 * yRatio
	button.Location = Point (xinput,formy)
	button.Click += form.setclose
	form.Controls.Add(button)
	form.MaximizeBox = False
	form.MinimizeBox = False
	form.FormBorderStyle = FormBorderStyle.FixedSingle
	
	#Adding Cancel button
	if IN[6] != None:
		cancelbutton = Button()
		cancelbutton.Text = IN[6]
		cancelbutton.Width = 120 * xRatio
		cancelbutton.Height = 20 * xRatio
		cancelbutton.Name = "Cancel"
		cancelbutton.Location = Point (xinput -120 * xRatio ,formy)
		cancelbutton.Click += form.setclose
		form.Controls.Add(cancelbutton)
		form.CancelButton = cancelbutton
	
	#Adding link to help
	
	if IN[5] != None :
		helplink = LinkLabel()
		helplink.Text = "Help"
		helplink.Tag = IN[5]
		helplink.Click += form.openurl
		helplink.Location = Point(formwidth - 65 * xRatio ,formy+30 * yRatio)
		form.Controls.Add(helplink)
	else:
		pass	
			
	form.ShowIcon = True
	form.Width = formwidth
	form.Height = formy + 120 * yRatio
	formfooterheight = form.Height - formheaderheight
	
	# Make formbody scrollable
	
	# if the form is longer than its maximum height, do the following:
	# modify the form height
	# make the formbody panel scrollable
	if form.Height + body.Height > IN[7] * yRatio and IN[7] * yRatio > 0:
		formbody_calculatedheight = IN[7] * yRatio - form.Height
		# make sure the formbody is at least 100 px high
		if formbody_calculatedheight < 100 * yRatio: formbody_calculatedheight = 100 * yRatio
		body.Height = formbody_calculatedheight
		form.Height = form.Height + formbody_calculatedheight
		# this (and apparently only this) will implement a vertical AutoScroll *only*
		# http://stackoverflow.com/a/28583501
		body.HorizontalScroll.Maximum = 0
		body.AutoScroll = False
		body.VerticalScroll.Visible = False
		body.AutoScroll = True
		body.BorderStyle = BorderStyle.Fixed3D
	else:
		form.Height = form.Height + body.Height
	# Move footer elements
	logo.Location = Point(logo.Location.X, logo.Location.Y + body.Height)
	button.Location = Point(button.Location.X, button.Location.Y + body.Height)
	if IN[6] != None: cancelbutton.Location = Point(cancelbutton.Location.X, cancelbutton.Location.Y + body.Height)
	if IN[5] != None: helplink.Location = Point(helplink.Location.X, helplink.Location.Y + body.Height)

	#Positionning the form at top left of current view
	#In The revit environment
	try:
		uiviews = uidoc.GetOpenUIViews()
		if doc.ActiveView.IsValidType(doc.ActiveView.GetTypeId()):
			activeviewid = doc.ActiveView.Id
			activeUIView = [v for v in uiviews if v.ViewId == activeviewid][0]
		else:
			activeUIView = uiviews[0]
		rect = activeUIView.GetWindowRectangle()
		form.StartPosition = FormStartPosition.Manual
		form.Location = Point(rect.Left-7 * xRatio,rect.Top)
	except:
		pass
	
		
	if IN[2]:
		if importcolorselection != 2:
			Application.Run(form)
			result = form.values
			OUT = result,True, form.cancelled 
		else:
			OUT = ['ColorSelection input is only available With Revit 2017'] , False, False
	else :
		OUT = ['Set toggle to true!'] , False, False
except:
	import traceback
	OUT = traceback.format_exc() , "error", "error"


