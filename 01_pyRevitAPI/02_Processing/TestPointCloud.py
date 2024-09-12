import open3d as o3d
import numpy as np
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
# Function to scale the point cloud
def scale_point_cloud(point_cloud, scale_factor):
    points = np.asarray(point_cloud.points)
    # Scaling the points by the factor
    scaled_points = points * scale_factor
    point_cloud.points = o3d.utility.Vector3dVector(scaled_points)
    return point_cloud
def pickObject():
    from Autodesk.Revit.UI.Selection import ObjectType
    ref = uidoc.Selection.PickObject(ObjectType.Element)
    return  doc.GetElement(ref.ElementId)
# Read the point cloud from file (replace 'point_cloud.ply' with your file)


pointCloud = pickObject()
pcd = o3d.io.read_point_cloud("pointCloud.ply")  # Replace with your point cloud file

# Scale factor (1/100)
scale_factor = 1 / 100

# Scale the point cloud
scaled_pcd = scale_point_cloud(pcd, scale_factor)
pointCloud = pickObject()

OUT = 1