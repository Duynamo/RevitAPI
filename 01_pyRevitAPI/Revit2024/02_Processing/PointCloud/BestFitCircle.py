from numpy import *
from matplotlib.pyplot import *
"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

from math import cos,sin,tan,radians
from scipy.optimize import least_squares
clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitAPI") 
import Autodesk
from Autodesk.Revit.DB import* 
from Autodesk.Revit.DB.Structure import*

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

# clr.AddReference("System.Windows.Forms")
# clr.AddReference("System.Drawing")
# clr.AddReference("System.Windows.Forms.DataVisualization")

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *
#endregion
#region ___Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView
#endregion

#region ___someFunctions
def uwList(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)
def flatten(nestedList):
    flatList = []
    for item in nestedList:
        if isinstance(item, list):
            flatList.extend(flatten(item))
        else:
            flatList.append(item)
    return flatList
#endregion
#source https://meshlogic.github.io/posts/jupyter/curve-fitting/fitting-a-circle-to-cluster-of-3d-points/
def fit_circle_3d(points):
    """
    Fit a circle to 3D points with normalization for numerical stability.
    Args:
        points: (N, 3) array of 3D points.
    Returns:
        center: (3,) array, the center of the circle.
        radius: The radius of the circle.
        normal: (3,) array, the normal vector of the circle's plane.
        zenith: Zenith angle (in radians) of the normal vector.
        azimuth: Azimuth angle (in radians) of the normal vector.
    """
    points = np.array(points)
    if points.ndim != 2 or points.shape[1] != 3:
        raise ValueError("Input points must be a 2D array with shape (N, 3).")
    # Normalize for numerical stability
    mean_point = np.mean(points, axis=0)
    scale_factor = np.mean(np.linalg.norm(points - mean_point, axis=1))
    points_normalized = (points - mean_point) / scale_factor
    # Initial guesses
    centroid = np.mean(points_normalized, axis=0)
    normal_guess = np.array([0, 0, 1])
    radius_guess = np.mean(np.linalg.norm(points_normalized - centroid, axis=1))
    initial_guess = np.hstack([centroid, normal_guess, radius_guess])
    # Residuals function
    def residuals(params):
        cx, cy, cz, nx, ny, nz, r = params
        n = np.array([nx, ny, nz])
        n = n / np.linalg.norm(n)  # Normalize the normal vector
        residuals_plane = np.dot(points_normalized - [cx, cy, cz], n)  # Distance to plane
        projected_points = points_normalized - np.outer(residuals_plane, n)
        distances = np.linalg.norm(projected_points - [cx, cy, cz], axis=1) - r
        return np.hstack([residuals_plane, distances])
    # Fit using least squares
    result = result = least_squares(residuals, initial_guess, method='trf', ftol=1e-12, xtol=1e-12)
    cx, cy, cz, nx, ny, nz, r = result.x
    # Rescale to original dimensions
    center = np.array([cx, cy, cz]) * scale_factor + mean_point
    normal = np.array([nx, ny, nz])
    radius = r * scale_factor
    # Compute zenith and azimuth
    zenith = np.arccos(normal[2] / np.linalg.norm(normal))
    azimuth = np.arctan2(normal[1], normal[0])
    return center, radius, normal, zenith, azimuth
temp=[]
def Create_Slide_Section():
    t=Transaction(doc,"highlight")
    t.Start()
    error=""
    distance=1/304.8
    try:
#region pick section
        point_box=view.GetSectionBox()
        pick=pointsss=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.PointOnElement)
        picked_point=pick.GlobalPoint
        Min_Box= point_box.Min
        Max_Box= point_box.Max
        #trans identity
        trans=Transform.Identity
        trans.BasisX=point_box.Transform.BasisX
        trans.BasisY=point_box.Transform.BasisY
        trans.BasisZ=point_box.Transform.BasisZ
        trans.Origin=XYZ(picked_point.X,picked_point.Y,Max_Box.Z)
        #box identity
        box=BoundingBoxXYZ()
        box.MaxEnabled=True
        box.MinEnabled=True 
        box_height  =5000/304.8
        box_width   =1/304.8
        box_depth   =5000/304.8
        box.Min=XYZ(-box_height/(2),-box_width/(2),-box_depth/(2))
        box.Max=XYZ(box_height/(2),box_width/(2),box_depth/(2))
        #higher than expected
        #lower the box
        box.Transform=trans
        view.SetSectionBox(box)
        view.IsSectionBoxActive=True   
        doc.ReloadLatest()
        t.Commit()
#endregion
        return 0
    except Exception as e:
        error=e
        t.Commit()
    return error
def find_distance(first, second):
    x = first.X - second.X
    y = first.Y - second.Y
    z = first.Z - second.Z
    return math.sqrt(x * x + y * y + z * z)
def fit_ellipse(x,y):
    x=np.mean(x)
    y=np.mean(y)
    D=np.vstack([x**2,x*y,y**2,x,y,np.ones_like(x)]).T
    D[:,-1]=1
    def residuals(params):
        return D @ params - 1
    initial_guess=np.ones(6) 
    result=least_squares(residuals,initial_guess)
    A,B,C,D,E,F=result.x            
    return A,B,C,D,E,F
def point_Interaction():
    pointCloudInstance=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element)
    picked_point=pointCloudInstance.GlobalPoint
    pointCloudInstance=doc.GetElement(pointCloudInstance)
    distance=10e-4  
    numberOfPoints=10e4
    boundingBox=view.get_BoundingBox(None)
    transform=boundingBox.Transform
#region filter 1
    Min=XYZ(-504,235,18)
    Max=XYZ(-499,238,23)
    # return Min,Max
    planes=[]
    #X boundaries
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisX,Min))
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisX.Negate(),Max))
    #Y boundaries
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisY,Min))
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisY.Negate(),Max))
    #Z boundaries
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisZ,Min))
    planes.append(Plane.CreateByNormalAndOrigin(XYZ.BasisZ.Negate(),Max))
    filter=PointClouds.PointCloudFilterFactory.CreateMultiPlaneFilter(planes)
    pointsCollection=pointCloudInstance.GetPoints(filter,distance,numberOfPoints) 
    points=[]
    #circle properties
    radius=0.0
    centre=XYZ()
    centre_X=0.0
    centre_Y=0.0
    centre_Z=0.0
    inside=0
    percentage=0.0
    trans=Transform.Identity
    temp_point=XYZ()
    for point in pointsCollection:
        temp_point=XYZ(point.X,point.Y,point.Z)
        break
    trans.Origin=XYZ(temp_point.X-picked_point.X,temp_point.Y-picked_point.Y,temp_point.Z-picked_point.Z)
    trans.BasisX=XYZ(1,0,0)
    trans.BasisY=XYZ(0,1,0)
    trans.BasisZ=XYZ(0,0,1)
    for point in pointsCollection:
        point=XYZ(point.X,point.Y,point.Z)
        point=trans.Inverse.OfPoint(point)
        points.append([point.X,point.Y,point.Z])
    picked_point2=XYZ(0,0,0)
    pointsCollection2=pointsCollection
#2nd array
    pointCloudInstance2=uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element)
    picked_point2=pointCloudInstance2.GlobalPoint
    pointCloudInstance2=doc.GetElement(pointCloudInstance2)
    pointsCollection2=pointCloudInstance2.GetPoints(filter,distance,numberOfPoints) 
    for point in pointsCollection2:
        temp_point=XYZ(point.X,point.Y,point.Z)
        break
    trans.Origin=XYZ(temp_point.X-picked_point2.X,temp_point.Y-picked_point2.Y,temp_point.Z-picked_point2.Z)
    trans.BasisX=XYZ(1,0,0)
    trans.BasisY=XYZ(0,1,0)
    trans.BasisZ=XYZ(0,0,1)
    for point in pointsCollection2:
        point=XYZ(point.X,point.Y,point.Z)
        point=trans.Inverse.OfPoint(point)
        points.append([point.X,point.Y,point.Z])
    centre=fit_circle_3d(points)[0]
    centre=XYZ(centre[0],centre[1],centre[2])
    radius=fit_circle_3d(points)[1]
    return picked_point,picked_point2,centre,radius,points
def create_curve(point1,point2,point3):
    error=""
    t=Transaction(doc,"create curve")
    t.Start()
    try:
        geoPlane=Plane.CreateByThreePoints(point1,point2,point3)
        sketchPlane=SketchPlane.Create(doc,geoPlane)
        arc=Autodesk.Revit.DB.Arc.Create(point1,point2,point3)
        curve=doc.Create.NewModelCurve(arc,sketchPlane)
    except Exception as e:
        error=e
    finally:
        t.Commit()
    return error
def create_circle(centre,radius,*args):
    picked_point=centre
    t=Transaction(doc," ")
    error=""
    radius=radius
    t.Start()
    try:
        transform=view.get_BoundingBox(None).Transform
        geoEllipse=Ellipse.CreateCurve(picked_point,radius,radius,transform.BasisX,transform.BasisY,-2*radius,2*radius)
        geoPlane=Plane.CreateByNormalAndOrigin(transform.BasisZ,picked_point)
        temp.append(transform.BasisZ)
        sketchPlane=SketchPlane.Create(doc,geoPlane) 
        arc=doc.Create.NewModelCurve(geoEllipse,sketchPlane)
        length=len(args)
        for i in range(0,length-1,100):
            args[i]=XYZ(args[i][0],args[i][1],args[i][2])
            args[i+1]=XYZ(args[i+1][0],args[i+1][1],args[i+1][2])
            geoLine = Line.CreateBound(centre, args[i])
            geoPlane=Plane.CreateByThreePoints(args[i],args[i+1],centre)
            sketchPlane=SketchPlane.Create(doc,geoPlane)
            line=doc.Create.NewModelCurve(geoLine, sketchPlane)
    except Exception as e:
        error=e
    finally:
        t.Commit()
    return error,picked_point,radius*304.8
def create_model_line(point1,point2,point3):
    error=""
    t=Transaction(doc,"Create Model Line")
    t.Start()
    try:
        geoLine = Line.CreateBound(point1, point2)
        geoPlane=Plane.CreateByThreePoints(point1,point2,point3)
        sketchPlane=SketchPlane.Create(doc,geoPlane)
        line=doc.Create.NewModelCurve(geoLine, sketchPlane)
    except Exception as e:
        error=e
    finally:
        t.Commit()
    return error
def create_modelLine_list(circle_list):
    error=[]
    pointList=[]
    temp_point=XYZ(0,0,0)
    length=len(circle_list)
    i=0
    for _ in range(0,length):
        if i>=length:
            break
        if length-i==2:
            t=create_model_line(circle_list[i],circle_list[length-1],XYZ(0,0,0))
            break
        vector_1=circle_list[i].Subtract(circle_list[i+1])
        vector_2=circle_list[i+2].Subtract(circle_list[i+1])
        angle=vector_1.AngleTo(vector_2)
        if angle<2.9:
            c=create_curve(circle_list[i],circle_list[i+2],circle_list[i+1])
            i=i+2
            error.append(c)
        else:
            temp_point=XYZ(0,0,0)
            t=create_model_line(circle_list[i],circle_list[i+1],temp_point)
            i+=1
            error.append(t)
    return error,length
pipes=[]
types=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsElementType().ToElements()
piping_system_types=Autodesk.Revit.DB.FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipingSystem).ToElements()
level_ids=FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
def create_pipe(p1,p2):
    error=[]
    t=Transaction(doc,"create pipes")
    t.Start()
    try:
        pipe_created=Autodesk.Revit.DB.Plumbing.Pipe.Create(doc,
													piping_system_types[0].Id,
													types[0].Id,
													level_ids[1].Id,
													p1,
													p2)
        pipes.append(pipe_created)
    except Exception as e:
        error.append(e)
    finally:
        t.Commit()
    return pipes
circle_list=[]
def create_centre_line():
    error=""
    try:
        for i in range(0,6):
            t=point_Interaction()
            circle_list.append(t[2])
    except Exception as e:
        error=e
    return error,len(circle_list),circle_list
def create_pipes(circle_list):
    error=[]
    for i in range(0,len(circle_list)-1):
        t=create_pipe(circle_list[i],circle_list[i+1])
        error.append(t)
    return error
connectors=[]
def get_connectors():
    for pipe in pipes:
        for connector in pipe.ConnectorManager.Connectors:
            connectors.append(connector)
    return connectors
def connect():
    error=[]
    for i in range(1,len(connectors)-1,2):
        t=Transaction(doc,"connect")
        t.Start()
        try:
            DocumentManager.Instance.CurrentDBDocument.Create.NewElbowFitting(connectors[i],connectors[i+1])
        except Exception as e:
            error=e
        finally:
            t.Commit()
    return error
# for i in range(0,7):
#     temp_point=uidoc.Selection.PickPoint()
#     circle_list.append(temp_point)
# def create_PipenConnector():
#     a=create_centre_line()
#     b=create_pipes(circle_list)
#     c=get_connectors()
#     temp.append(a)
#     temp.append(b)
#     temp.append(c)
# create_PipenConnector()
# OUT=connect()
get=point_Interaction()
# OUT=create_circle(get[2],get[3],get[4]),get[0],get[1]
OUT=get
