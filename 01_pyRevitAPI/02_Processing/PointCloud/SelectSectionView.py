"""Copyright by: vudinhduybm@gmail.com"""
#region ___import all Library
import clr
import sys 
import System   
import math
import collections

from math import cos,sin,tan,radians

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
class XYZMath:
    # Define a precision of double data
    PRECISION = 1e-10
    @staticmethod
    def find_mid_point(first, second):
        """
        Find the middle point of the line.

        :param first: The start point of the line (XYZ)
        :param second: The end point of the line (XYZ)
        :return: The middle point of the line (XYZ)
        """
        x = (first.X + second.X) / 2
        y = (first.Y + second.Y) / 2
        z = (first.Z + second.Z) / 2
        return XYZ(x, y, z)

    @staticmethod
    def find_distance(first, second):
        """
        Find the distance between two points.

        :param first: The first point (XYZ)
        :param second: The second point (XYZ)
        :return: The distance between the two points (float)
        """
        x = first.X - second.X
        y = first.Y - second.Y
        z = first.Z - second.Z
        return math.sqrt(x * x + y * y + z * z)
    @staticmethod
    def find_direction(first, second):
        """
        Find the normalized direction vector from 'first' to 'second'.

        :param first: The start point (XYZ)
        :param second: The end point (XYZ)
        :return: The direction vector (XYZ)
        """
        x = second.X - first.X
        y = second.Y - first.Y
        z = second.Z - first.Z
        distance = XYZMath.find_distance(first, second)
        return XYZ(x / distance, y / distance, z / distance)
    @staticmethod
    def find_right_direction(view_direction):
        """
        Find the direction rotated 90 degrees around the Z-axis from the given view direction.

        :param view_direction: The input direction vector (XYZ)
        :return: The right direction vector (XYZ)
        """
        # Rotate 90 degrees around the Z axis
        x = -view_direction.Y
        y = view_direction.X
        z = view_direction.Z  # Z remains unchanged
        
        return XYZ(x, y, z)
    @staticmethod
    def find_up_direction(view_direction):
        """
        Find the up direction, which is always aligned with the Z-axis in this example.

        :param view_direction: The input direction vector (XYZ)
        :return: The up direction vector (XYZ)
        """
        # Up direction is always aligned with the Z-axis
        return XYZ(0, 0, 1)
    @staticmethod
    def find_floor_view_direction(curve_array):
        """
        Find the view direction for the floor based on the curve array.
        Assumes the floor is always level, and the first curve provides the direction.

        :param curve_array: A list or array-like object of Curve objects
        :return: The view direction vector (XYZ)
        """
        # Get the first curve in the array
        curve = curve_array[0]
        
        # Get the start and end points of the curve
        first = curve.get_end_point(0)
        second = curve.get_end_point(1)
        
        # Calculate and return the direction vector
        return XYZMath.find_direction(first, second)
def distance(xyz1, xyz2):
    d = 0.0
    d = math.sqrt(  (xyz2[0] - xyz1[0])**2 + 
                    (xyz2[1] - xyz1[1])**2 + 
                    (xyz2[2] - xyz1[2])**2)
    return d
def createSection2():


    t=Transaction(doc,"create section view")
    t.Start()
    trans=Transform.Identity
    trans.Origin=XYZ(0,0,0)
    trans.BasisX=XYZ(0,0,1)
    trans.BasisY=XYZ(1,0,0)
    trans.BasisZ=XYZ(0,1,0)
    section_box=BoundingBoxXYZ()
    section_box.Min=XYZ(-1,-1,-1)
    section_box.Max=XYZ(1,1,1)
    section_box.Transform=trans
    section_type_id=doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
    new_section=ViewSection.CreateSection(doc,section_type_id,section_box)
    t.Commit()

def tranform_point(point,transform):
    x = point.X
    y = point.Y
    z = point.Z

    # Transform basis vectors and origin of the new coordinate system
    b0 = transform.BasisX
    b1 = transform.BasisY
    b2 = transform.BasisZ
    origin = transform.Origin

    # Transform the origin of the old coordinate system in the new coordinate system
    x_temp = x * b0.X + y * b1.X + z * b2.X + origin.X
    y_temp = x * b0.Y + y * b1.Y + z * b2.Y + origin.Y
    z_temp = x * b0.Z + y * b1.Z + z * b2.Z + origin.Z

    return XYZ(x_temp, y_temp, z_temp)

def createBoundingBox(XYZ1,XYZ2) :
    box=BoundingBoxXYZ()
    #find origin
    trans=Transform.Identity
    #get root vector
    vector=XYZ1.Subtract(XYZ2)
    vector=vector.Normalize().Negate()
    #create transform
    #set origin to O(0,0,0)
    midpoint=XYZMath.find_mid_point(XYZ1,XYZ2)
    #define length
    trans.BasisY=vector
    #define depth
    trans.BasisX=XYZ(0,0,1)
    #define height
    trans.BasisZ=trans.BasisX.CrossProduct(trans.BasisY).Normalize()
    box_depth=2000/304.8
    box_height=5000/304.8
    box_length=XYZ1.DistanceTo(XYZ2)
    trans.Origin=midpoint
    minPoint=Autodesk.Revit.DB.XYZ(-box_height/2,-box_length/2,0)
    maxPoint=Autodesk.Revit.DB.XYZ(box_height/2,box_length/2,box_depth)
    temp.append(minPoint)
    temp.append(maxPoint)
    #-----Box * Trans=Bounding box
    box.Max=maxPoint
    box.Min=minPoint
    box.MaxEnabled=True
    box.MinEnabled=True
    box.Transform=trans
    return box,trans.BasisX,trans.BasisY,trans.BasisZ
#create bouding box perpendicular by 2 points
temp=[]
def createBoundingBox_2(XYZ1,XYZ2) :
    box=BoundingBoxXYZ()
    #find origin
    trans=Transform.Identity
    #get root vector
    vector=XYZ1.Subtract(XYZ2)
    vector=vector.Normalize()
    midpoint=XYZMath.find_mid_point(XYZ1,XYZ2)
    #create transform
    trans.BasisX=XYZ(0,0,1)
    trans.BasisZ=vector
    trans.BasisY=vector.CrossProduct(trans.BasisX).Normalize()
    box_depth=3000/304.8
    box_height=5000/304.8
    box_length=3000/304.8
    trans.Origin=midpoint
    minPoint=Autodesk.Revit.DB.XYZ(-box_height/2,-box_length/2,0)
    maxPoint=Autodesk.Revit.DB.XYZ(box_height/2,box_length/2,box_depth)
    temp.append(minPoint)
    temp.append(maxPoint)
    # -----Box * Trans=Bounding box
    box.Max=maxPoint
    box.Min=minPoint
    box.MaxEnabled=True
    box.MinEnabled=True
    box.Transform=trans
    return box,trans.BasisX,trans.BasisY,trans.BasisZ

def getFamilyId():
    elements=Autodesk.Revit.DB.FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
    for e in elements:
        v = e if isinstance(e,ViewFamilyType) else None
        if v is not None:
            return v.Id
    return 0
#perpendicular to axis
def createSection2(XYZ1,XYZ2):
    error=""
    string=""
    t=Transaction(doc,"CreateSectionView")
    t.Start()
    try:                                         
        familyId=doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
        box=createBoundingBox_2(XYZ1,XYZ2)[0]
        ViewSection.CreateSection(doc,familyId,box)
        t.Commit()
        return box                                                              
    except Exception as e:
        t.Commit()
        error=e
    return error
#2point
def createSection1(XYZ1,XYZ2):
    error=""
    string=""
    t=Transaction(doc,"CreateSectionView")
    t.Start()
    try:                                         
        familyId=doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeSection)
        box=createBoundingBox(XYZ1,XYZ2)[0]
        ViewSection.CreateSection(doc,familyId,box)
        t.Commit()
        return box                                                              
    except Exception as e:
        t.Commit()
        error=e
    return error
def selection():
    first_Point=uidoc.Selection.PickPoint()
    second_Point=uidoc.Selection.PickPoint()
    createSection1(first_Point,second_Point)
    return first_Point,second_Point
points_dictionary=[]
def create_point_dict(point,vector,length,x):
    index=int(length/x)
    vector=vector.Multiply(x)
    XYZ1=point
    for i in range(1,index):
        XYZ2=XYZ1.Add(vector)
        points_dictionary.append(XYZ1)
        points_dictionary.append(XYZ2)
        XYZ1=XYZ2
def get_points():
    error_lists=[]
    list_pipes = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element , "s")
    list_length=[6,7,8,9]
    divided_points=[]
    points=[]
    lengths=[]
    vectors=[]
    for i in range (0,len(list_pipes)):
        list_pipes[i]=doc.GetElement(list_pipes[i])
        points.append(list_pipes[i].Location.Curve.Origin)
        vectors.append(list_pipes[i].Location.Curve.Direction)
        lengths.append(list_pipes[i].Location.Curve.Length)
        create_point_dict(points[i],vectors[i],lengths[i],list_length[i])
    return points,lengths,vectors,list_length
def createSection3(dictionary):
    for i in range(0,len(dictionary)-1):
        createSection2(dictionary[i],dictionary[i+1])
        i+=1
    return 0
# get_points()
# OUT = createSection3(points_dictionary)
