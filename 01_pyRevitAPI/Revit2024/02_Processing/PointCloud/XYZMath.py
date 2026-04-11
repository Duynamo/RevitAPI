import math
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
# Placeholder class for Autodesk.Revit.DB.XYZ equivalent in Python
class XYZ:
    def __init__(self, x, y, z):
        self.X = x
        self.Y = y
        self.Z = z

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
