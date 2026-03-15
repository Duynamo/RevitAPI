#region ___import Python Package
from aqa.math import *
from varmain.primitiv import *
from varmain.custom import *

#endregion

#region __metaData

@activate (Group="Flange",
          TooltipShort="SO_Flange",
          TooltipLong="Slip On Flange",
          FirstPortEndtypes="FL",
          LengthUnit="mm",
          Ports="2")
@group ("MainDimensions")
@param (D=LENGTH, TooltipShort="Outer Diameter of the Flange")
@param (D1=LENGTH, TooltipShort="Pipe Diameter of the Flange")
@param (T=LENGTH, TooltipLong="Thickness of the Flange")
@param (OF=LENGTH, TooltipLong="Weld offset")
@group (Name="meaningless enum")
@param (K=ENUM)
@enum  (1, "align X")
@enum  (2, "align Y")
@enum  (3, "align Z")

#endregion

def SO_Flange(s, D=210.0, D1=114.3, T=10, OF =- 1, K=1, ** kw):
    F = CYLINDER(s, R=D/2, H=T, O=D1/2).rotateY(90)
    return None

flange = SO_Flange

OUT = flange