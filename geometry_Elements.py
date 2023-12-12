import abc
from os import remove
import clr 
import System
import math 
from System.Collections.Generic import *

clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import*
from  Autodesk.Revit.UI.Selection import*

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
view = doc.ActiveView
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
"""-----------------------PART 3: GEOMETRIC ELEMENTS.-----------------------"""

def lstFlattenL2(_list):
    re = []
    for i in _list:
        for j in i:
            re.append(j)
    return re
def GetGeoElement(element): # Get geometry of elements.
    geo = []
    opt = Options() # Geometrical analysis
    opt.ComputeReferences = True
    opt.IncludeNonVisibleObjects = True
    opt.DetailLevel = ViewDetailLevel.Fine
    geoByElement = element.get_Geometry(opt)
    geo = [i for i in geoByElement]
    return geo
def GetSolidFromGeo(lstGeo): #Get Solid from geometry.
    sol = []
    for i in lstGeo:
        if i.GetType() == Solid and i.Volume > 0:
            sol.append(i)
        elif i.GetType() == GeometryInstance:
            var = i.SymbolGeometry
            for j in var:
                if j.Volume > 0:
                    sol.append(j)
    return sol
def GetPlanarFormSolid(solids): #Ger planarFace Not null value
    plaf = []
    for i in solids:
        var = i.Faces
        for j in var:
            if j.Reference != None:
                plaf.append(j)
    return plaf
def Isparael(p,q):
    return p.CrossProduct(q).IsZeroLength()
def FilterVerticalPlanar(lstPlface): # Get Vertical PlanarFace.
    faV = []
    x = XYZ.BasisX
    for i in lstPlface:
        faNomal = i.FaceNormal
        check = Isparael(x,faNomal)
        if check == True:
            faV.append(i)
    return faV
def FilterHorizontalPlanar(lstPlface): # Get Horizontal PlanarFaces
    faH = []
    z = XYZ.BasisZ
    for i in lstPlface:
        faNomal = i.FaceNormal
        check = Isparael(z,faNomal)
        if check == True:
            faH.append(i)
    return  faH
def RemoveFaceNone(lstPlanars): # Get planarFaces Not Null Value
    pfaces = []
    for i in lstPlanars:
        if i.Reference != None:
            pfaces.append(i)
    return pfaces
def GetFaceVertical(planar):  # Get Vertical PlannarFaces 
    re = []
    remove = RemoveFaceNone(planar)
    for i in remove:
        var = i.FaceNormal
        rad = var.AngleTo(XYZ.BasisZ)
        if 30<(rad*180/3.14)<170:
            re.append(i)
    return re
def GetMaxface(planars):  #Get GetMaxface 
    _Area =[]
    _Pface = []
    for i in planars:
        _Area.append(i.Area)
    for j in planars:
        if j.Area == max(_Area):
            _Pface.append(j)
    return _Pface
def GetMinface(planars): # Get GetMinface 
    _Area =[]
    _Pface = []
    for i in planars:
        _Area.append(i.Area)
    for j in planars:
        if j.Area == min(_Area):
            _Pface.append(j)
    return _Pface
def GetAreaface(planars): # Get Area Face 
    _Area =[]
    for i in planars:
        _Area.append(i.Area)
    return _Area
def GetNear_Minface(planars):
    _Area =[]
    _Pface = []
    for i in planars:
        _Area.append(i.Area)
    for j in planars:
        if j.Area < (min(_Area) + 1):
            _Pface.append(j)
    return _Pface
def GetNear_Maxface(planars): 
    _Area =[]
    _Pface = []
    for i in planars:
        _Area.append(i.Area)
    for j in planars:
        if j.Area >= (max(_Area) - 1):
            _Pface.append(j)
    return _Pface
def RightFace(planar,viewin): # Get Right PlannarFaces of a Element
    direc = view.RightDirection
    for i in planar:
        var = i.FaceNormal
        if var.IsAlmostEqualTo(direc):
            return i
def LeftFace(planar,viewin): # Get Right PlannarFaces of a Element
    direc = view.RightDirection
    for i in planar:
        var = -1*i.FaceNormal
        if var.IsAlmostEqualTo(direc):
            return i
def GetRightOrLeftFace(lstPlanar,reason,viewin):
    if reason == True: return RightFace(lstPlanar,viewin)
    elif reason == False: return LeftFace(lstPlanar,viewin)
def GetTopOrBotFace(lstPlanar, reason): # Get Top or Bottom of Faces
    r1 = []
    for i in lstPlanar:
        if i.FaceNormal.Z == 1 and reason == True:
            return i
    for i in lstPlanar:
        if i.FaceNormal.Z == -1 and reason == False:
            return i
def GetTopFaceEle(lstplanar):
    re = []
    remove =RemoveFaceNone(lstplanar)
    for i in remove:
        var = i.FaceNormal
        rad = var.AngleTo(XYZ.BasisZ)
        if (rad*180)/3.14 < 10:
            re.append (i)
    return re
def GetBotFaceEle(lstplanar):
    re = []
    remove =RemoveFaceNone(lstplanar)
    for i in remove:
        var = i.FaceNormal
        rad = var.AngleTo(XYZ.BasisZ)
        if (rad*180)/3.14 > 120:
            re.append (i)
    return re
def RightLstFace(lstplanar,viewin): # Get Right PlannarFaces of Elements
    direc = []
    for i in lstplanar:
        direc.append(view.RightDirection)
    faces = []
    for t,j in zip(direc,lstplanar):
        var = j.FaceNormal
        if var.IsAlmostEqualTo(t):
            faces.append(j)
    return faces
    #OUT = [RightLstFace(i, view) for i in getFaceFraming ]
def LeftLstFace(lstplanar,viewin): # Get Left PlannarFaces of Elements
    direc = []
    for i in lstplanar:
        direc.append(view.RightDirection)
    faces = []
    for t,j in zip(direc,lstplanar):
        var = -1*j.FaceNormal
        if var.IsAlmostEqualTo(t):
            faces.append(j)
    return faces
    #OUT = [LeftLstFace(i, view) for i in getFaceFraming ]
def Getlst_RightOrLeftFace(lstPlanar,reason,viewin):
    if reason == True: return RightLstFace(lstPlanar,viewin)
    elif reason == False: return LeftLstFace(lstPlanar,viewin)
def RetrieveEdgesFace(lstsPlanar): # Get Revit.DB.Line of Planarface
    re = []
    var = lstsPlanar.EdgeLoops
    for i in var:
        for j in i:
            re.append(j.AsCurve())
    return re
def GetRefFromEdgeLoop(lstPlanar):
    re = []
    var = lstPlanar.EdgeLoops
    for i in var:
        for j in i:
            re.append(j.Reference)
    return re
    #OUT = [GetRefFromEdgeLoop(j) for i in getMaxFaceFraming for j in i]
def GetRefFromPlanar(lstPlanar): # Get Reference of PlanarFace
    re = []
    for i in lstPlanar:
        re.append(i.Reference)
    return re
def GetRefArrayFromRef(ref): # Get ReferenceArray form Ref
    re = ReferenceArray()
    for i in ref:
        re.Append(i)
    return re
def GetLineMin(lstLine): #Get min lines of listline
    _length = []
    for i in lstLine:
        _length.append(i.Length)
    for j in lstLine:
        if j.Length == min(_length):
            return j
def GetLineMax(lstLine): #Get max lines of listline
    _length = []
    for i in lstLine:
        _length.append(i.Length)
    for j in lstLine:
        if j.Length == max(_length):
            return j
def LineOffset(line, distance, direc): #Offset Revit.DB.Line
    convert = distance/304.8
    newVector = None
    if direc == "X": newVector = XYZ(convert,0,0)
    if direc == "Y": newVector = XYZ(0,convert,0)
    if direc == "Z": newVector = XYZ(0,0,convert)

    trans = Transform.CreateTranslation(newVector)
    lineMove = line.CreateTransformed(trans)
    return lineMove
"""_______________________________"""

fls = FilteredElementCollector(doc, view.Id).ToElements()
AllEle = []
eleFloors = []
eleFraming =[]
for i in fls:
    try:
        if i.Category.Name == "Structural Framing":
            eleFraming.append(i)
        if i.Category.Name == "Floors":
            eleFloors.append(i)
        if i.Category.Name == "Structural Framing" or i.Category.Name == "Floors":
            AllEle.append(i)
    except:
        pass
#OUT = AllEle
getGeoFraming = [GetGeoElement(i) for i in eleFraming]
getGeoFloors = [GetGeoElement(i) for i in eleFloors]
getGeoAllEle = [GetGeoElement(i) for i in AllEle]
OUT = getGeoFraming, getGeoFloors , getGeoAllEle

getSolidFraming = [GetSolidFromGeo(i) for i in getGeoFraming]
getSolidFloors = [GetSolidFromGeo(i) for i in getGeoFloors]
getSolidAllEle = [GetSolidFromGeo(i) for i in getGeoAllEle]
OUT = getSolidFraming, getSolidFloors , getSolidAllEle

getFaceFraming = [GetPlanarFormSolid(i) for i in getSolidFraming]
getFaceFloors = [GetPlanarFormSolid(i) for i in getSolidFloors]
getFaceAllEle = [GetPlanarFormSolid(i) for i in getSolidAllEle]
OUT = getFaceFraming, getFaceFloors,getFaceAllEle

getFaHoriFraming = [FilterHorizontalPlanar(i) for i in getFaceFraming]
getFaHoriFloor = [FilterHorizontalPlanar(i) for i in getFaceFloors]
getFaHoriAllEle = [FilterHorizontalPlanar(i) for i in getFaceAllEle]
OUT = getFaHoriAllEle

getFaVerFraming = [FilterVerticalPlanar(i) for i in getFaceFraming]
getFaVerFloors = [FilterVerticalPlanar(i) for i in getFaceFloors]
getFaVerAllEle = [FilterVerticalPlanar(i) for i in getFaceAllEle]
OUT = getFaVerAllEle

rightFaceFraimg  =  [GetRightOrLeftFace(i,True,view) for i in getFaVerFraming]
leftFaceFraimg = [GetRightOrLeftFace(i,False,view) for i in getFaVerFraming]
getLR_Vframing = rightFaceFraimg, leftFaceFraimg

getLineFraming = [RetrieveEdgesFace(i) for i in leftFaceFraimg]
OUT = getLineFraming

linePlaceH = [GetLineMin(i) for i in getLineFraming]
OUT = linePlaceH

dimLineH = [LineOffset(i,IN[1],"X") for i in linePlaceH]
OUT = dimLineH

getB_Framing = [GetTopOrBotFace((i),False) for i in getFaceFraming]
OUT = getB_Framing
getBline_Framing = [RetrieveEdgesFace(i) for i in getB_Framing]
OUT = getBline_Framing

linePlaceV = [GetLineMin(i) for i in getBline_Framing]
OUT = linePlaceV

dimLineV = [LineOffset(i,IN[1],"Z") for i in linePlaceV]
OUT = dimLineV, dimLineH

OUT = getFaHoriAllEle
####
re = []
reduceZ = set()
for i in getFaHoriAllEle:
    for j in i:
        if j.Origin.Z not in reduceZ:
            reduceZ.add(j.Origin.Z)
            re.append(j)
OUT = re

getFaHoriAllEle = re



getFaceH_Allele = [GetRefFromPlanar(getFaHoriAllEle)]
OUT = getFaceH_Allele

###_______________________________________________________###
getFaceH_Allele = lstFlattenL2(getFaceH_Allele)
OUT = getFaceH_Allele

refArrayFraming = [GetRefArrayFromRef(getFaceH_Allele)]
OUT = refArrayFraming[0]
TransactionManager.Instance.EnsureInTransaction(doc)
dimV = doc.Create.NewDimension(view, dimLineH[0], refArrayFraming[0])
TransactionManager.Instance.TransactionTaskDone()
###_______________________________________________________###



###_______________________________________________________###
getT_Framing = [GetTopOrBotFace((i),True) for i in getFaceFraming]
getTB_Framing = getT_Framing, getB_Framing
getHRefFraming = lstFlattenL2([GetRefFromPlanar(i) for i in getTB_Framing])
#OUT = getHRefFraming
refArrayBTFraming = [GetRefArrayFromRef(getHRefFraming)][0]
OUT = refArrayBTFraming
dimLineH1 =  [LineOffset(i,IN[2],"X") for i in dimLineH]
OUT = dimLineH,dimLineH1
TransactionManager.Instance.EnsureInTransaction(doc)
dimV = doc.Create.NewDimension(view, dimLineH1[0], refArrayBTFraming)
TransactionManager.Instance.TransactionTaskDone()
###_______________________________________________________###



###_______________________________________________________###
getRef_VFraming = lstFlattenL2([GetRefFromPlanar(i) for i in getLR_Vframing])
OUT = getRef_VFraming
refArrayRLFraming = [GetRefArrayFromRef(getRef_VFraming)][0]
OUT = refArrayBTFraming
TransactionManager.Instance.EnsureInTransaction(doc)
dimV = doc.Create.NewDimension(view, dimLineV[0], refArrayRLFraming)
TransactionManager.Instance.TransactionTaskDone()
###_______________________________________________________###
def GetLineVertical(lstLine): # Get vertical Line by Isparalel
    re = []
    NorRevit = XYZ.BasisZ
    for i in lstLine:
        NorLo = i.Direction
        if Isparael(NorLo, NorRevit):
            re.append(i)
    return re

def GetIntersection(face, line):
    re = []
    results = clr.Reference[IntersectionResultArray]()
    intersect = face.Intersect(line, results)
    if intersect == SetComparisonResult.Overlap:
        var1 = results.Item[0]
        var2 = var1.XYZPoint
        re.append(var2)
    return re

def PointOffset(lstPoint, dis, face):
    re = []
    if len(lstPoint) > 1:
        ptsX = [i.X for i in lstPoint]
        maxX = max(ptsX)
        for i in lstPoint:
            if i.X == maxX:
                 re.append(XYZ(-1*dis/304.8+ i.X, i.Y, i.Z)) # Right
            else:
                re.append(XYZ(1*dis/304.8 + i.X, i.Y, i.Z)) # Left
    else:
        ptsCompare = face.Origin.X # for A floor display
        for i in lstPoint:
            if i.X > ptsCompare: re.append(XYZ(-1*dis/304.8 + i.X, i.Y, i.Z))
            else: re.append(XYZ(1*dis/304.8 + i.X, i.Y, i.Z))
    return re


###----------------------------------------CropView-----------------------------####
cropview = view.GetCropRegionShapeManager().GetCropShape()
OUT = cropview
####---------------------------------------GetLineVertical----------------------####
lineCrop = lstFlattenL2([GetLineVertical(i) for i in cropview])
OUT = lineCrop


####---------------------------------------GetTopFloorsFace----------------------####
get_TopPlanarFloor = lstFlattenL2([GetTopFaceEle(i) for i in getFaceFloors])
OUT = get_TopPlanarFloor
####---------------------------------------GetTopFloorsFace----------------------####


ptsInterect = []
for i,j in zip(get_TopPlanarFloor,lineCrop):
    ptsInterect.append(GetIntersection(i,j))
ptsInterect = lstFlattenL2(ptsInterect)
OUT = ptsInterect

pts_putFa = PointOffset(ptsInterect,50,view)

familyBreak = []

TransactionManager.Instance.EnsureInTransaction(doc)
for i in pts_putFa:
    putFa_Break  = doc.Create.NewFamilyInstance(i,UnwrapElement(IN[3]),view)
    familyBreak.append(putFa_Break)
TransactionManager.Instance.TransactionTaskDone()

OUT = familyBreak

re = []
for i in familyBreak:
    re.append(i.flipHand() )
OUT = re

OUT = [i.L for i in UnwrapElement(familyBreak)]