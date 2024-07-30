import clr
import sys 
import System   
import math
from System.Collections.Generic import *
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

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms.DataVisualization")
clr.AddReference("DSCoreNodes")
from DSCore.List import Flatten

import System.Windows.Forms 
from System.Windows.Forms import *
import System.Drawing
from System.Drawing import *

"""_____________________________"""
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = uiapp.ActiveUIDocument
view = doc.ActiveView
"""_________"""

faces = []
dsfaces = []

faces.append(uidoc.Selection.PickObjects(ObjectType.Face,'Select Faces'))
for f in faces:
	for r in f:
		e = uidoc.Document.GetElement(r)
		dsface = e.GetGeometryObjectFromReference(r).ToProtoType(True)
		[f.Tags.AddTag("RevitFaceReference", r) for f in dsface]
		dsfaces.append(dsface)


#region _join
import clr

# Import DesignScript Grometry Library
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

# Import ToDSType(bool) extension method
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

# Import DocumentManager and TransactionManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application

from System.Collections.Generic import *

# Import RevitAPI
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

# 이 노드에 대한 입력값은 IN 변수에 리스트로 저장됩니다.
dataEnteringNode = IN

# 코드를 이 선 아래에 배치
Elements = UnwrapElement(IN[0])

joinlog = []

for element in Elements:
    
    for element2 in Elements:
        
        if(element.Id == element2.Id): continue
        
        if (JoinGeometryUtils.AreElementsJoined(doc,element,element2)) == False:
            
            try:
                TransactionManager.Instance.EnsureInTransaction(doc)

                JoinGeometryUtils.JoinGeometry(doc,element,element2)

                TransactionManager.Instance.TransactionTaskDone()
                
                joinlog.append("성공")

            except:

                joinlog.append("실패")

OUT = joinlog
[출처] Dynamo Python - Join|작성자 소주장군


#endregion

#region _un join
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument

# Import ToDSType(bool) extension method
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

#입력
elements = UnwrapElement(IN[0])

#출력 할것
log = []

#트렌젝션 시작
TransactionManager.Instance.EnsureInTransaction(doc)

for element in elements:
	try:
		StructuralFramingUtils.DisallowJoinAtEnd(element,0)
		StructuralFramingUtils.DisallowJoinAtEnd(element,1)
		log.append("성공")
	except:
		log.append("실패")
# End Transaction
TransactionManager.Instance.TransactionTaskDone()

OUT = log
[출처] DisallowJoin (결합 금지 설정)|작성자 소주장군

#endregion