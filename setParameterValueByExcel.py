"""Copyright by: vudinhduybm@gmail.com"""
##########################################
import clr 

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

#Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

def famDoc_setParameter(famDoc, p, t, v):
    try:
        famMan = famDoc.FamilyManager
        famMan.CurrentType = t
        famMan.Set(p, v)
        return True
    except:
        return False

famDocs = IN[0]
famNames = [f.Title for f in famDocs]

famTypesList , famTypeNamesList , famParamsList , famParamNamesList = [],[],[],[]

for f in famDocs :
    famMan = f.FamilyManager
    famTypes = list(famMan.Types)
    typeNames = [f.Name for f in famTypes]
    famTypesList.append(famTypes)
    famTypeNamesList.append(typeNames)
    params = famMan.GetParameters()
    paramNames = [p.Definition.Name for p in params]
    famParamsList.append(params)
    famParamNamesList.append(paramNames)

outcomesList = []
paramNames   = IN[1]
excelData    = IN[2]

for row in excelData:
    outcomes = []
    famInd = famNames.index(row[0])
    famDoc = famDocs[famInd]
    typeInd = famTypeNamesList[famInd].index(row[1])
    famType = famTypesList[famInd][typeInd]
    values = row[2:]
    t = Transaction(famDoc, i[0] + " - " + i[1])
    t.Start()
    for p,v in zip(paramNames, values):
        parInd = famParamNamesList[famInd].index(p)
        famPar = famParamsList[famInd][parInd]
        setParam = famDoc_setParameter(famDoc, famPar, famType, v)
        outcomes.append(setParam)
        t.Commit()
    outcomesList.append(outcomes)

    outcomesList.append(row[0])

OUT = outcomesList
"""Copyright by: vudinhduybm@gmail.com"""
##########################################
import clr 

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

#Current doc/app/ui
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
view = doc.ActiveView

def famDoc_setParameter(famDoc, p, t, v):
    try:
        famMan = famDoc.FamilyManager
        famMan.CurrentType = t
        famMan.Set(p, v)
        return True
    except:
        return False

famDocs = IN[0]
famNames = [f.Title for f in famDocs]

famTypesList , famTypeNamesList , famParamsList , famParamNamesList = [],[],[],[]

for f in famDocs :
    famMan = f.FamilyManager
    famTypes = list(famMan.Types)
    typeNames = [f.Name for f in famTypes]
    famTypesList.append(famTypes)
    famTypeNamesList.append(typeNames)
    params = famMan.GetParameters()
    paramNames = [p.Definition.Name for p in params]
    famParamsList.append(params)
    famParamNamesList.append(paramNames)

outcomesList = []
paramNames   = IN[1]
excelData    = IN[2]

# for row in excelData:
#     outcomes = []
#     famInd = famNames.index(row[0])
#     famDoc = famDocs[famInd]
#     typeInd = famTypeNamesList[famInd].index(row[1])
#     famType = famTypesList[famInd][typeInd]
#     values = row[2:]
#     t = Transaction(famDoc, i[0] + " - " + i[1])
#     t.Start()
#     for p,v in zip(paramNames, values):
#         parInd = famParamNamesList[famInd].index(p)
#         famPar = famParamsList[famInd][parInd]
#         setParam = famDoc_setParameter(famDoc, famPar, famType, v)
#         outcomes.append(setParam)
#         t.Commit()
#     outcomesList.append(outcomes)

#     outcomesList.append(row[0])

OUT = famTypeNamesList , famParamNamesList
