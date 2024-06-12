
import clr
import math

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
doc = DocumentManager.Instance.CurrentDBDocument

clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
#The inputs to this node will be stored as a list in the IN variables.
if isinstance(IN[0], list):
	fittings = UnwrapElement(IN[0])
	toggle = 0
else:
	toggle = 1
	fittings = [UnwrapElement(IN[0])]
	
	
def getConnSysType(connector):
	domain = connector.Domain
	if domain == Domain.DomainHvac:
		return connector.DuctSystemType.ToString()
	elif domain == Domain.DomainPiping:
		return connector.PipeSystemType.ToString()
	elif domain == Domain.DomainElectrical:
		return connector.ElectricalSystemType.ToString()
	else:
		return None

ref = []
conns = []
for fitting in fittings:
	refs = []
	connlist = []
	try:
		connectors = fitting.MEPModel.ConnectorManager.Connectors
	except:
		try:
			connectors = fitting.ConnectorManager.Connectors
		except:			
			ref.append(None)
			conns.append(None)
			continue
	for conn in connectors:
		connlist.append(conn)
		description.append(conn.Description)
		for c in conn.AllRefs:
			refs.append(c.Owner)

	ref.append(refs)
	conns.append(connlist)



#Assign your output to the OUT variable.
if toggle:
	OUT = refs,connlist
else:
	OUT = ref, conns