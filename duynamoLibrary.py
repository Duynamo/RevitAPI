#def 01
#to Flatten nested lists to expect Level
def flattenTo1dList(arr):
    result = []
    def recursive_flatten(subarray):
        for item in subarray:
            if is_arrayitem(item):
                recursive_flatten(item)
            else:
                result.append(item) 
    recursive_flatten(arr)        
    return result

def isArray(array):
    return "List" in obj.__class__.__name__

def getArrayRank(array):
    if isArray(array)
        return 1 + max(getArrayRank(item) for item in array)
    else: 
        return 0

def flattenArrayByOptionals(objects, level = 1):
    result = []
    for item in objects:
        newRank = getArrayRank(item)
        if isArray(item) and newRank >= level :
            arr = flattenTo1dList(item)
            result.append(arr)
        else:
            result.append(item)
    return result

##selectionFilter
class selectionFilter(ISelectionFilter):
    def __init__(self, category1, category2,category3, category4):
        self.category1 = category1
        self.category2 = category2
        self.category3 = category3
        self.category4 = category4
    
    def AllowElement(self, element):
        if element.Category.Name == self.category1 or element.Category.Name == self.category2 or element.Category.Name == self.category3 or element.Category.Name == self.category4:
            return True
        else:
            return False
    def AllowReference(reference, point):
        return False
    
ele = selectionFilter('Structural Columns', 'Structural Framing', 'Pipes','Pipe Fittings')
ele.copy()

##pickObjects
def pickObjects():
    from Autodesk.Revit.UI.Selection import ObjectType
    refs = uidoc.Selection.PickObjects(ObjectType.Element)
    return  [doc.GetElement(i.ElementId) for i in refs], [ref for ref in refs]



def pickObjects():
    from Autodesk.Revit.UI.Selection import ObjectType
    refs = uidoc.Selection.PickObjects(ObjectType.Element)
    return  [doc.GetElement(i.ElementId) for i in refs]

lstPipes = pickObjects()
flat_lstPipes = [pipe for pipe in lstPipes]

pipe1 = flat_lstPipes[0]
pipe2 = flat_lstPipes[1]

# Tạo một biến lưu khoảng cách gần nhất
closest_distance = float('inf')  # Gán giá trị lớn nhất cho khoảng cách ban đầu
closest_connector_pipe2 = None
closest_connector_pipe1 = None

# Lặp qua connectors của ống thứ nhất
for connector1 in pipe1.ConnectorManager.Connectors:
    for connector2 in pipe2.ConnectorManager.Connectors:
        distance = connector1.Origin.DistanceTo(connector2.Origin)
        if distance < closest_distance:
            closest_distance = distance
            closest_connector_pipe1 = connector1
            closest_connector_pipe2 = connector2
            
if closest_connector_pipe1 and closest_connector_pipe2:
	cor_ConnectPipe1 = closest_connector_pipe1.Origin.ToPoint()
	cor_ConnectPipe2 = closest_connector_pipe2.Origin.ToPoint()

TransactionManager.Instance.EnsureInTransaction(doc)
pipeJoin = doc.Create.NewPipe(closest_connector_pipe1, closest_connector_pipe2)
TransactionManager.Instance.TransactionTaskDone()


OUT = cor_ConnectPipe1, cor_ConnectPipe2