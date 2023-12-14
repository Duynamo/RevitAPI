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
