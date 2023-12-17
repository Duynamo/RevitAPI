import clr
from System.Collections.Generic import List
from Autodesk.Revit.DB import BuiltInCategory, FilteredElementCollector, LogicalOrFilter
from Autodesk.Revit.DB import ElementId

# Assuming 'doc' is your Revit Document object
doc = __revit__.ActiveUIDocument.Document

def get_built_in_categories(categories):
    # Retrieve BuiltInCategory objects for the given category names
    built_in_categories = []
    for category in categories:
        built_in_category = BuiltInCategory.Invalid
        try:
            built_in_category = BuiltInCategory[str(category)]
        except:
            pass
        if built_in_category != BuiltInCategory.Invalid:
            built_in_categories.append(built_in_category)
    return built_in_categories

categories_to_filter = ['Walls', 'Doors', 'Windows']  # Replace with your desired categories

# Get BuiltInCategory objects from category names
built_in_categories = get_built_in_categories(categories_to_filter)

# Filter elements for each built-in category
element_collectors = [FilteredElementCollector(doc).OfCategory(cat) for cat in built_in_categories]

# Combine filters using LogicalOrFilter to filter elements by multiple categories
combined_filter = LogicalOrFilter(element_collectors)

# Get elements that belong to any of the specified categories
elements = combined_filter.ToElements()

# Get element IDs for the filtered elements
element_ids = List[ElementId]()
for element in elements:
    element_ids.Add(element.Id)

# Output the filtered element IDs
OUT = element_ids
