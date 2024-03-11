def pickFace():
	elements = []
	planarFace = []
	refs = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "pick face")
	for r in refs:
		e = uidoc.Document.GetElement(r)
		dsface = e.GetGeometryObjectFromReference(r).ToProtoType(True)
		[f.Tags.AddTag("RevitFaceReference", r) for f in dsface]
		dsfaces.append(dsface)
		curveLoop = dsface.GetEdgesAsCurveLoops()
		elements.append(e)
		planarFace.append(dsface)
	return  refs, elements, face ,curveLoop