for pipe in collector:
    # Define tag type (e.g., Pipe tag type)
    tag_type = TagElementType.Pipe
    
    # Create a reference for the pipe
    pipe_reference = Reference(pipe)
    
    # Create a new tag
    tag = IndependentTag.Create(doc, active_view.Id, pipe_reference, True, tag_type, TagOrientation.Horizontal, XYZ.BasisX)
    
	tag_mode = TagMode.TM_ADDBY_CATEGORY
    tag = Tag.Create(doc, active_view.Id, pipe_reference, True, tag_type, tag_mode)
    # Add the tag to the list
    
	pipeLocation = (pipe.Location as LocationCurve).Curve.Evaluate(0.5, True)

    # Set the location of the tag
    tagLocation = XYZ(pipeLocation.X, pipeLocation.Y, pipeLocation.Z + offset)  # Adjust the offset as needed
    tag.TagHeadPosition = tagLocation 
    pipe_tags.append(tag)


