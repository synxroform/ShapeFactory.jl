using ShapeFactory

# Rhino API also available.


rhino = dispaxxx("Rhino.Interface.6")

so = rhino.GetScriptObject()
so.AddLine.([0, 0, 0], [10, 10, 10])

