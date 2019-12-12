using ShapeFactory


# concentric arcs with random angles
# rad   :start radius
# delta :self explanatory
# num   :number of circles
#
#
function random_concentric(rad::Real, delta::Real, num::Integer, idx::Integer = 1)
    
    app = dispatch("CATIA.Application")
    prt = app.ActiveDocument.Part
    fac = prt.HybridShapeFactory

    rad_dim = prt.Parameters.CreateDimension("min_radius_idx$(idx)", "Length", 0)
    rad_dim.ValuateFromString("$(rad)mm")

    rdt_dim = prt.Parameters.CreateDimension("rad_delta_idx$(idx)", "Length", 0)
    rdt_dim.ValuateFromString("$(delta)mm")

    bdy = prt.HybridBodies.Add()
    bdy.Name = "concentric_circle_idx$(idx)"

    cen = fac.AddNewPointCoord(0, 0, 0)
    pln = prt.FindObjectByName("xy plane")

    rds = rad:delta:(rad + delta * num-1)
    a_a = rand(0:300, num)
    a_b = rand.((:).(a_a .+ 10, 360))

    con = fac.AddNewCircleCtrRadWithAngles.(cen, pln, false, rds, a_a, a_b)

    sub = prt.Relations.SubList(bdy, false)
    exp = x -> "min_radius_idx$(idx) + rad_delta_idx$(idx) * $(x)"
    sub.CreateFormula.("", "", getproperty.(con, :Radius), exp.(1:num))
    
    bdy.AppendHybridShape.(con)
    prt.Update()
end