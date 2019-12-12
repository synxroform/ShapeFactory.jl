# apollonian gasket constructed in sketch environment
# please dont` use large subdivision levels if you cant` afford supercomputer 
#

using ShapeFactory

app = dispatch("CATIA.Application")
prt = app.ActiveDocument.Part

xy = prt.FindObjectByName("xy plane")

gst = prt.HybridBodies.Add()
gst.Name = "Output"

sk1 = gst.HybridSketches.Add(xy)
f2d = sk1.OpenEdition()
con = sk1.Constraints


function base_circle(rad)
    pt0 = f2d.CreatePoint(0, 0)
    con.AddMonoEltCst.(0, pt0)

    ax0 = f2d.CreateLine(0, 0, 0, rad * 3)
    ax0.StartPoint = pt0
    ax0.Construction = true
    con.AddMonoEltCst(13, ax0)

    c0 = f2d.CreateClosedCircle(0, 0, rad)
    c0.CenterPoint = pt0
    rad = con.AddMonoEltCst.(14, c0)
    rad.Dimension.ValuateFromString("$(rad)mm")

    return c0, ax0
end


function base_triad(cc0, ax0)
    
    cx = f2d.CreateClosedCircle(0, 0, cc0.Radius - 1)
    cx.CenterPoint = cc0.CenterPoint
    cx.Construction = true

    tri = map((2pi/3):(2pi/3):2pi) do ang
        pt = f2d.CreatePoint(sin(ang) * cx.Radius, cos(ang)*cx.Radius)
        pt.Construction = true
        cc = f2d.CreateClosedCircle(sin(ang), cos(ang), 1)
        cc.CenterPoint = pt
        rad = con.AddMonoEltCst.(14, cc)
        rad.Dimension.ValuateFromString("$(cc0.Radius / 2)mm")
        cc => rad
    end
    
    tri, rad = first.(tri), last.(tri)

    con.AddBiEltCst.(2, getproperty.(tri, :CenterPoint), cx) 

    push!(tri, tri[1])
    for n in 1:3
        con.AddBiEltCst(4, tri[n], tri[n+1])
    end
    
    con.AddBiEltCst(2, tri[1].CenterPoint, ax0)
    map(x -> x.Deactivate(), rad)
    con.AddBiEltCst.(4, tri[1:end-1], cc0)

    return tri[1:end-1]
end


function circle_tri(a, b, c)
    loop = (a, b, c, a)
    cntr = [0, 0, 0]
    for n in 1:3
        pt    = f2d.CreatePoint(0, 0)
        pc    = con.AddBiEltCst.(2, pt, (loop[n], loop[n+1]))
        cntr += pt.GetCoordinates([0, 0, 0])[2] ./ 3
        map(x -> x.Deactivate(), pc)
    end
    cir = f2d.CreateClosedCircle(cntr[1], cntr[2], 0.01)
    con.AddBiEltCst.(4, cir, sort([a, b, c], by=x -> x.Radius, rev=true))
    return cir
end


function subdivide(a, b, c, stop)
    loop = (a, b, c, a)
    x = circle_tri(a, b, c)
    if stop > 1
        for n in 1:3
            subdivide(loop[n], loop[n+1], x, stop-1)
        end
    end
end

c0, ax = base_circle(10)
a, b, c = base_triad(c0, ax)
subdivide(a, b, c, 2)

loop = (a, b, c, a)
for n in 1:3
    subdivide(loop[n], loop[n+1], c0, 3)
end

sk1.CloseEdition()
prt.Update()