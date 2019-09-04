
using ShapeFactory


cat = dispatch("CATIA.Application")
prt = cat.ActiveDocument.Part


# running rectangle example
#
function ex_a(num::Integer, rad::Real)    
    
    pln = prt.FindObjectByName("xy plane")
    gst = prt.FindObjectByName("Geometrical Set.1")

    sk1 = gst.HybridSketches.Add(pln)
    f2d = sk1.OpenEdition()
    con = sk1.Constraints

    
    ang0 = prt.Parameters.CreateDimension("master_angle", "ANGLE", 0)
    ang0.ValuateFromString("7deg")

    pt0 = f2d.CreatePoint(0, 0)
    con.AddMonoEltCst.(0, pt0) # fixed

    ax0 = f2d.CreateLine(0, 0, 0, rad * 3)
    ax0.StartPoint = pt0
    ax0.Construction = true
    con.AddMonoEltCst(13, ax0) # vertical

    cor = [-1 1; 1 1; 1 -1; -1 -1] .* rad
    edg = []

    for n in 1:num
       
        pts = f2d.CreatePoint.(cor[:,1], cor[:,2])
        if n == 1
            con.AddMonoEltCst.(0, pts) # fixed
        else
            con.AddBiEltCst.(2, pts, edg) # on line
        end

        run = [pts circshift(pts, -1)]
        edg = f2d.CreateLine.(1:4, 0, 0, 0)
        
        setproperty!.(edg, :StartPoint, run[:, 1])
        setproperty!.(edg, :EndPoint, run[:, 2])
        
        con.AddBiEltCst.(11, edg[2:end], edg[1:end-1]) # perp

        if n != 1
            ax = f2d.CreateLine(0, 0, 0, 100)
            ax.StartPoint = pt0
            ax.Construction = true
            ang = con.AddBiEltCst(6, ax, ax0) # angle
            ang.AngleSector = 0
            ax0 = ax
            prt.Relations.CreateFormula("", "", ang.Dimension, "master_angle")
            con.AddBiEltCst(11, edg[1], ax) # perpendicular
        end

    end

    sk1.CloseEdition()
    prt.Update()

end