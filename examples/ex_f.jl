#   ribbon modulation example
#

using ShapeFactory

app = dispatch("CATIA.Application")
sel = app.ActiveDocument.Selection
prt = app.ActiveDocument.Part

a0 = prt.Parameters.CreateDimension("from_angle", "ANGLE", 0)
a0.ValuateFromString("45deg")

a1 = prt.Parameters.CreateDimension("to_angle", "ANGLE", 0)
a1.ValuateFromString("1deg")

rw = prt.Parameters.CreateDimension("ribbon_width", "LENGTH", 0)
rw.ValuateFromString("300mm")

rh = prt.Parameters.CreateDimension("ribbon_height", "LENGTH", 0)
rh.ValuateFromString("1000mm")

ri = prt.Parameters.CreateDimension("ribbon_inset", "LENGTH", 0)
ri.ValuateFromString("50mm")

fi = prt.Parameters.CreateDimension("flange_inset", "LENGTH", 0)
fi.ValuateFromString("20mm")

rn = 10


gst = prt.HybridBodies.Add()
gst.Name = "Output"

fac = prt.HybridShapeFactory
xyp = prt.FindObjectByName("xy plane")

ydr = fac.AddNewDirectionByCoord(0, 1, 0)
zdr = fac.AddNewDirectionByCoord(0, 0, 1)


function Base.repr(idsp::Ptr{IDispatch})
    return join(split(idsp.Name, "\\")[2:end], "\\")
end


function relate(idsp::Ptr{IDispatch}, frm::String)
    name = replace(repr(idsp), "\\" => "_")
    prt.Relations.CreateFormula(name, "", idsp, frm)
end


function blend_surface(edges, flip)
    n = (1, 2)
    e = fac.AddNewExtrude.(edges, 0, (2 * -flip, 2 * flip), zdr)
    b = fac.AddNewBlend()
    b.Coupling = 4
    b.SetCurve.(n, edges)
    b.SetSupport.(n, e)
    b.SetContinuity.(n, 1)
    b.SetTransition.(n, (flip, -flip))
    b.RuledDevelopableSurface = true
    return b
end


function flange_surface(edges, plane, limit1, limit2, flip)
    pin = fac.AddNewPointOnCurveFromDistance(edges[2], 0, false)
    ppn = fac.AddNewLineAngle(edges[2], xyp, pin, false, 0, 1, 90, false)
    relate(pin.Ratio, "$(repr(fi))")
    ext = fac.AddNewExtrude(pin, 0, 1, zdr)
    relate(ext.EndOffset, "$(repr(rh)) * -$(flip)")
    fln = fac.AddNewExtrude(ext, 0, 10, fac.AddNewDirection(ppn))
    fln.SecondLimitType = 2
    fln.SecondUptoElement = plane
    s1 = fac.AddNewHybridSplit(fln, limit1, flip)
    s2 = fac.AddNewHybridSplit(s1, limit2, flip)
    gst.AppendHybridShape(s2)
    sel.Add.((fln, s1, pin, ppn, ext, edges...))
end


function main_edges(pts, angle)
    edg = fac.AddNewLineAngle.(ydr, xyp, pts[end-1:end], false, 0, 1, 1, (false, true))
    relate(edg[1].EndOffset, "$(repr(rw)) + $(repr(ri))")
    relate(edg[1].Angle, angle)
    relate(edg[2].BeginOffset, "-$(repr(ri))")
    relate(edg[2].EndOffset, "$(repr(rw))")
    relate(edg[2].Angle, angle)
    return edg
end


function ribbons()

    px = [fac.AddNewPointCoord(0, 0, 0)]
    dt = 1 / (rn-1)
    pl = nothing
    bl = nothing

    for n in 0:rn-1
        ndt = n * dt 
        sgn = 1 - ((n % 2) * 2)
        ang = "$(repr(a1)) * $(ndt) + $(repr(a0)) * (1- $(ndt))"
        frm = "$(repr(rw))/cos($(ang))"
        frm =  n == 0 ? frm : "$(repr(px[end].Y)) + ($(frm))"

        push!(px, fac.AddNewPointCoord(0, 0, 0))
        relate(px[end].Z, "$((n+1) % 2) * -$(repr(rh))")
        relate(px[end].Y, frm)
        gst.AppendHybridShape(px[end])

        edges = main_edges(px[end-1:end], ang)
        gst.AppendHybridShape.(edges)

        blend = blend_surface(edges, sgn)
        gst.AppendHybridShape(blend)

        if pl != nothing
            flange_surface(edges, pl, bl, blend, sgn)
        end

        pl = fac.AddNewPlaneAngle(xyp, edges[1], 90, false)
        bl = blend
    end
end

ribbons()
sel.VisProperties.SetShow(1)
prt.Update()