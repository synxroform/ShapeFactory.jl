# hydra-growing example
#

using ShapeFactory


compass = CartesianIndex.(((1, 0), (0, -1), (-1, 0), (0, 1)))


function scale_idx(idx, n)
    return idx + (idx - one(idx)) * (n-1)
end


function stencil(grid, index, hv)
    for di in CartesianIndices((0:1, 0:1))
        grid[index + di] = (grid[index + di] + hv) % 4
    end
end


function hydra_grid(height, width, maxgen=4)
    sub_grid = zeros(Int8, height * 2, width * 2)
    top_grid = Set(CartesianIndices((height, width)))
    while length(top_grid) > 0
        group = Dict{CartesianIndex, BitArray}()
        group[pop!(top_grid)] = [0, 0, 0, 0]

        gen = 0
        while length(group) > 0
            indexa, edgesa = rand(group)
            side = first(rand(filter(x -> last(x) == 0, collect(enumerate(edgesa)))))
            indexb = indexa + compass[side]
            edgesa[side] = 1
            if all(edgesa)
                pop!(group, indexa)
            end
            if indexb in top_grid
                di = compass[side]
                stencil(sub_grid, scale_idx(indexa, 2) + di, 2 + (side-1)%2)
                group[indexb] = [0, 0, 0, 0]
                delete!(top_grid, indexb)
                if (gen += 1) > maxgen
                    break
                end
            end
        end
    end
    return sub_grid
end


grid = hydra_grid(10, 10)

app = dispatch("CATIA.Application")
prt = app.ActiveDocument.Part


gst = prt.HybridBodies.Add()
gst.Name = "Output"

fac = prt.HybridShapeFactory
xyp = prt.FindObjectByName("xy plane")

pts = fac.AddNewPointCoord.((1000, 0, 1000, 0), (1000, 1000, 0, 0), 0) 
ccs = fac.AddNewCircleCtrRadWithAngles.(pts, xyp, false, 500, 
                                        (180, 270, 90, 0), (270, 360, 180, 90))

pts = fac.AddNewPointCoord.((500, 0, 500, 1000), (0, 500, 1000, 500), 0)
lns = fac.AddNewLinePtPt.(pts[1:2], pts[3:4])

gst.AppendHybridShape.(ccs)
gst.AppendHybridShape.(lns)


for y in 0:size(grid, 1)-1
    for x in 0:size(grid, 2)-1
        mx, my = x * 1000, y * 1000
        xv = fac.AddNewDirectionByCoord(mx, my, 1)
        xd = sqrt(mx^2 + my^2)
        st = grid[y + 1, x + 1]
        if st > 1
            tx = fac.AddNewTranslate(lns[st-1], xv, xd)
        else
            tx = fac.AddNewTranslate(ccs[((x+st)%2 + ((y+st)%2)*2) + 1], xv, xd)
        end
        tx.VectorType = 0
        gst.AppendHybridShape(tx)
    end
end

prt.Update()