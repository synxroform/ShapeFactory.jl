
# .................... cubic b-spline functions ...............................

# n / 0 == 0
#
function zero_divide(a, b)
    x = a / b
    return ifelse(isfinite(x), x, 0)
end


# uniform-knot p:degree n:number_of_points
#
function knot(p, n)
    base = [ifelse(k <= p, 0.0, 1.0) for k in (0:p + n)]
    if n - p == 1
        return base
    else
        for k in 0:n-p
            base[p + k + 1] = k / (n - p)
        end
    end
    return base
end


# b-spline basis p:degree i:number u:parameter k:knot
#
function bbasis(p, i, u, k)
    if p == 0
        return ifelse(k[i] <= u <= k[i+1], 1, 0)
    else
        m0 = zero_divide(u - k[i], k[i+p] - k[i])
        m1 = zero_divide(k[i+p+1] - u, k[i+p+1] - k[i+1])
        return m0 * bbasis(p-1, i, u, k) + m1 * bbasis(p-1, i+1, u, k)
    end
end


# b-spline basis first-derivative
#
function bbasis_fd(p, i, u, k)
    return zero_divide(p, k[i+p] - k[i]) * bbasis(p-1, i, u, k) -
        zero_divide(p, k[i+p+1] - k[i+1]) * bbasis(p-1, i+1, u, k)
end


# b-spline curve, points: as columns, t: 1d vector
#
function bcurve(points, u, bas=bbasis)
    knt = Ref(knot(3, size(points, 2)))
    tmp = zeros(size(points, 1), length(u), size(points, 2))
    for n in axes(points, 2)
        tmp[:,:,n] = points[:,n] .* bas.(3, n, u, knt)'
    end
    return dropdims(sum(tmp, dims=3), dims=3)
end


# b-spline surface, points: size(cp_u, cp_v, 3), uv: size(nu, nv, 2) 
#
function bsurface(points, uv, bu=bbasis, bv=bbasis)
    ku, kv = knot.(3, size(points))     
    uv_tmp = zeros(size(uv, 1), size(uv, 2), size(points, 3))
    
    for I in CartesianIndices(uv[:,:,1])
        for J in CartesianIndices(points[:,:,1])
            uv_tmp[I,:] += points[J,:] .* 
                bu(3, J[1], uv[I, 1], ku) .* 
                bv(3, J[2], uv[I, 2], kv)
        end
    end
    return uv_tmp
end


function mgrid(nu, nv, a=0:1, b=0:1)
    uv = zeros(nu, nv, 2)
    uv[:,:,1] .= first(a) : ((last(a)-first(a)) / (nu-1)) : last(a)
    uv[:,:,2] .= (first(b) : ((last(b)-first(b)) / (nv-1)) : last(b))'
    return uv
end


function rnd_grid(x, y)
    uv = mgrid(x, y, 0:x, 0:y)
    az = rand(Float32, x, y)
    return cat(uv, az, dims = 3)
end