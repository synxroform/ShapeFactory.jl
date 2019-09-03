
# VARIANT GETTERS_SETTERS


function com_variantset(var::Ref{Variant}, data::T) where {T}
    unsafe_store!(Ptr{T}(pointer_from_objref(var) + 8), data)
    return var
end

function com_variantset(pvar::Ptr{Variant}, data::T) where {T}
    unsafe_store!(Ptr{T}(pvar + 8), data)
end

function com_variantget(var::Ref{Variant}, ::Type{T}) where {T}
    return unsafe_load(Ptr{T}(pointer_from_objref(var) + 8))
end

function com_variantget(pvar::Ptr{Variant}, ::Type{T}) where {T}
    return unsafe_load(Ptr{T}(pvar + 8))
end



# VARIANT CONSTRUCTORS


function variant_free(var::Variant)
    if !(gettype(var.vt) <: Ptr)
        ccall((:VariantClear, "oleaut32.dll"), Cvoid, (Ref{Variant},), var)
    end
end


# This is not a copy constructor, just plumbing for for some algorithms.
# Variant structure is immutable therefore it can be copied just as v1 = v2.
#
Variant(pass::Variant) = pass


# Convert julia string to Variant with newly allocated BSTR.
#
function Variant(value::String) 
    str = push!(transcode(UInt16, value), '\0')
    var = Variant(ccall((:SysAllocString, "oleaut32.dll"), Ptr{UInt16}, (Ptr{UInt16},), str))
    return var
end


# Allocate SafeArray of variants to be used wtih COM methods.
#
# vars : Array of variants.
# return: Variant with allocated SafeArray.
#
function com_createarray(vars::Array{Variant}) ::Variant
     
    abnds = SafeArrayBound.(collect(size(vars)), 0)
    array = ccall((:SafeArrayCreate, "oleaut32.dll"), Ptr{SafeArray}, (Cushort, Cuint, Ptr{SafeArrayBound}), 
                VT_VARIANT, ndims(vars), pointer(abnds))
    if array != C_NULL
        pdata = Ref(Ptr{Variant}(0))
        ccall((:SafeArrayAccessData, "oleaut32.dll"), Cint, (Ptr{SafeArray}, Ref{Ptr{Variant}}), array, pdata)
        for n in 0:length(vars) - 1
            unsafe_store!(pdata[] + sizeof(Variant)*n, vars[n+1])
        end
        ccall((:SafeArrayUnaccessData, "oleaut32.dll"), Cint, (Ptr{SafeArray},), array)
    end
    return com_variantset(Ref(Variant(VT_VARIANT | VT_ARRAY, 0, 0, 0, 0)), array)[]
end



function Variant(value::Array{T}) where {T}
    return com_createarray(Variant.(value))
end


function Variant(value::T) where {T <: Union{Ptr, Number}}
    var = Ref(Variant(getindex(T), 0, 0, 0, 0))
    com_variantset(var, value)
    return var[]
end


# CONVERTERS

# Return the typed value of Variant, with special handling of arrays. 
#
function value(var::Variant)
    if (var.vt & VT_ARRAY) > 0
        ret = convert(Array, var)
        if ret[1] === nothing
            popfirst!(ret)
            if length(ret) == 1
                return ret[1]
            end
        end
        return ret
    else
        return convert(get(indextype, var.vt & VT_TYPEMASK, Nothing), var)
    end
end


function Base.convert(::Type{Nothing}, var::Variant)
    return nothing
end


function Base.convert(::Type{T}, var::Variant) where {T <: Union{Ptr, Number}}
    value = com_variantget(Ref(var), T)
    if var.vt == 8
        return convert(String, value)
    else
        return value
    end
end


Base.String(var::Variant) = convert(String, var)
#
function Base.convert(::Type{String}, var::Variant)
    return String(convert(Ptr{UInt16}, var))
end


Base.String(str::Ptr{UInt16}) = convert(String, str)
#
function Base.convert(::Type{String}, str::Ptr{UInt16})
    buf = Vector{UInt16}()
    for n in 1:typemax(Int16)
        chr = unsafe_load(str, n)
        if chr == zero(UInt16)
            break
        end
        push!(buf, chr)
    end
    return transcode(String, buf)
end


function Base.convert(::Type{Array}, var::Variant)
    if (var.vt & VT_ARRAY) > 0
        if (var.vt & VT_BYREF) > 0
            apntr = unsafe_load(com_variantget(Ref(var), Ptr{Ptr{SafeArray}}))
        else
            apntr = com_variantget(Ref(var), Ptr{SafeArray})
        end
        array = unsafe_load(apntr)
        abnds = getfield.(unsafe_load.(Ptr{SafeArrayBound}(apntr + 24), 1:array.cDims), :cElements)
        if length(abnds) > 1
            abnds[1], abnds[2] = abnds[2], abnds[1]
        end
        vars = unsafe_load.(Ptr{Variant}(array.pvData), 1:prod(abnds))
        return reshape(convert.(gettype.(getfield.(vars, :vt)), vars), Tuple(abnds))
    end
    return nothing
end