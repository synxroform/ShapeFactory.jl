module ShapeFactory


export dispatch, dispaxxx, IDispatch

include("ShapeDispatch.jl")
include("ShapeVariant.jl")
include("ShapeSimplex.jl")

# ITypeInfo machinery.

function com_getfuncdesc(type_info::Ptr{ITypeInfo}, idx::Cuint)
    desc = Ref(Ptr{FUNCDESC}(0))
    func = unsafe_load(unsafe_load(Ptr{Ptr{Ptr{Cvoid}}}(type_info)), 6)
    ccall(func, Cint, (Ptr{ITypeInfo}, Cuint, Ref{Ptr{FUNCDESC}}), type_info, idx, desc)
    return desc[]
end


function com_getfuncname(type_info::Ptr{ITypeInfo}, fd::FUNCDESC)
    name = Ref(Ptr{UInt16}(0))
    func = unsafe_load(unsafe_load(Ptr{Ptr{Ptr{Cvoid}}}(type_info)), 13)
    ccall(func, Cint, (Ptr{ITypeInfo}, Clong, Ref{Ptr{UInt16}}, Ptr{Nothing}, Ptr{Nothing}, Ptr{Nothing}), 
            type_info, fd.memid, name, C_NULL, C_NULL, C_NULL)
    return Symbol(String(name[]))
end


function com_gettypeinfo(idsp::Ptr{IDispatch})
    info = Ref(Ptr{ITypeInfo}(0))
    func = unsafe_load(unsafe_load(Ptr{Ptr{Ptr{Cvoid}}}(idsp)), 5)
    ccall(func, Cint, (Ptr{IDispatch}, Cuint, Cint, Ref{Ptr{ITypeInfo}}), idsp, 0, LCID, info)
    return info[]
end


function com_gettypeattr(type_info::Ptr{ITypeInfo})
    attr = Ref(Ptr{TYPEATTR}(0))
    func = unsafe_load(unsafe_load(Ptr{Ptr{Ptr{Cvoid}}}(type_info)), 4)
    ccall(func, Cint, (Ptr{ITypeInfo}, Ref{Ptr{TYPEATTR}}), type_info, attr)
    data = unsafe_load(attr[])
    com_releasetypeattr(type_info, attr[])
    return data
end


function com_releasetypeattr(type_info::Ptr{ITypeInfo}, attr::Ptr{TYPEATTR})
    func = unsafe_load(unsafe_load(Ptr{Ptr{Ptr{Cvoid}}}(type_info)), 20)
    ccall(func, Cint, (Ptr{ITypeInfo}, Ptr{TYPEATTR}), type_info, attr)
end


# In order to call something like cat.ActiveDocument.Part.Update(), i need to distiguish between
# properties and methods, to do so i need some information about symbol, but i cant'get it
# exclusively by symbol name. FUNCDESC retrieved by index, conversion if memid to index handled by
# ITypeInfo2 interface which seems broken in CATIA COM Implementation. 
# Therefore there is some dancing with tambourine.
#
function com_gettypedata(obj::Ptr{IDispatch})
    info = com_gettypeinfo(obj)
    attr = com_gettypeattr(info)

    get!(type_cache, hash(attr.guid)) do
        tdata = TypeData(info, Dict{Symbol, FuncData}())
        for n in 0:attr.cFuncs - 1
            fd = unsafe_load(com_getfuncdesc(info, Cuint(n)))
            name = com_getfuncname(info, fd)
            meth = get!(tdata.func, name) do 
                fdata = FuncData(fd.memid, 0, Array{UInt16}(undef, fd.cParams))
                if fd.cParams > 0
                    for p in 0:fd.cParams - 1
                        fdata.param_flags[p + 1] = unsafe_load(Ptr{UInt16}(fd.lprgelemdescParam + 32 * p + 24))
                    end
                end
                fdata
            end
            meth.invoke_mask |= fd.invkind
        end
        tdata
    end
end


# IDispatch machinery


function com_invoke(obj::ComObject, name::Symbol, flags::UInt16, args::Ref{DISPPARAMS}, ret::Ptr{Variant})
    fptr = unsafe_load(unsafe_load(Ptr{Ptr{Ptr{Cvoid}}}(obj.ptr)), 7)
    ccall(fptr, Cint, (Ptr{IDispatch}, Clong, Ref{GUID}, UInt32, UInt16, Ref{DISPPARAMS}, Ptr{Variant}, Ptr{Cvoid}, Ptr{Cvoid}), 
    obj.ptr, obj.dat.func[name].memid, IID_NULL, LCID, flags, args, ret, C_NULL, C_NULL)
end


function com_callmethod(obj::ComObject, name::Symbol, flag::UInt16 = 0x0001)
    var = Ref(Variant(0, 0, 0, 0, 0))
    arg = Ref(DISPPARAMS(C_NULL, C_NULL, 0, 0))
    com_invoke(obj, name, flag, arg, Ptr{Variant}(pointer_from_objref(var)))
    ret = value(var[])
    variant_free(var[])
    return ret
end


# this function can be simpler, but it must handle [out] arguments.
#
function com_callfunc(obj::ComObject, name::Symbol, args::Array{Variant})
    flags = obj.dat.func[name].param_flags
    packs = deepcopy(args)
    out_n = 0
    for n in 1:length(args)
        if (flags[n] & PARAMFLAG_OUT) > 0
            com_variantset(pointer(packs, n), Int128(0))
            if (packs[n].vt & VT_ARRAY) > 0
                unsafe_store!(Ptr{Cshort}(pointer(packs, n)), VT_BYREF | VT_ARRAY | VT_VARIANT)
                com_variantset(pointer(packs, n), Ptr{Ptr{SafeArray}}(pointer(args, n) + 8))
            else
                unsafe_store!(Ptr{Cshort}(pointer(packs, n)), VT_BYREF | VT_ARRAY | VT_VARIANT)
                com_variantset(pointer(packs, n), pointer(args, n))
            end
            out_n += 1
        end
    end
    var = Ref(Variant(0, 0, 0, 0, 0))
    arg = Ref(DISPPARAMS(pointer(reverse!(packs)), C_NULL, length(args), 0))
    com_invoke(obj, name, 0x0001, arg, Ptr{Variant}(pointer_from_objref(var)))
    
    if out_n > 0
        ret = value.(reverse!(push!(filter(v -> (v.vt & VT_BYREF) > 0 , packs), var[])))
    else
        ret = value(var[])
    end
    
    map(x -> variant_free.(x), [args, packs, [var[]]])
    return ret
end


function comcall(obj::ComObject, name::Symbol, args)
    if length(args) == 0
        return com_callmethod(obj, name)
    else
        return com_callfunc(obj, name, collect(Variant.(args)))
    end
end


function Base.getproperty(idsp::Ptr{IDispatch}, name::Symbol)
    dat = com_gettypedata(idsp)
    if !(name in keys(dat.func))
        error("SF: there is no member $(name) on COM object.")
    end
    obj = ComObject(idsp, dat)
    if (dat.func[name].invoke_mask & 0x1) > 0
        return (args...) -> comcall(obj, name, args)
    else
        return com_callmethod(obj, name, 0x0002)
    end
end


# special treatment of arrays
#
function Base.getproperty(idspa::Array{Ptr{IDispatch}}, name::Symbol)
    dat = com_gettypedata(idspa[1])
    if !(name in keys(dat.func))
        error("SF: there is no member $(name) on COM object.")
    end
    objs = ComObject.(idspa, Ref(dat))
    if (dat.func[name].invoke_mask & 0x1) > 0
        return (args...) -> comcall.(objs, name, Ref(args))
    else
        return com_callmethod.(objs, name, 0x0002)
    end
end


function Base.setproperty!(idsp::Ptr{IDispatch}, name::Symbol, value)
    dat = com_gettypedata(idsp)
    if !(name in keys(dat.func))
        error("SF: there is no member $(name) on COM object.")
    end
    obj = ComObject(idsp, dat)
    var = Ref(Variant(value))
    pid = [Int32(-3)]
    arg = Ref(DISPPARAMS(Ptr{Variant}(pointer_from_objref(var)), pointer(pid), 1, 1))
    com_invoke(obj, name, 0x0004, arg, Ptr{Variant}(C_NULL))
    variant_free(var[])
end


function dispatch(name::String)

    wname = push!(transcode(UInt16, name), '\0')
    ccall((:CoInitializeEx, "ole32.dll"), Cint, (Ptr{Cvoid}, Cuint), C_NULL, 0)
    
    guid = Ref(GUID(0, 0, 0, (0, 0, 0, 0, 0, 0, 0, 0)))
    ccall((:CLSIDFromProgID, "ole32.dll"), Cint, (Ptr{UInt16}, Ref{GUID}), wname, guid)

    pdisp = Ref(Ptr{IDispatch}(0))
    err = ccall((:CoCreateInstance, "ole32.dll"), Cint, (Ref{GUID}, Ptr{IUnknown}, Cuint, GUID, Ref{Ptr{IDispatch}}), 
            guid, C_NULL, CLSCTX_SERVER, IID_IDispatch, pdisp)

    if (err != 0)
        error("Failed to create IDispatch")
    end

    return pdisp[]
end


dispaxxx(name::String) = Ptr{IDispaxxx}(dispatch(name))

end