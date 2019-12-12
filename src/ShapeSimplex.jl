# This machinery included for servers without ITypeInfo support such as Rhino
# just call dispaxxx("Rhino.Interface") to handle naked IDispatch
# use this functionality on your own risk


function com_invoke(idspx::Ptr{IDispaxxx}, name::Symbol, flags::UInt16, args::Ref{DISPPARAMS}, ret::Ptr{Variant})

    fmid, fptr = unsafe_load.(unsafe_load(Ptr{Ptr{Ptr{Cvoid}}}(idspx)), (6, 7))
    wname = [push!(transcode(UInt16, String(name)), '\0')]
    memid = Ref(Clong(0))

    # idispatch getidsofnames
    ccall(fmid, Cint, (Ptr{IDispaxxx}, Ref{GUID}, Ref{Ptr{UInt16}}, Cuint, UInt32, Ref{Clong}),
        idspx, IID_NULL, wname, 1, LCID,  memid)    

    # idispatch invoke
    ccall(fptr, Cint, (Ptr{IDispaxxx}, Clong, Ref{GUID}, UInt32, UInt16, Ref{DISPPARAMS}, Ptr{Variant}, Ptr{Cvoid}, Ptr{Cvoid}), 
        idspx, memid[], IID_NULL, LCID, flags, args, ret, C_NULL, C_NULL)
end


function com_callfunc(idspx::Ptr{IDispaxxx}, name::Symbol, args::Array{Variant})
    var = Ref(Variant(0, 0, 0, 0, 0))
    arg = Ref(DISPPARAMS(pointer(reverse!(args)), C_NULL, length(args), 0))
    com_invoke(idspx, name, 0x0001, arg, Ptr{Variant}(pointer_from_objref(var)))
    
    ret = value(var[])
    variant_free.(args)
    variant_free(var[])
    return ret
end



function com_callmethod(idspx::Ptr{IDispaxxx}, name::Symbol, flag::UInt16 = 0x0001)
    var = Ref(Variant(0, 0, 0, 0, 0))
    arg = Ref(DISPPARAMS(C_NULL, C_NULL, 0, 0))
    
    com_invoke(idspx, name, flag, arg, Ptr{Variant}(pointer_from_objref(var)))
    
    ret = value(var[])
    variant_free(var[])
    if isa(ret, Ptr{IDispatch})
        return Ptr{IDispaxxx}(ret)
    else
        return ret
    end
end


function Base.getproperty(idspx::Ptr{IDispaxxx}, name::Symbol)
    return (args...) -> 
        if length(args) > 0
            com_callfunc(idspx, name, collect(Variant.(args)))
        else
            com_callmethod(idspx, name)
        end
end