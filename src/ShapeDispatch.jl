
# Windows COM Structures constants and types

struct GUID
    d1::Cuint
    d2::Cushort
    d3::Cushort
    d4::NTuple{8, Cuchar}
end

const IID_ITypeInfo  = GUID(0x00020401, 0x0000, 0x0000, (0xc0,0x00, 0x00,0x00,0x00,0x00,0x00,0x46))
const IID_IDispatch  = GUID(0x00020400, 0x0000, 0x0000, (0xc0,0x00,0x00,0x00,0x00,0x00,0x00,0x46))
const IID_NULL       = GUID(0, 0, 0, (0, 0, 0, 0, 0, 0, 0, 0))

const CLSCTX_ALL = 0x1 | 0x2 | 0x4 | 0x10
const CLSCTX_SERVER = 0x1 | 0x4 | 0x10
const LCID = ccall((:GetUserDefaultLCID, "kernel32.dll"), Cint, ())


const PARAMFLAG_OUT  = UInt16(0x2)
const VT_USERDEFINED = UInt16(29)
const VT_VARIANT     = UInt16(12)
const VT_ARRAY       = UInt16(0x2000)
const VT_BYREF       = UInt16(0x4000)
const VT_TYPEMASK    = UInt16(0xfff)


abstract type IUnknown end
abstract type IDispatch end
abstract type IDispaxxx end
abstract type ITypeInfo end


struct Variant
    vt::Cshort
    r1::Cshort
    r2::Cshort
    r3::Cshort
    data::UInt128
end

struct SafeArrayBound
    cElements::Culong
    lLbound::Clong
end

struct SafeArray
    cDims::Cushort
    fFeatures::Cushort
    cbElements::Culong
    cLocks::Culong
    pvData::Ptr{Cvoid}
    bounds::SafeArrayBound
end

struct DISPPARAMS
    rgvarg::Ptr{Variant}
    rgdispidNamedArgs::Ptr{Int32}
    cArgs::Cuint
    cNamedArgs::Cuint
end

struct TYPEDESC
    union::NTuple{14, Cuchar}
    vt::UInt16
end

struct IDLDESC
    dwReserved::Ptr{Cvoid}
    wIDLFlags::UInt16
end

struct ELEMDESC
    tdesc::TYPEDESC
    paramdesc::UInt128 # 10 padded to 16
end

struct TYPEATTR
    guid::GUID
    lcid::UInt32
    dwReserved::UInt32
    memidConstructor::UInt32
    memidDestructor::UInt32
    lpstrSchema::Ptr{Cvoid}
    cbSizeInstance::UInt32
    typekind::UInt32
    cFuncs::UInt32
    cVars::UInt32
    cImplTypes::UInt32
    cbSizeVft::UInt32
    cbAlignment::UInt32
    wTypeFlags::UInt32
    wMajorVerNum::UInt32
    wMinorVerNum::UInt32
    tdescAlias::TYPEDESC
    idldescType::IDLDESC
end

struct EXCEPINFO
    wCode::Cushort
    wReserved::Cushort
    bstrSource::Ptr{UInt16}
    bstrDescription::Ptr{UInt16}
    bstrHelpFile::Ptr{UInt16}
    dwHelpContext::Cushort
    pvReserved::Ptr{Cvoid}
    pfnDeferredFillIn::Ptr{Cvoid}
    scode::Clong
end

struct FUNCDESC
    memid::Clong
    lprgscode::Ptr{Cvoid}
    lprgelemdescParam::Ptr{ELEMDESC}
    funckind::UInt32
    invkind::UInt32
    callconv::UInt32
    cParams::UInt16
    cParamsOpt::UInt16
    oVft::UInt16
    cScodes::UInt16
    elemdescFunc::ELEMDESC
    wFuncFlags::UInt16
end


mutable struct FuncData
    memid::Clong
    invoke_mask::UInt32
    param_flags::Array{UInt16}
end


mutable struct TypeData
    info::Ptr{ITypeInfo}
    func::Dict{Symbol, FuncData}
end


struct ComObject
    ptr::Ptr{IDispatch}
    dat::TypeData
end

# Global Dictionaries

type_cache = Dict{UInt64, TypeData}()


const typeindex = Dict(
    Ptr{Cvoid}      => [26 1 24],
    Ptr{IDispatch}  => [9],
    Ptr{IUnknown}   => [13],
    Ptr{Int32}      => [37],
    Ptr{UInt32}     => [38],
    Ptr{UInt16}     => [8 30 31],
    Variant         => [12],
    Array           => [27],
    Int8            => [16],
    Int16           => [2],
    Int32           => [3 22],
    Int64           => [20],
    UInt8           => [17],
    UInt16          => [18 29],
    UInt32          => [19 23 25],
    UInt64          => [21],
    Float32         => [4],
    Float64         => [5 7],
    GUID            => [72],
    Bool            => [11]
)

function reversedict(dict::Dict)
    rev = Dict()
    for (k, v) in dict
        for n in v
            rev[n] = k
        end
    end
    return rev
end

const indextype = reversedict(typeindex)


function gettype(idx::Cshort)
    if (idx & VT_ARRAY) > 0
        return Array
    else
        return get(indextype, idx, Nothing)
    end
end


function getindex(::Type{T}) where {T}
    if T == Int64
        # Int64 not suppported in CATIA.
        return typeindex[Int32][1] 
    else
        return typeindex[T][1]
    end
end

