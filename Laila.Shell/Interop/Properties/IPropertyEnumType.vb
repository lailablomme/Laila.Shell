Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

<ComImport, Guid("11E1FBF9-2D56-4A6B-8DB3-7CD193A471F2"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IPropertyEnumType
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetEnumType(<[Out]> ByRef penumtype As PropEnumType)

    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetValue(<[Out]> ppropvar As PROPVARIANT)

    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetRangeMinValue(<[Out]> ppropvar As PROPVARIANT)

    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetRangeSetValue(<[Out]> ppropvar As PROPVARIANT)

    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetDisplayText(<Out, MarshalAs(UnmanagedType.LPWStr)> ByRef ppszDisplay As String)
End Interface
