Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

<ComImport, Guid("9B6E051C-5DDD-4321-9070-FE2ACB55E794"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IPropertyEnumType2
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
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    <PreserveSig>
    Function GetImageReference(<Out, MarshalAs(UnmanagedType.LPWStr)> ByRef ppszImageRes As String) As HResult
End Interface
