Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace Interop.Properties
    <ComImport, Guid("A99400F4-3D84-4557-94BA-1242FB2CC9A6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IPropertyEnumTypeList
        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub GetCount(<[Out]> ByRef pctypes As UInt32)

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub GetAt(<[In]> itype As UInt32, <[In]> ByRef riid As Guid, <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppv As IPropertyEnumType)

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub GetConditionAt(<[In]> index As UInt32, <[In]> ByRef riid As Guid, ByRef ppv As IntPtr)

        <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Sub FindMatchingIndex(<[In]> propvarCmp As PROPVARIANT, <[Out]> ByRef pnIndex As UInt32)
    End Interface
End Namespace

