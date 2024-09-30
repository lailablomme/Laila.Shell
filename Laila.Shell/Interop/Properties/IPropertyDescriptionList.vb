Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

<ComImport, Guid("1F9FC1D0-C39B-4B26-817F-011967D3440E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IPropertyDescriptionList

    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetCount(ByRef pcElem As UInt32)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetAt(<[In]> iElem As UInt32, <[In]> ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IPropertyDescription2)
End Interface