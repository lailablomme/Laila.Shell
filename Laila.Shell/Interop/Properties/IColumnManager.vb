Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Namespace Interop.Properties
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("d8ec27bb-3f3b-4042-b10a-4acfd924d453")>
    Public Interface IColumnManager
        <MethodImpl(MethodImplOptions.PreserveSig Or MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Function SetColumnInfo(ByRef propkey As PROPERTYKEY, ByRef pcmci As CM_COLUMNINFO) As Integer

        <MethodImpl(MethodImplOptions.PreserveSig Or MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Function GetColumnInfo(ByRef propKey As PROPERTYKEY, <[In], Out> ByRef pcmci As CM_COLUMNINFO) As Integer

        <MethodImpl(MethodImplOptions.PreserveSig Or MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Function GetColumnCount(dwFlags As CM_ENUM_FLAGS, ByRef puCount As UInt32) As Integer

        <MethodImpl(MethodImplOptions.PreserveSig Or MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Function GetColumns(dwFlags As CM_ENUM_FLAGS, <MarshalAs(UnmanagedType.LPArray), Out> rgkeyOrder As PROPERTYKEY(), cColumns As UInt32) As Integer

        <MethodImpl(MethodImplOptions.PreserveSig Or MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
        Function SetColumns(<MarshalAs(UnmanagedType.LPArray)> rgkeyOrder As PROPERTYKEY(), cVisible As UInt32) As Integer
    End Interface
End Namespace

