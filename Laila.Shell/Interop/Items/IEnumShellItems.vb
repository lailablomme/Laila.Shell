Imports System.Runtime.InteropServices

Namespace Interop.Items
    <ComImport>
    <Guid("70629033-e363-4a28-a567-0db78006e6d7")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IEnumShellItems
        <PreserveSig>
        Function [Next](<[In]> celt As UInteger, <Out, MarshalAs(UnmanagedType.LPArray)> rgelt() As IShellItem, <Out> ByRef pceltFetched As UInteger) As Integer
        <PreserveSig>
        Function Skip(<[In]> celt As UInteger) As Integer
        <PreserveSig>
        Function Reset() As Integer
        <PreserveSig>
        Function Clone(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppenum As IEnumShellItems) As Integer
    End Interface
End Namespace
