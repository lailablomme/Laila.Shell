Imports System.Runtime.InteropServices

Namespace Interop.Items
    <ComImport>
    <Guid("1C9CD5BB-98E9-4491-A60F-31AACC72B83C")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IObjectWithSelection
        <PreserveSig>
        Function SetSelection(<[In], MarshalAs(UnmanagedType.Interface)> psia As IShellItemArray) As Integer

        <PreserveSig>
        Function GetSelection(<Out> <MarshalAs(UnmanagedType.Interface)> ByRef ppsia As IShellItemArray) As Integer
    End Interface
End Namespace
