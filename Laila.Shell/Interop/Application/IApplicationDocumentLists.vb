Imports System.Runtime.InteropServices

Namespace Interop.Application
    <ComImport>
    <Guid("3C594F9F-9F30-47A1-979A-C9E83D3D0A06")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IApplicationDocumentLists
        Sub SetAppID(<MarshalAs(UnmanagedType.LPWStr)> appID As String)
        Sub GetList(listType As APPDOCLISTTYPE, cItemsDesired As UInteger, <[In]> ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef items As IObjectArray)
    End Interface
End Namespace
