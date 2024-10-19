Imports System.Runtime.InteropServices

<ComImport>
<Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IObjectArray
    Sub GetCount(ByRef cObjects As UInteger)
    Sub GetAt(iIndex As UInteger, <MarshalAs(UnmanagedType.LPStruct)> ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppvObject As Object)
End Interface