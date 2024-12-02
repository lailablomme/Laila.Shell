Imports System.Runtime.InteropServices

<ComImport>
<Guid("79EAC9E4-BAF9-11CE-8C82-00AA004BA90B")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IEnumUnknown
    Sub [Next](ByVal celt As Integer, <MarshalAs(UnmanagedType.Interface)> ByRef rgelt As Object, ByRef pceltFetched As Integer)
    Sub Skip(ByVal celt As Integer)
    Sub Reset()
    Sub Clone(<MarshalAs(UnmanagedType.Interface)> ByRef ppenum As IEnumUnknown)
End Interface