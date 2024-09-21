Imports System.Runtime.InteropServices

<ComImport>
<Guid("70629033-e363-4a28-a567-0db78006e6d7")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IEnumShellItems
    Sub [Next](<[In]> celt As UInteger, <Out, MarshalAs(UnmanagedType.LPArray)> rgelt() As IShellItem, <Out> ByRef pceltFetched As UInteger)
    Sub Skip(<[In]> celt As UInteger)
    Sub Reset()
    Sub Clone(<Out, MarshalAs(UnmanagedType.Interface)> ByRef ppenum As IEnumShellItems)
End Interface