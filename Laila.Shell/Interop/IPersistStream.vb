Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

<ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000010c-0000-0000-C000-000000000046")>
Public Interface IPersist
    Sub GetClassID(ByRef pClassID As Guid)
End Interface

<ComImport, Guid("00000109-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IPersistStream
    Inherits IPersist
    Function IsDirty() As Integer
    Function Load(<[In], MarshalAs(UnmanagedType.Interface)> pstm As IStream) As Integer
    Function Save(<[In], MarshalAs(UnmanagedType.Interface)> pstm As IStream, fClearDirty As Boolean) As Integer
    Function GetSizeMax(ByRef pcbSize As Long) As Integer
End Interface