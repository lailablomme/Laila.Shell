Imports System
Imports System.Runtime.InteropServices

Namespace Interop.Storage
    <ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000010C-0000-0000-C000-000000000046")>
    Public Interface IPersist
        <PreserveSig>
        Function GetClassID(ByRef pClassID As Guid) As Integer
    End Interface
End Namespace
