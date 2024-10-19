Imports System
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000D-0000-0000-C000-000000000046")>
Public Interface IEnumSTATSTG
    <PreserveSig>
    Function [Next](
        celt As UInteger,
        <Out> rgelt() As STATSTG,
        <Out> ByRef pceltFetched As UInteger) As Integer

    <PreserveSig>
    Function Skip(
        celt As UInteger) As Integer

    Sub Reset()

    Sub Clone(
        <Out> ByRef ppenum As IEnumSTATSTG)
End Interface
