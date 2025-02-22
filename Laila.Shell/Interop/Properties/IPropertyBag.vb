Imports System.Runtime.InteropServices

Namespace Interop.Properties
    <ComImport>
    <Guid("55272A00-42CB-11CE-8135-00AA004BB851")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IPropertyBag
        <PreserveSig>
        Function Read(
        <MarshalAs(UnmanagedType.LPWStr)> ByVal pszPropName As String,
        <Out> ByRef pVar As Object,
        <[In]> ByVal pErrorLog As IErrorLog
    ) As Integer

        <PreserveSig>
        Function Write(
        <MarshalAs(UnmanagedType.LPWStr)> ByVal pszPropName As String,
        <[In]> ByRef pVar As PROPVARIANT
    ) As Integer
    End Interface
End Namespace
