Imports System.Runtime.InteropServices

Namespace Interop.Folders
    <ComImport>
    <Guid("a879e3c4-af77-44fb-8f37-ebd1487cf920")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IQueryParserManager
        Sub CreateLoadedParser(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszCatalog As String, ByVal langidForKeywords As Integer, ByVal riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppQueryParser As Object)
        Sub InitializeOptions(ByVal dwOptions As Integer)

        Sub SetOption(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszOptionName As String, ByVal pszOptionValue As String)
    End Interface
End Namespace


