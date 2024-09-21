Imports System.Runtime.InteropServices

<ComImport(), Guid("0000010b-0000-0000-C000-000000000046"),
   InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IPersistFile

    <PreserveSig()> Function IsDirty() As Integer
    <PreserveSig()> Function Load(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszFileName As String,
            ByVal dwMode As Integer) As HResult
    <PreserveSig()> Function Save(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszFileName As String,
            <MarshalAs(UnmanagedType.Bool)> ByVal fRemember As Boolean) As HResult
    <PreserveSig()> Function SaveCompleted(<MarshalAs(UnmanagedType.LPWStr)> ByVal pszFileName As String) As HResult
    ' CAUTION: This method is untested! It returns a value, so probably should be StringBuilder, but this
    '      disagrees with the C# definition above.
    <PreserveSig()> Function GetCurFile(<MarshalAs(UnmanagedType.LPWStr)> ByVal ppszFileName As Text.StringBuilder) As HResult
End Interface