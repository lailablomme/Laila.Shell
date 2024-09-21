Imports System.Runtime.InteropServices

<ComImport>
<Guid("3127CA40-446E-11CE-8135-00AA004BB851")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IErrorLog
    <PreserveSig>
    Function AddError(
        <MarshalAs(UnmanagedType.LPWStr)> ByVal pszPropName As String,
        <[In]> ByRef pExcepInfo As EXCEPINFO
    ) As Integer
End Interface