Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure SHFILEOPSTRUCT
    Public hwnd As IntPtr
    Public wFunc As Integer
    <MarshalAs(UnmanagedType.LPWStr)>
    Public pFrom As String
    <MarshalAs(UnmanagedType.LPWStr)>
    Public pTo As String
    Public fFlags As Short
    Public fAnyOperationsAborted As Boolean
    Public hNameMappings As IntPtr
    Public lProgressDialog As IntPtr
End Structure