Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure EXCEPINFO
    Public wCode As UShort
    Public wReserved As UShort
    <MarshalAs(UnmanagedType.BStr)>
    Public bstrSource As String
    <MarshalAs(UnmanagedType.BStr)>
    Public bstrDescription As String
    <MarshalAs(UnmanagedType.BStr)>
    Public bstrHelpFile As String
    Public dwHelpContext As Integer
    Public pvReserved As IntPtr
    Public pfnDeferredFillIn As IntPtr
    Public scode As Integer
End Structure