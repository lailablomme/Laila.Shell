Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure DROPFILES
    Public pFiles As UInteger ' Offset to file list
    Public pt As WIN32POINT
    Public fNC As Boolean ' True if mouse coords are in screen coordinates
    Public fWide As Boolean ' True if file list is Unicode (WCHAR)
End Structure