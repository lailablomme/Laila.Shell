Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure DROPFILES
    Public pFiles As UInteger ' Offset of file list
    Public pt As System.Windows.Point  ' Drop point (client coordinates)
    Public fNC As Boolean ' Is it on NonClient area and pt is in screen coordinates
    Public fWide As Boolean ' Wide character flag
End Structure