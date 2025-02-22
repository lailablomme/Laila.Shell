Imports System.Runtime.InteropServices

Namespace Interop.Items
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure LogFontW
        Public lfHeight As Integer
        Public lfWidth As Integer
        Public lfEscapement As Integer
        Public lfOrientation As Integer
        Public lfWeight As Integer
        Public lfItalic As Byte
        Public lfUnderline As Byte
        Public lfStrikeOut As Byte
        Public lfCharSet As Byte
        Public lfOutPrecision As Byte
        Public lfClipPrecision As Byte
        Public lfQuality As Byte
        Public lfPitchAndFamily As Byte
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
        Public lfFaceName As String
    End Structure
End Namespace
