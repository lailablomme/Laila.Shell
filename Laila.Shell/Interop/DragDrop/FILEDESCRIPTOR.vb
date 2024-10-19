Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
Public Structure FILEDESCRIPTOR
    Public dwFlags As UInteger
    Public ftCreationTime As FILETIME
    Public ftLastAccessTime As FILETIME
    Public ftLastWriteTime As FILETIME
    Public nFileSizeHigh As UInteger
    Public nFileSizeLow As UInteger
    Public dwFileAttributes As UInteger
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
    Public cFileName As String
End Structure