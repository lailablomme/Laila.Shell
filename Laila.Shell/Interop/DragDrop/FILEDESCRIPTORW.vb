Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.Windows

Namespace Interop.DragDrop
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure FILEDESCRIPTORW
        Public dwFlags As UInteger
        Public clsid As Guid
        Public sizel As WIN32SIZE
        Public pointl As WIN32POINT
        Public dwFileAttributes As UInteger
        Public ftCreationTime As Long
        Public ftLastAccessTime As Long
        Public ftLastWriteTime As Long
        Public nFileSizeHigh As UInteger
        Public nFileSizeLow As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public cFileName As String
    End Structure
End Namespace
