Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure WIN32_FILE_ATTRIBUTE_DATA
    Public dwFileAttributes As FILE_ATTRIBUTE
    Public ftCreationTime As Long
    Public ftLastAccessTime As Long
    Public ftLastWriteTime As Long
    Public nFileSizeLow As UInteger
    Public nFileSizeHigh As UInteger
End Structure
