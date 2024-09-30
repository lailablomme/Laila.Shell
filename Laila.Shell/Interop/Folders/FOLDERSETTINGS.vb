Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure FOLDERSETTINGS
    Public ViewMode As UInt32       ' View mode (FOLDERVIEWMODE values)
    Public fFlags As UInt32        ' View options (FOLDERFLAGS bits)
End Structure
