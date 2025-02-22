Imports System.Runtime.InteropServices

Namespace Interop.Folders
    <StructLayout(LayoutKind.Sequential)>
    Public Structure FOLDERSETTINGS
        Public ViewMode As Int32       ' View mode (FOLDERVIEWMODE values)
        Public fFlags As UInt32        ' View options (FOLDERFLAGS bits)
    End Structure
End Namespace

