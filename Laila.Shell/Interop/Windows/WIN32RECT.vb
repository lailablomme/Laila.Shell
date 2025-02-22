Imports System.Runtime.InteropServices

Namespace Interop.Windows
    <StructLayout(LayoutKind.Sequential)>
    Public Structure WIN32RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure
End Namespace

