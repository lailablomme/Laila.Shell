Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.Windows

Namespace Interop.DragDrop
    <StructLayout(LayoutKind.Sequential)>
    Public Structure SHDRAGIMAGE
        Public sizeDragImage As WIN32SIZE
        Public ptOffset As WIN32POINT
        Public hbmpDragImage As IntPtr
        Public crColorKey As Integer
    End Structure
End Namespace

