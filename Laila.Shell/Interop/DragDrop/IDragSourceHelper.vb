Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports Laila.Shell.Interop.Windows

Namespace Interop.DragDrop
    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("DE5BF786-477A-11D2-839D-00C04FD918D0")>
    Public Interface IDragSourceHelper
        <PreserveSig()>
        Function InitializeFromBitmap(ByRef pshdi As SHDRAGIMAGE, ByVal pDataObject As ComTypes.IDataObject) As Integer

        <PreserveSig()>
        Function InitializeFromWindow(ByVal hwnd As IntPtr, ByRef ppt As WIN32POINT, ByVal pDataObject As ComTypes.IDataObject) As Integer
    End Interface
End Namespace
