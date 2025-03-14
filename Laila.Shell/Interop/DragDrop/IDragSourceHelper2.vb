Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports Laila.Shell.Interop.Windows

Namespace Interop.DragDrop
    <ComImport(), Guid("83E07D0D-0C5F-4163-BF1A-60B274051E40"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDragSourceHelper2
        <PreserveSig()>
        Function InitializeFromBitmap(ByRef pshdi As SHDRAGIMAGE, ByVal pDataObject As IDataObject_PreserveSig) As Integer

        <PreserveSig()>
        Function InitializeFromWindow(ByVal hwnd As IntPtr, ByRef ppt As WIN32POINT, ByVal pDataObject As IDataObject_PreserveSig) As Integer
        <PreserveSig()>
        Function SetFlags(ByVal dwFlags As Integer) As Integer
    End Interface
End Namespace
