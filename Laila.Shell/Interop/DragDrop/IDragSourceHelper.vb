Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("DE5BF786-477A-11D2-839D-00C04FD918D0")>
Public Interface IDragSourceHelper
    <PreserveSig()>
    Function InitializeFromBitmap(ByRef pshdi As SHDRAGIMAGE, ByVal pDataObject As IDataObject) As Integer

    <PreserveSig()>
    Function InitializeFromWindow(ByVal hwnd As IntPtr, ByRef ppt As System.Windows.Point, ByVal pDataObject As IDataObject) As Integer
End Interface