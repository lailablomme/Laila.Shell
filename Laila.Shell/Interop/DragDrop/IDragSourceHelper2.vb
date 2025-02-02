Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComImport(), Guid("83E07D0D-0C5F-4163-BF1A-60B274051E40"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IDragSourceHelper2
    <PreserveSig()>
    Function InitializeFromBitmap(ByRef pshdi As SHDRAGIMAGE, ByVal pDataObject As ComTypes.IDataObject) As Integer

    <PreserveSig()>
    Function InitializeFromWindow(ByVal hwnd As IntPtr, ByRef ppt As System.Windows.Point, ByVal pDataObject As ComTypes.IDataObject) As Integer
    <PreserveSig()>
    Sub SetFlags(ByVal dwFlags As Integer)
End Interface