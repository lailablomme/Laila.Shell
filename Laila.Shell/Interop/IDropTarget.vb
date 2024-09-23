Imports System.Runtime.InteropServices

<ComImport(), Guid("00000122-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IDropTarget
    <PreserveSig>
    Function DragEnter(<[In]> pDataObj As IntPtr, grfKeyState As UInteger, pt As System.Windows.Point, ByRef pdwEffect As UInteger) As Integer

    <PreserveSig>
    Function DragOver(grfKeyState As UInteger, pt As System.Windows.Point, ByRef pdwEffect As UInteger) As Integer

    <PreserveSig>
    Function DragLeave() As Integer

    <PreserveSig>
    Function Drop(<[In]> pDataObj As IntPtr, grfKeyState As UInteger, pt As System.Windows.Point, ByRef pdwEffect As UInteger) As Integer
End Interface