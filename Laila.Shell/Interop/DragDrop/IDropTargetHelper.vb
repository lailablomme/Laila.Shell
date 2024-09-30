Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComImport>
<Guid("4657278B-411B-11D2-839A-00C04FD918D0")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IDropTargetHelper
    <PreserveSig>
    Sub DragEnter(<[In]> hwndTarget As IntPtr, <[In], MarshalAs(UnmanagedType.Interface)> pDataObject As Object, <[In]> ByRef ppt As WIN32POINT, <[In]> dwEffect As Integer)
    <PreserveSig>
    Sub DragLeave()
    <PreserveSig>
    Sub DragOver(<[In]> ByRef ppt As WIN32POINT, <[In]> dwEffect As Integer)
    <PreserveSig>
    Sub Drop(<[In], MarshalAs(UnmanagedType.Interface)> pDataObject As Object, <[In]> ByRef ppt As WIN32POINT, <[In]> dwEffect As Integer)
    <PreserveSig>
    Sub Show(<[In]> fShow As Boolean)
End Interface