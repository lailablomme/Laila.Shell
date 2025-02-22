Imports System.Runtime.InteropServices

Namespace Interop.ContextMenu
    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("3409E930-5A39-11D1-83FA-00A0C90DC849")>
    Public Interface IContextMenuCB
        <PreserveSig()>
        Function CallBack(
            ByVal psf As IntPtr,
            ByVal hwndOwner As IntPtr,
            ByVal pdtobj As IntPtr,
            ByVal uMsg As UInteger,
            ByVal wParam As UInteger,
            ByVal lParam As UInteger
        ) As Integer
    End Interface
End Namespace
