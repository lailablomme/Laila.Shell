Imports System.Runtime.InteropServices

Namespace Interop.DragDrop
    <ComImport(), Guid("00000121-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IDropSource
        <PreserveSig>
        Function QueryContinueDrag(<[In], MarshalAs(UnmanagedType.Bool)> fEscapePressed As Boolean, <[In]> grfKeyState As Integer) As Integer

        <PreserveSig>
        Function GiveFeedback(<[In]> dwEffect As DROPEFFECT) As Integer
    End Interface
End Namespace
