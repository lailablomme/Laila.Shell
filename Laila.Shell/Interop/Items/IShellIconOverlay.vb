Imports System.Runtime.InteropServices

Namespace Interop.Items
    <ComImport>
    <Guid("7d688a70-c613-11d0-999b-00c04fd655e1")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IShellIconOverlay
        <PreserveSig>
        Function GetOverlayIndex(<[In]> ByVal pidl As IntPtr, ByRef pIndex As Integer) As Integer

        <PreserveSig>
        Function GetOverlayIconIndex(<[In]> ByVal pidl As IntPtr, ByRef pIconIndex As Integer) As Integer
    End Interface
End Namespace
