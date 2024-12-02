Imports System.Runtime.InteropServices
Imports System.Windows

<ComImport>
<Guid("8895b1c6-b41f-4c1c-a562-0d564250836f")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IPreviewHandler
    <PreserveSig>
    Function SetWindow(hwnd As IntPtr, <[In]> ByRef prc As WIN32RECT) As Integer

    <PreserveSig>
    Function SetRect(<[In]> ByRef prc As WIN32RECT) As Integer

    <PreserveSig>
    Function DoPreview() As Integer

    <PreserveSig>
    Function Unload() As Integer

    <PreserveSig>
    Function SetFocus() As Integer

    <PreserveSig>
    Function QueryFocus(<Out> ByRef phwnd As IntPtr) As Integer
End Interface