Imports System.Runtime.InteropServices

Namespace Interop.Items
    <ComImport>
    <Guid("196bf9a5-b346-4ef0-aa1e-5dcdb76768b1")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IPreviewHandlerVisuals
        <PreserveSig>
        Function SetBackgroundColor(ByVal color As UInteger) As Integer

        <PreserveSig>
        Function SetFont(ByRef plf As LogFontW) As Integer

        <PreserveSig>
        Function SetTextColor(ByVal color As UInteger) As Integer
    End Interface
End Namespace
