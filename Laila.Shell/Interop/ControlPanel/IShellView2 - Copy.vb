
Imports System.Runtime.InteropServices
Imports System.Windows

<ComImport, Guid("88E39E80-3578-11CF-AE69-08002B2E1262")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IShellView2
    <PreserveSig>
    Function GetWindow(ByRef hwnd As IntPtr) As Integer

    <PreserveSig>
    Function ContextSensitiveHelp(fEnterMode As Integer) As Integer

    ' Method declarations
    <PreserveSig>
    Function TranslateAccelerator(<[In]> ByRef pmsg As System.Windows.Forms.Message) As Integer

    <PreserveSig>
    Function EnableModeless(<[In]> fEnable As Boolean) As Integer

    <PreserveSig>
    Function UIActivate(<[In]> uState As UInteger) As Integer

    <PreserveSig>
    Function Refresh() As Integer

    <PreserveSig>
    Function CreateViewWindow(
        <[In]> <MarshalAs(UnmanagedType.Interface)> psvPrevious As IShellView,
        <[In]> ByRef lpfs As FOLDERSETTINGS,
        <[In]> <MarshalAs(UnmanagedType.Interface)> psb As IShellBrowser,
        <[In], Out> ByRef prcView As WIN32RECT,
        <Out> ByRef phWnd As IntPtr) As Integer

    <PreserveSig>
    Function DestroyViewWindow() As Integer

    <PreserveSig>
    Function GetCurrentInfo(<Out> ByRef lpfs As FOLDERSETTINGS) As Integer

    <PreserveSig>
    Function AddPropertySheetPages(
        <[In]> dwReserved As UInteger,
        <[In]> pfn As IntPtr,
        <[In]> lparam As IntPtr) As Integer

    <PreserveSig>
    Function SaveViewState() As Integer

    <PreserveSig>
    Function SelectItem(
        <[In]> pidlItem As IntPtr,
        <[In]> uFlags As Int32) As Integer

    <PreserveSig>
    Function GetItemObject(
        <[In]> uItem As UInteger,
        <[In]> ByRef riid As Guid,
        <Out> ByRef ppv As IntPtr) As Integer

    ' Corrected methods for IShellView2

    <PreserveSig>
    Function CreateViewWindow2(ByRef lpParams As SV2CVW2_PARAMS) As Integer

    <PreserveSig>
    Function GetView(ByRef pvid As Guid, uView As Int32) As Integer

    <PreserveSig>
    Function HandleRename(pidlNew As IntPtr) As Integer

    <PreserveSig>
    Function SelectAndPositionItem(pidlItem As IntPtr, uFlags As UInt32, ByRef ppt As WIN32POINT) As Integer

End Interface

' Define necessary types

'<StructLayout(LayoutKind.Sequential)>
'Public Structure SHELLVIEWID
'    Public dwID As UInt32
'End Structure