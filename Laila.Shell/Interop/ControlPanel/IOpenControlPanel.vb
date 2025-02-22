Imports System.Runtime.InteropServices
Imports System.Text

Namespace Interop.ControlPanel
    <ComImport>
    <Guid("D11AD862-66DE-4DF4-BF6C-1F5621996AF1")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IOpenControlPanel
        ''' <summary>
        ''' Retrieves the current view of the Control Panel.
        ''' </summary>
        ''' <param name="pView">An output pointer to the current view.</param>
        ''' <returns>HRESULT indicating success or failure.</returns>
        <PreserveSig>
        Function GetCurrentView(<Out> ByRef pView As CPVIEW) As Integer

        ''' <summary>
        ''' Retrieves the path of a Control Panel item by its canonical name.
        ''' </summary>
        ''' <param name="pszName">
        ''' The canonical name of the Control Panel item.
        ''' </param>
        ''' <param name="pszPath">
        ''' A buffer that receives the path.
        ''' </param>
        ''' <param name="cchPath">The size of the pszPath buffer, in characters.</param>
        ''' <returns>HRESULT indicating success or failure.</returns>
        <PreserveSig>
        Function GetPath(
        <[In], MarshalAs(UnmanagedType.LPWStr)> pszName As String,
        <[In], [Out], MarshalAs(UnmanagedType.LPWStr)> pszPath As StringBuilder,
        <[In]> cchPath As UInteger) As Integer

        ''' <summary>
        ''' Opens a Control Panel item.
        ''' </summary>
        ''' <param name="pszName">The canonical name of the Control Panel item to open.</param>
        ''' <param name="pszPage">The specific page within the item to open, if any.</param>
        ''' <param name="punkSite">An optional site pointer for contextual navigation.</param>
        ''' <returns>HRESULT indicating success or failure.</returns>
        <PreserveSig>
        Function Open(
        <[In], MarshalAs(UnmanagedType.LPWStr)> pszName As String,
        <[In], MarshalAs(UnmanagedType.LPWStr)> pszPage As String,
        <[In], MarshalAs(UnmanagedType.IUnknown)> punkSite As Object) As Integer
    End Interface

    ''' <summary>
    ''' Enumeration for Control Panel views.
    ''' </summary>
    Public Enum CPVIEW
        Classic = 0
        Category = 1
    End Enum
End Namespace
