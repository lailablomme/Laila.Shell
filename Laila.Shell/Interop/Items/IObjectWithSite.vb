Imports System.Runtime.InteropServices
Imports Laila.Shell.Controls.ExplorerMenu
Imports Laila.Shell.Interop.Folders

Namespace Interop.Items

    <ComImport>
    <Guid("FC4801A3-2BA9-11CF-A229-00AA003D7352")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IObjectWithSite
        <PreserveSig>
        Function SetSite(<[In], MarshalAs(UnmanagedType.Interface)> punkSite As IServiceProvider) As Integer

        <PreserveSig>
        Function GetSite(<[In]> ByRef riid As Guid, <Out> <MarshalAs(UnmanagedType.IUnknown)> ByRef ppvSite As Object) As Integer
    End Interface


    <ComImport>
    <Guid("6D5140C1-7436-11CE-8034-00AA006009FA")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IServiceProvider
        <PreserveSig>
        Function QueryService(
            <[In]> ByRef guidService As Guid,
            <[In]> ByRef riid As Guid,
            <Out> ByRef ppvObject As IntPtr
        ) As Integer
    End Interface
End Namespace