Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.Windows

Namespace Interop.Items
    <ComImport>
    <Guid("7F9185B0-CB92-43C5-80A9-92277A4F7B54")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IExecuteCommand
        <PreserveSig>
        Function SetKeyState(<[In]> grfKeyState As UInteger) As Integer

        <PreserveSig>
        Function SetParameters(<[In], MarshalAs(UnmanagedType.LPWStr)> pszParameters As String) As Integer

        <PreserveSig>
        Function SetPosition(<[In]> pt As WIN32POINT) As Integer

        <PreserveSig>
        Function SetShowWindow(<[In]> nShow As Integer) As Integer

        <PreserveSig>
        Function SetNoShowUI(<[In]> fNoShowUI As Boolean) As Integer

        <PreserveSig>
        Function SetDirectory(<[In], MarshalAs(UnmanagedType.LPWStr)> pszDirectory As String) As Integer

        <PreserveSig>
        Function Execute() As Integer
    End Interface
End Namespace