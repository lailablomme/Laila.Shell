Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.Windows

Namespace Interop.Items
    <ComImport>
    <Guid("BDDACB60-7657-47AE-8445-D23E1ACF82AE")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IExplorerCommandState
        <PreserveSig>
        Function GetState(
            <[In], MarshalAs(UnmanagedType.Interface)> psia As IShellItemArray,
            <[In]> grfKeyState As MK,
            <Out> ByRef pCmdState As EXPCMDSTATE
        ) As Integer
    End Interface
End Namespace