Imports System.Runtime.InteropServices
Imports System.Text

Namespace Interop.ContextMenu
    <ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), GuidAttribute("000214e4-0000-0000-c000-000000000046")>
    Public Interface IContextMenu
        <PreserveSig()>
        Function QueryContextMenu(ByVal hmenu As IntPtr,
            ByVal iMenu As Integer,
            ByVal idCmdFirst As Integer,
            ByVal idCmdLast As Integer,
            ByVal uFlags As Integer) As Integer
        <PreserveSig()>
        Function InvokeCommand(ByRef pici As CMInvokeCommandInfoEx) As Integer

        <PreserveSig()>
        Function GetCommandString(ByVal idcmd As Integer,
            ByVal uflags As Integer,
            ByVal reserved As Integer,
            ByVal commandstring As StringBuilder,
            ByVal cch As Integer) As Integer
    End Interface
End Namespace
