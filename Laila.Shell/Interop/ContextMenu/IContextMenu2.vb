Imports System.Runtime.InteropServices
Imports System.Text

<ComImport(),
InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
GuidAttribute("000214f4-0000-0000-c000-000000000046")>
Public Interface IContextMenu2
    ' IContextMenu methods

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

    'IContextMenu2 method
    <PreserveSig()>
    Function HandleMenuMsg(
    ByVal uMsg As Integer,
    ByVal wParam As IntPtr,
    ByVal lParam As IntPtr
    ) As Integer

End Interface