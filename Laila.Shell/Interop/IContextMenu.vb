Imports System.Runtime.InteropServices
Imports System.Text

<ComImport(), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214E4-0000-0000-C000-000000000046")>
Public Interface IContextMenu
    ' Define necessary methods here
    Function QueryContextMenu(ByVal hMenu As IntPtr, ByVal indexMenu As UInteger, ByVal idCmdFirst As UInteger, ByVal idCmdLast As UInteger, ByVal uFlags As UInteger) As Integer
    Function InvokeCommand(ByRef pici As CMINVOKECOMMANDINFO) As Integer
    Function GetCommandString(ByVal idCmd As UInteger, ByVal uType As UInteger, ByVal pReserved As IntPtr, ByVal pszName As StringBuilder, ByVal cchMax As UInteger) As Integer
End Interface