﻿Imports System.Runtime.InteropServices
Imports System.Text

Namespace Interop.ContextMenu
    <ComImport(),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
     GuidAttribute("bcfce0a0-ec17-11d0-8d10-00a0c90f2719")>
    Public Interface IContextMenu3
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

        'IContextMenu3 method
        <PreserveSig()>
        Function HandleMenuMsg2(
            ByVal uMsg As Integer,
            ByVal wParam As IntPtr,
            ByVal lParam As IntPtr,
            ByVal plResult As IntPtr
        ) As Integer
    End Interface
End Namespace
