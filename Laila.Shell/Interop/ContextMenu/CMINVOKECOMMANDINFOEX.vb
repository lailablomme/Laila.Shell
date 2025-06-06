﻿Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.Windows

Namespace Interop.ContextMenu
    <StructLayout(LayoutKind.Sequential)>
    Public Structure CMInvokeCommandInfoEx
        Public cbSize As Integer
        Public fMask As Integer
        Public hwnd As IntPtr
        Public lpVerb As IntPtr
        <MarshalAs(UnmanagedType.LPStr)> Public lpParameters As String
        <MarshalAs(UnmanagedType.LPStr)> Public lpDirectory As String
        Public nShow As Integer
        Public dwHotKey As Integer
        Public hIcon As IntPtr
        <MarshalAs(UnmanagedType.LPStr)> Public lpTitle As String
        Public lpVerbW As IntPtr
        <MarshalAs(UnmanagedType.LPWStr)> Public lpParametersW As String
        <MarshalAs(UnmanagedType.LPWStr)> Public lpDirectoryW As String
        <MarshalAs(UnmanagedType.LPWStr)> Public lpTitleW As String
        Public ptInvoke As WIN32POINT
    End Structure
End Namespace
