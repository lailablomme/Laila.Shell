Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
Public Structure CMINVOKECOMMANDINFOEX
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
    Public ptInvoke As System.Drawing.Point
End Structure