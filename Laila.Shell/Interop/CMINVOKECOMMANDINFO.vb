﻿Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure CMINVOKECOMMANDINFO
    Public cbSize As UInteger
    Public fMask As UInteger
    Public hwnd As IntPtr
    <MarshalAs(UnmanagedType.LPStr)>
    Public lpVerb As String
    <MarshalAs(UnmanagedType.LPStr)>
    Public lpParameters As String
    <MarshalAs(UnmanagedType.LPStr)>
    Public lpDirectory As String
    Public nShow As Integer
    Public dwHotKey As UInteger
    Public hIcon As IntPtr
End Structure