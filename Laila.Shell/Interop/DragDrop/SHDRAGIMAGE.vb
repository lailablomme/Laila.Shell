Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure SHDRAGIMAGE
    Public sizeDragImage As WIN32SIZE
    Public ptOffset As WIN32POINT
    Public hbmpDragImage As IntPtr
    Public crColorKey As Integer
End Structure

<StructLayout(LayoutKind.Sequential)>
Public Structure WIN32SIZE
    Public Width As Integer
    Public Height As Integer
End Structure

<StructLayout(LayoutKind.Sequential)>
Public Structure WIN32POINT
    Public x As Integer
    Public y As Integer
End Structure
