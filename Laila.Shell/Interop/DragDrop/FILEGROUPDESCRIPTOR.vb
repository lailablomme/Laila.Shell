﻿Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Public Structure FILEGROUPDESCRIPTOR
    Public cItems As UInteger
    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1)>
    Public fd As FILEDESCRIPTOR()  ' Array of FILEDESCRIPTOR
End Structure