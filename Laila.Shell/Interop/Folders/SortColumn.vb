﻿Imports System.Runtime.InteropServices

Namespace Interop.Folders
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure SortColumn
        <MarshalAs(UnmanagedType.LPWStr)>
        Public propertyName As String
        Public direction As Integer ' Sort direction: 0 = ascending, 1 = descending
    End Structure
End Namespace

