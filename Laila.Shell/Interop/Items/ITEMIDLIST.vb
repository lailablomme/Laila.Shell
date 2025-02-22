Imports System.Runtime.InteropServices

Namespace Interop.Items
    ''' <summary>
    ''' Contains a list of item identifiers.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure ITEMIDLIST
        ''' <summary>
        ''' A list of item identifiers.
        ''' </summary>
        <MarshalAs(UnmanagedType.Struct)> Public mkid As SHITEMID
    End Structure
End Namespace
