Imports System.Runtime.InteropServices

Namespace Interop.Items
    Public Structure SHITEMID
        ''' <summary>
        ''' The size of identifier, in bytes, including cb itself.
        ''' </summary>
        Public cb As UShort
        ''' <summary>
        ''' A variable-length item identifier.
        ''' </summary>
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1)>
        Public abID() As Byte
    End Structure
End Namespace

