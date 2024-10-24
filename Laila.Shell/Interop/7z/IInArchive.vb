Imports System.Runtime.InteropServices

Namespace SevenZip
    <ComImport>
    <Guid("23170F69-40C1-278A-0000-000600600000")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IInArchive
        <PreserveSig>
        Function Open(<[In], MarshalAs(UnmanagedType.Interface)> ByVal stream As IInStream,
                  ByVal max As IntPtr, ' ref ULong max
                  <[In], MarshalAs(UnmanagedType.Interface)> ByVal callback As IArchiveOpenCallback) As Integer
        Sub Close()
        Function GetNumberOfItems() As UInteger
        Sub GetProperty(ByVal index As UInteger, ByVal pid As ITEMPROPID, ByRef value As PROPVARIANT)
        <PreserveSig>
        Function Extract(<[In], MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=1)> ByVal indexes As UInteger(),
                     ByVal count As UInteger,
                     ByVal test As Integer,
                     <[In], MarshalAs(UnmanagedType.Interface)> ByVal callback As IArchiveExtractCallback) As Integer
        Sub GetArchiveProperty(ByVal aid As ARCHIVEPROPID, ByRef value As PROPVARIANT)
        Function GetNumberOfProperties() As UInteger
        Sub GetPropertyInfo(ByVal index As UInteger,
                        <Out, MarshalAs(UnmanagedType.BStr)> ByRef name As String,
                        ByRef pid As ITEMPROPID,
                        ByRef type As UShort)
        Function GetNumberOfArchiveProperties() As UInteger
        Sub GetArchivePropertyInfo(ByVal index As UInteger,
                               <Out, MarshalAs(UnmanagedType.BStr)> ByRef name As String,
                               ByRef pid As ITEMPROPID,
                               ByRef type As UShort)
    End Interface
End Namespace