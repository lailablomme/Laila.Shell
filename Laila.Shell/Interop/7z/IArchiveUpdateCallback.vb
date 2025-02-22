Imports Laila.Shell.Interop.Properties
Imports Laila.Shell.SevenZip
Imports System.Runtime.InteropServices

Namespace Interop.SevenZip
    <ComImport>
    <Guid("23170F69-40C1-278A-0000-000600800000")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IArchiveUpdateCallback
        <PreserveSig>
        Function SetTotal(bytes As ULong) As Integer
        <PreserveSig>
        Function SetCompleted(bytes As IntPtr) As Integer
        <PreserveSig>
        Function GetUpdateItemInfo(index As UInteger,
                                <[In], Out> ByRef newdata As Integer,
                                <[In], Out> ByRef newprop As Integer,
                                <[In], Out> ByRef indexInArchive As UInteger) As Integer
        <PreserveSig>
        Function GetProperty(index As UInteger,
                         pid As ITEMPROPID,
                         <[In], Out> ByRef value As PROPVARIANT) As Integer
        <PreserveSig>
        Function GetStream(<MarshalAs(UnmanagedType.U4)> index As UInteger,
                       <Out> ByRef stream As ISequentialInStream) As Integer
        <PreserveSig>
        Function SetOperationResult(result As Integer) As Integer
        <PreserveSig>
        Function EnumProperties(enumerator As IntPtr) As Long
    End Interface
End Namespace