Imports Laila.Shell.SevenZip
Imports System.Runtime.InteropServices

<ComImport>
<Guid("23170F69-40C1-278A-0000-000600800000")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IArchiveUpdateCallback

    ' Notifies the total bytes of target files.
    <PreserveSig>
    Function SetTotal(bytes As ULong) As Integer

    ' Notifies the bytes to be archived.
    <PreserveSig>
    Function SetCompleted(bytes As IntPtr) As Integer

    ' Gets information of updating item.
    <PreserveSig>
    Function GetUpdateItemInfo(index As UInteger,
                                <[In], Out> ByRef newdata As Integer,
                                <[In], Out> ByRef newprop As Integer,
                                <[In], Out> ByRef indexInArchive As UInteger) As Integer

    ' Gets the property information according to the specified arguments.
    <PreserveSig>
    Function GetProperty(index As UInteger,
                         pid As PROPID,
                         <[In], Out> ByRef value As PROPVARIANT) As Integer

    ' Gets the stream according to the specified arguments.
    <PreserveSig>
    Function GetStream(<MarshalAs(UnmanagedType.U4)> index As UInteger,
                       <Out> ByRef stream As ISequentialInStream) As Integer

    ' Sets the specified operation result.
    <PreserveSig>
    Function SetOperationResult(result As Integer) As Integer

    ' EnumProperties 7-zip internal function.
    <PreserveSig>
    Function EnumProperties(enumerator As IntPtr) As Long

End Interface