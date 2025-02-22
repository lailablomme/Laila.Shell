Imports Laila.Shell.SevenZip
Imports System.Runtime.InteropServices

Namespace Interop.SevenZip
    <ComImport>
    <Guid("23170F69-40C1-278A-0000-000600200000")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IArchiveExtractCallback
        <PreserveSig>
        Function SetTotal(ByVal bytes As ULong) As Integer
        <PreserveSig>
        Function SetCompleted(ByVal bytes As IntPtr) As Integer
        <PreserveSig>
        Function GetStream(ByVal index As UInteger,
                       <MarshalAs(UnmanagedType.Interface)> ByRef stream As ISequentialOutStream,
                       ByVal mode As ProcessingMode) As Integer
        <PreserveSig>
        Function PrepareOperation(ByVal mode As ProcessingMode) As Integer
        <PreserveSig>
        Function SetOperationResult(ByVal code As Integer) As Integer
    End Interface
End Namespace