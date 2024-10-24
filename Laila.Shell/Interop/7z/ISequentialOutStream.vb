Imports System.Runtime.InteropServices

Namespace SevenZip
    <ComImport>
    <Guid("23170F69-40C1-278A-0000-000300020000")> ' ISequentialOutStream GUID
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface ISequentialOutStream
        <PreserveSig>
        Function Write(data As IntPtr, size As UInt32, processedSize As IntPtr) As Integer
    End Interface
End Namespace